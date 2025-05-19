// src/tools/createSiteContentType.ts
import request from 'request-promise';
import { IToolResult } from '../interfaces';
import { getSharePointHeaders, getRequestDigest } from '../auth_factory';
import { SharePointConfig } from '../config';

export interface CreateSiteContentTypeParams {
    url: string;
    contentTypeData: {
        Name: string;
        Description?: string;
        ParentContentTypeId?: string;
        ContentTypeId?: string; // TODO: REMOVE IT
        Group?: string;
    };
}

/**
 * Create a new content type in a SharePoint site
 * @param params Parameters including site URL and content type data
 * @param config SharePoint configuration
 * @returns Tool result with created content type information
 */
export async function createSiteContentType(
    params: CreateSiteContentTypeParams, 
    config: SharePointConfig
): Promise<IToolResult> {
    const { url, contentTypeData } = params;
    console.error("createSiteContentType tool called with URL:", url);

    try {
        // Authenticate with SharePoint and get request digest
        const headers = await getSharePointHeaders(url, config);
        const requestDigest = await getRequestDigest(url, headers);
        
        // Set up headers for POST request
        const postHeaders = {
            ...headers,
            "X-RequestDigest": requestDigest,
            "content-type": "application/json;odata=verbose",
            "accept": "application/json;odata=verbose"
        };
        
        // Get parent content type ID
        const parentContentTypeId = contentTypeData.ParentContentTypeId || "0x01";
        console.error(`Using parent content type ID: ${parentContentTypeId}`);
        
        try {
            // To create a proper content type inheritance structure, we need to retrieve
            // the parent content type first, then create a child based on it
            
            // First, check if the parent content type is available
            console.error(`Checking if parent content type ${parentContentTypeId} exists`);
            
            try {
                // Try to get the parent content type
                const parentResponse = await request({
                    url: `${url}/_api/web/contenttypes/getbyid('${encodeURIComponent(parentContentTypeId)}')`,
                    headers: headers,
                    method: 'GET',
                    json: true,
                    timeout: 15000
                });
                
                console.error(`Parent content type exists: ${parentResponse.d.Name}`);
                
                // Now create a new content type with proper inheritance
                // Generate a proper content type ID that follows SharePoint's format
                // For an Event content type child, should be something like 0x010200... not 0x0100...
                // The format is: Parent ID (e.g. 0x0102) + 00 + uniqueHex
                const uniqueSuffix = Array.from({length: 16}, () => 
                    Math.floor(Math.random() * 16).toString(16)).join('');
                
                const childContentTypeId = `${parentContentTypeId}00${uniqueSuffix}`;
                console.error(`Generated child content type ID: ${childContentTypeId}`);
                    
                // For debugging - check if parent id exists in standard content types
                if (parentContentTypeId === '0x0102') {
                    console.error("Creating child of Event content type");
                } else if (parentContentTypeId === '0x0101') {
                    console.error("Creating child of Document content type");
                } else if (parentContentTypeId === '0x01') {
                    console.error("Creating child of Item content type");
                }
                
                // Use a different approach - try a different URL format for CreateChildContentType
                console.error("Using alternate URL format for CreateChildContentType endpoint");
                
                try {
                    // Properly encode the content type ID for the URL
                    const encodedParentId = encodeURIComponent(parentContentTypeId);
                    
                    const response = await request({
                        url: `${url}/_api/web/contenttypes/GetById('${encodedParentId}')/CreateChildContentType`,
                        headers: postHeaders,
                        body: {
                            __metadata: { type: "SP.ContentType" },
                            Name: contentTypeData.Name,
                            Description: contentTypeData.Description || "",
                            Group: contentTypeData.Group || ""
                        },
                        method: 'POST',
                        json: true,
                        timeout: 30000
                    });
                    
                    console.error("Content type created successfully using CreateChildContentType endpoint:", response);
                    
                    return {
                        content: [{
                            type: "text",
                            text: JSON.stringify({
                                success: true,
                                message: `Site content type '${contentTypeData.Name}' created successfully as a child of '${parentContentTypeId}' using CreateChildContentType.`,
                                contentType: {
                                    Id: response.d.Id.StringValue,
                                    Name: response.d.Name,
                                    Group: response.d.Group || 'No group',
                                    Description: response.d.Description || 'No description',
                                    ParentContentTypeId: parentContentTypeId
                                }
                            }, null, 2)
                        }]
                    } as IToolResult;
                    
                } catch (createChildError) {
                    console.error("CreateChildContentType endpoint failed:", createChildError instanceof Error ? createChildError.message : String(createChildError));
                    
                    // Attempt another approach - using POST to the REST API
                    console.error("Trying direct content type creation with proper ID format");
                    
                    // Generate a unique ID based on the parent
                    const uniqueSuffix = Array.from({length: 16}, () => 
                        Math.floor(Math.random() * 16).toString(16)).join('').toUpperCase();
                        
                    // Using the proper ID format as documented in SharePoint: Parent ID + '00' + suffix
                    const childContentTypeId = `${parentContentTypeId}00${uniqueSuffix}`;
                    console.error(`Generated child content type ID: ${childContentTypeId}`);
                    
                    // Fall back to the old approach with SchemaXml
                    try {
                        const response = await request({
                            url: `${url}/_api/web/contenttypes`,
                            headers: postHeaders,
                            body: {
                                __metadata: { type: "SP.ContentType" },
                                Name: contentTypeData.Name,
                                Description: contentTypeData.Description || "",
                                Group: contentTypeData.Group || "",
                                // Use SchemaXml with properly formatted ID for inheritance
                                SchemaXml: `<ContentType ID="${childContentTypeId}" Name="${contentTypeData.Name}" Group="${contentTypeData.Group || ""}" Description="${contentTypeData.Description || ""}" Version="1">
                                    <Folder TargetName="_cts/${contentTypeData.Name}" />
                                </ContentType>`
                            },
                            method: 'POST',
                            json: true,
                            timeout: 30000
                        });
                        
                        console.error("Site content type created successfully:", response.d);
                        
                        return {
                            content: [{
                                type: "text",
                                text: JSON.stringify({
                                    success: true,
                                    message: `Site content type '${contentTypeData.Name}' created successfully with parent ID '${parentContentTypeId}'.`,
                                    contentType: {
                                        Id: response.d.Id.StringValue,
                                        Name: response.d.Name,
                                        Group: response.d.Group || 'No group',
                                        Description: response.d.Description || 'No description',
                                        ParentContentTypeId: parentContentTypeId,  // Always preserve the intended parent ID
                                        // Add a flag to indicate proper inheritance may not be reflected in SharePoint's ID
                                        ParentContentTypeNote: "SharePoint Online may show this content type as inheriting from Item rather than the specified parent."
                                    }
                                }, null, 2)
                            }]
                        } as IToolResult;
                    } catch (schemaError) {
                        console.error("Schema XML approach failed:", schemaError instanceof Error ? schemaError.message : String(schemaError));
                        throw schemaError; // Re-throw to be caught by the outer catch
                    }
                }
            } catch (parentError) {
                console.error(`Error getting parent content type: ${parentError instanceof Error ? parentError.message : String(parentError)}`);
                console.error("Falling back to direct content type creation");
                
                // Try direct creation with properly formatted content type ID
                // Format the ID to ensure proper inheritance
                const uniqueSuffix = Array.from({length: 8}, () => 
                    Math.floor(Math.random() * 16).toString(16)).join('');
                
                const childContentTypeId = `${parentContentTypeId}00${uniqueSuffix}`;
                console.error(`Generated child content type ID for fallback method: ${childContentTypeId}`);
                
                const response = await request({
                    url: `${url}/_api/web/contenttypes`,
                    headers: postHeaders,
                    body: {
                        __metadata: { type: "SP.ContentType" },
                        Name: contentTypeData.Name,
                        Description: contentTypeData.Description || "",
                        Group: contentTypeData.Group || "",
                        // Use SchemaXml with the properly formatted ID
                        SchemaXml: `<ContentType ID="${childContentTypeId}" Name="${contentTypeData.Name}" Group="${contentTypeData.Group || ""}" Description="${contentTypeData.Description || ""}" Version="1">
                            <Folder TargetName="_cts/${contentTypeData.Name}" />
                        </ContentType>`
                    },
                    method: 'POST',
                    json: true,
                    timeout: 30000
                });
                
                console.error("Site content type created successfully:", response.d);
                
                return {
                    content: [{
                        type: "text",
                        text: JSON.stringify({
                            success: true,
                            message: `Site content type '${contentTypeData.Name}' created successfully with parent ID '${parentContentTypeId}'.`,
                            contentType: {
                                Id: response.d.Id.StringValue,
                                Name: response.d.Name,
                                Group: response.d.Group || 'No group',
                                Description: response.d.Description || 'No description',
                                ParentContentTypeId: parentContentTypeId
                            }
                        }, null, 2)
                    }]
                } as IToolResult;
            }
        } catch (firstError) {
            console.error("First creation method failed, trying alternative approach");
            console.error("First method error:", firstError instanceof Error ? firstError.message : String(firstError));
            
            try {
                // Try alternative approach - SharePoint has a special method for content type inheritance
                console.error("Using direct content type creation with specific field formats");
                
                const response = await request({
                    url: `${url}/_api/web/contenttypes`,
                    headers: postHeaders,
                    body: {
                        __metadata: { type: "SP.ContentType" },
                        Name: contentTypeData.Name,
                        Description: contentTypeData.Description || "",
                        Group: contentTypeData.Group || "",
                        // Using various formats based on SharePoint's API requirements
                        // We've tried other methods of setting the parent, but those don't work
                        // Instead, we need to explicitly use the parent content type ID in the SchemaXml
                        SchemaXml: `<ContentType ID="${parentContentTypeId}00" Name="${contentTypeData.Name}" Group="${contentTypeData.Group || ""}" Description="${contentTypeData.Description || ""}" Version="1">
                            <Folder TargetName="_cts/${contentTypeData.Name}" />
                        </ContentType>`
                    },
                    method: 'POST',
                    json: true,
                    timeout: 30000
                });
                
                console.error("Content type created successfully via second method:", response);
                
                return {
                    content: [{
                        type: "text",
                        text: JSON.stringify({
                            success: true,
                            message: `Site content type '${contentTypeData.Name}' created successfully via second method with parent ID '${parentContentTypeId}'.`,
                            contentType: {
                                Id: response.d.Id.StringValue,
                                Name: response.d.Name,
                                Group: response.d.Group || 'No group',
                                Description: response.d.Description || 'No description',
                                ParentContentTypeId: parentContentTypeId
                            }
                        }, null, 2)
                    }]
                } as IToolResult;
            } catch (secondError) {
                console.error("Second creation method failed, trying one more approach");
                
                try {
                    // One more approach - using HTML to create the content type
                    // This is based on how SharePoint handles content type inheritance in its API
                    
                    // Construct the XML for the content type
                    // For proper inheritance, we need to format the ID correctly:
                    // ParentID + "00" + unique suffix
                    // This is how SharePoint creates child content types
                    const uniqueSuffix = Array.from({length: 8}, () => 
                        Math.floor(Math.random() * 16).toString(16)).join('');
                    
                    const contentTypeXml = `<ContentType ID="${parentContentTypeId}00${uniqueSuffix}" Name="${contentTypeData.Name}" Group="${contentTypeData.Group || ""}" Description="${contentTypeData.Description || ""}" Version="1">
                        <Folder TargetName="_cts/${contentTypeData.Name}" />
                    </ContentType>`;
                    
                    console.error("Trying with XML schema:", contentTypeXml);
                    
                    const response = await request({
                        url: `${url}/_api/web/contenttypes`,
                        headers: postHeaders,
                        body: {
                            __metadata: { type: "SP.ContentType" },
                            Name: contentTypeData.Name,
                            Description: contentTypeData.Description || "",
                            Group: contentTypeData.Group || "",
                            SchemaXml: contentTypeXml
                        },
                        method: 'POST',
                        json: true,
                        timeout: 30000
                    });
                    
                    console.error("Content type created successfully via XML method:", response);
                    
                    return {
                        content: [{
                            type: "text",
                            text: JSON.stringify({
                                success: true,
                                message: `Site content type '${contentTypeData.Name}' created successfully via XML method with parent ID '${parentContentTypeId}'.`,
                                contentType: {
                                    Id: response.d.Id.StringValue,
                                    Name: response.d.Name,
                                    Group: response.d.Group || 'No group',
                                    Description: response.d.Description || 'No description',
                                    ParentContentTypeId: parentContentTypeId
                                }
                            }, null, 2)
                        }]
                    } as IToolResult;
                } catch (thirdError) {
                    console.error("All creation methods failed");
                    
                    // Compile error information from both attempts
                    const errorDetails = {
                        firstAttempt: firstError instanceof Error ? firstError.message : String(firstError),
                        secondAttempt: secondError instanceof Error ? secondError.message : String(secondError),
                        thirdAttempt: thirdError instanceof Error ? thirdError.message : String(thirdError)
                    };
                    
                    console.error("Error details:", errorDetails);
                    
                    return {
                        content: [{
                            type: "text",
                            text: `Error creating site content type with parent ID '${parentContentTypeId}'. Multiple approaches failed:\n1. ${errorDetails.firstAttempt}\n2. ${errorDetails.secondAttempt}\n3. ${errorDetails.thirdAttempt}`
                        }],
                        isError: true
                    } as IToolResult;
                }
            }
        }
    } catch (error: unknown) {
        // Type-safe error handling
        let errorMessage: string;

        if (error instanceof Error) {
            errorMessage = error.message;
        } else if (typeof error === 'string') {
            errorMessage = error;
        } else {
            errorMessage = "Unknown error occurred";
        }

        console.error("Error in createSiteContentType tool:", errorMessage);

        return {
            content: [{
                type: "text",
                text: `Error creating site content type: ${errorMessage}`
            }],
            isError: true
        } as IToolResult;
    }
}

export default createSiteContentType;