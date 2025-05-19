// src/tools/createListContentType.ts
import request from 'request-promise';
import { IToolResult } from '../interfaces';
import { getSharePointHeaders, getRequestDigest } from '../auth_factory';
import { SharePointConfig } from '../config';

export interface CreateListContentTypeParams {
    url: string;
    listTitle: string;
    contentTypeData: {
        Name: string;
        Description?: string;
        ParentContentTypeId?: string;
        Group?: string;
    };
}

/**
 * Create a new content type in a SharePoint list
 * @param params Parameters including site URL, list title, and content type data
 * @param config SharePoint configuration
 * @returns Tool result with created content type information
 */
export async function createListContentType(
    params: CreateListContentTypeParams, 
    config: SharePointConfig
): Promise<IToolResult> {
    const { url, listTitle, contentTypeData } = params;
    console.error("createListContentType tool called with URL:", url, "list title:", listTitle);

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
        
        // Prepare data for content type creation
        const postData: any = {
            __metadata: { type: "SP.ContentType" },
            Name: contentTypeData.Name
        };
        
        // Add optional fields if provided
        if (contentTypeData.Description) {
            postData.Description = contentTypeData.Description;
        }
        
        if (contentTypeData.Group) {
            postData.Group = contentTypeData.Group;
        }
        
        if (contentTypeData.ParentContentTypeId) {
            postData.ParentContentTypeId = contentTypeData.ParentContentTypeId;
        }
        
        console.error("Making request to create content type:", JSON.stringify(postData));
        
        try {
            // First attempt - direct creation
            const response = await request({
                url: `${url}/_api/web/lists/getByTitle('${encodeURIComponent(listTitle)}')/contenttypes`,
                headers: postHeaders,
                body: postData,
                method: 'POST',
                json: true,
                timeout: 30000
            });
            
            console.error("Content type created successfully:", response.d);
            
            return {
                content: [{
                    type: "text",
                    text: JSON.stringify({
                        success: true,
                        message: `Content type '${contentTypeData.Name}' created successfully.`,
                        contentType: {
                            Id: response.d.Id.StringValue,
                            Name: response.d.Name,
                            Group: response.d.Group || 'No group',
                            Description: response.d.Description || 'No description'
                        }
                    }, null, 2)
                }]
            } as IToolResult;
        } catch (error) {
            console.error("First creation method failed, trying alternative approach");
            
            // Alternative approach - try adding content type by ID from parent web
            try {
                // Get the parent content type ID (default to Item content type)
                const parentContentTypeId = contentTypeData.ParentContentTypeId || "0x01";
                
                // Use the AddAvailableContentType endpoint
                const addResponse = await request({
                    url: `${url}/_api/web/lists/getByTitle('${encodeURIComponent(listTitle)}')/ContentTypes/AddAvailableContentType`,
                    headers: postHeaders,
                    body: {
                        contentTypeId: parentContentTypeId
                    },
                    method: 'POST',
                    json: true,
                    timeout: 30000
                });
                
                console.error("Content type added successfully via alternative method:", addResponse);
                
                return {
                    content: [{
                        type: "text",
                        text: JSON.stringify({
                            success: true,
                            message: `Content type based on ID '${parentContentTypeId}' added successfully.`,
                            contentType: {
                                Id: parentContentTypeId,
                                Name: contentTypeData.Name,
                                Group: contentTypeData.Group || 'No group',
                                Description: contentTypeData.Description || 'No description'
                            },
                            note: "Content type was added using the AddAvailableContentType method."
                        }, null, 2)
                    }]
                } as IToolResult;
            } catch (addError) {
                // If both methods fail, return the original error
                console.error("Both creation methods failed");
                
                // Type-safe error handling
                let errorMessage: string;
                
                if (error instanceof Error) {
                    errorMessage = error.message;
                } else if (typeof error === 'string') {
                    errorMessage = error;
                } else {
                    errorMessage = "Unknown error occurred";
                }
                
                console.error("Error in createListContentType tool:", errorMessage);
                
                return {
                    content: [{
                        type: "text",
                        text: `Error creating content type: ${errorMessage}`
                    }],
                    isError: true
                } as IToolResult;
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

        console.error("Error in createListContentType tool:", errorMessage);

        return {
            content: [{
                type: "text",
                text: `Error creating content type: ${errorMessage}`
            }],
            isError: true
        } as IToolResult;
    }
}

export default createListContentType;
