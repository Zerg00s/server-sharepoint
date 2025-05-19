// src/tools/updateListContentType.ts
import request from 'request-promise';
import { IToolResult } from '../interfaces';
import { getSharePointHeaders, getRequestDigest } from '../auth_factory';
import { SharePointConfig } from '../config';

export interface UpdateListContentTypeParams {
    url: string;
    listTitle: string;
    contentTypeId: string;
    updateData: {
        Name?: string;
        Description?: string;
        Group?: string;
    };
}

/**
 * Update a content type in a SharePoint list
 * @param params Parameters including site URL, list title, content type ID, and update data
 * @param config SharePoint configuration
 * @returns Tool result with updated content type information
 */
export async function updateListContentType(
    params: UpdateListContentTypeParams, 
    config: SharePointConfig
): Promise<IToolResult> {
    const { url, listTitle, contentTypeId, updateData } = params;
    console.error("updateListContentType tool called with URL:", url, "list title:", listTitle, "contentTypeId:", contentTypeId);

    try {
        // Authenticate with SharePoint and get request digest
        const headers = await getSharePointHeaders(url, config);
        const requestDigest = await getRequestDigest(url, headers);
        
        // Set up headers for MERGE request
        const mergeHeaders = {
            ...headers,
            "X-RequestDigest": requestDigest,
            "content-type": "application/json;odata=verbose",
            "accept": "application/json;odata=verbose",
            "IF-MATCH": "*",
            "X-HTTP-Method": "MERGE"
        };
        
        // Prepare data for content type update
        const updateBody: any = {
            __metadata: { type: "SP.ContentType" }
        };
        
        // Add fields to update if provided
        if (updateData.Name) {
            updateBody.Name = updateData.Name;
        }
        
        if (updateData.Description !== undefined) {
            updateBody.Description = updateData.Description;
        }
        
        if (updateData.Group !== undefined) {
            updateBody.Group = updateData.Group;
        }
        
        // Check if there are actually fields to update
        if (Object.keys(updateBody).length <= 1) {
            return {
                content: [{
                    type: "text",
                    text: "No fields provided for update."
                }],
                isError: true
            } as IToolResult;
        }
        
        console.error("Making request to update content type:", JSON.stringify(updateBody));
        
        try {
            // Make request to SharePoint API to update content type
            // The contenttype endpoint in SharePoint uses the content type ID
            // Need to properly encode the content type ID in the URL
            const encodedContentTypeId = encodeURIComponent(contentTypeId);
            await request({
                url: `${url}/_api/web/lists/getByTitle('${encodeURIComponent(listTitle)}')/contenttypes('${encodedContentTypeId}')`,
                headers: mergeHeaders,
                body: updateBody,
                method: 'POST',
                json: true,
                timeout: 30000
            });
            
            console.error("Content type updated successfully");
            
            // Get the updated content type to return in the response
            const getResponse = await request({
                url: `${url}/_api/web/lists/getByTitle('${encodeURIComponent(listTitle)}')/contenttypes('${encodedContentTypeId}')`,
                headers: headers,
                method: 'GET',
                json: true,
                timeout: 15000
            });
            
            return {
                content: [{
                    type: "text",
                    text: JSON.stringify({
                        success: true,
                        message: `Content type with ID '${contentTypeId}' updated successfully.`,
                        contentType: {
                            Id: getResponse.d.Id.StringValue,
                            Name: getResponse.d.Name,
                            Group: getResponse.d.Group || 'No group',
                            Description: getResponse.d.Description || 'No description'
                        }
                    }, null, 2)
                }]
            } as IToolResult;
        } catch (error) {
            // Try an alternative approach for updating
            console.error("Standard update method failed, trying alternative approach");
            
            try {
                // Sometimes we need to use a different URL format for certain content types
                // Try with the ID without quotes
                await request({
                    url: `${url}/_api/web/lists/getByTitle('${encodeURIComponent(listTitle)}')/contenttypes/${contentTypeId}`,
                    headers: mergeHeaders,
                    body: updateBody,
                    method: 'POST',
                    json: true,
                    timeout: 30000
                });
                
                console.error("Content type updated successfully via alternative method");
                
                return {
                    content: [{
                        type: "text",
                        text: JSON.stringify({
                            success: true,
                            message: `Content type with ID '${contentTypeId}' updated successfully via alternative method.`,
                            note: "Content type was updated using an alternative URL format."
                        }, null, 2)
                    }]
                } as IToolResult;
            } catch (alternativeError) {
                // If both methods fail, return the original error
                // Type-safe error handling
                let errorMessage: string;
                
                if (error instanceof Error) {
                    errorMessage = error.message;
                } else if (typeof error === 'string') {
                    errorMessage = error;
                } else {
                    errorMessage = "Unknown error occurred";
                }
                
                console.error("Error in updateListContentType tool:", errorMessage);
                
                return {
                    content: [{
                        type: "text",
                        text: `Error updating content type: ${errorMessage}`
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

        console.error("Error in updateListContentType tool:", errorMessage);

        return {
            content: [{
                type: "text",
                text: `Error updating content type: ${errorMessage}`
            }],
            isError: true
        } as IToolResult;
    }
}

export default updateListContentType;
