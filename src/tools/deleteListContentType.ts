// src/tools/deleteListContentType.ts
import request from 'request-promise';
import { IToolResult } from '../interfaces';
import { getSharePointHeaders, getRequestDigest } from '../auth_factory';
import { SharePointConfig } from '../config';

export interface DeleteListContentTypeParams {
    url: string;
    listTitle: string;
    contentTypeId: string;
    confirmation: string;
}

/**
 * Delete a content type from a SharePoint list
 * @param params Parameters including site URL, list title, content type ID, and confirmation
 * @param config SharePoint configuration
 * @returns Tool result with deletion status
 */
export async function deleteListContentType(
    params: DeleteListContentTypeParams, 
    config: SharePointConfig
): Promise<IToolResult> {
    const { url, listTitle, contentTypeId, confirmation } = params;
    console.error("deleteListContentType tool called with URL:", url, "list title:", listTitle, "contentTypeId:", contentTypeId);

    // Validate confirmation matches content type ID
    if (confirmation !== contentTypeId) {
        return {
            content: [{
                type: "text",
                text: "Confirmation text does not match content type ID. Deletion aborted."
            }],
            isError: true
        } as IToolResult;
    }

    try {
        // Authenticate with SharePoint and get request digest
        const headers = await getSharePointHeaders(url, config);
        const requestDigest = await getRequestDigest(url, headers);
        
        // Set up headers for DELETE request
        const deleteHeaders = {
            ...headers,
            "X-RequestDigest": requestDigest,
            "IF-MATCH": "*",
            "X-HTTP-Method": "DELETE"
        };
        
        console.error("Making request to delete content type with ID:", contentTypeId);
        
        // Make request to SharePoint API to delete content type
        const encodedContentTypeId = encodeURIComponent(contentTypeId);
        await request({
            url: `${url}/_api/web/lists/getByTitle('${encodeURIComponent(listTitle)}')/contenttypes('${encodedContentTypeId}')`,
            headers: deleteHeaders,
            method: 'POST',
            timeout: 30000
        });
        
        console.error("Content type deleted successfully");
        
        return {
            content: [{
                type: "text",
                text: JSON.stringify({
                    success: true,
                    message: `Content type with ID '${contentTypeId}' deleted successfully from list '${listTitle}'.`
                }, null, 2)
            }]
        } as IToolResult;
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

        console.error("Error in deleteListContentType tool:", errorMessage);

        return {
            content: [{
                type: "text",
                text: `Error deleting content type: ${errorMessage}`
            }],
            isError: true
        } as IToolResult;
    }
}

export default deleteListContentType;
