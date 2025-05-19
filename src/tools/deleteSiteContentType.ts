// src/tools/deleteSiteContentType.ts
import request from 'request-promise';
import { IToolResult } from '../interfaces';
import { getSharePointHeaders, getRequestDigest } from '../auth_factory';
import { SharePointConfig } from '../config';

export interface DeleteSiteContentTypeParams {
    url: string;
    contentTypeId: string;
    confirmation: string;
}

/**
 * Delete a content type from a SharePoint site
 * @param params Parameters including site URL, content type ID, and confirmation
 * @param config SharePoint configuration
 * @returns Tool result with deletion status
 */
export async function deleteSiteContentType(
    params: DeleteSiteContentTypeParams, 
    config: SharePointConfig
): Promise<IToolResult> {
    const { url, contentTypeId, confirmation } = params;
    console.error("deleteSiteContentType tool called with URL:", url, "contentTypeId:", contentTypeId);

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
        
        console.error("Making request to delete site content type with ID:", contentTypeId);
        
        try {
            // Make request to SharePoint API to delete content type
            const encodedContentTypeId = encodeURIComponent(contentTypeId);
            await request({
                url: `${url}/_api/web/contenttypes('${encodedContentTypeId}')`,
                headers: deleteHeaders,
                method: 'POST',
                timeout: 30000
            });
            
            console.error("Site content type deleted successfully");
            
            return {
                content: [{
                    type: "text",
                    text: JSON.stringify({
                        success: true,
                        message: `Site content type with ID '${contentTypeId}' deleted successfully.`
                    }, null, 2)
                }]
            } as IToolResult;
        } catch (error) {
            // Try alternative approach
            console.error("Standard deletion method failed, trying alternative approach");
            
            try {
                // Try with the ID without quotes
                await request({
                    url: `${url}/_api/web/contenttypes/${contentTypeId}`,
                    headers: deleteHeaders,
                    method: 'POST',
                    timeout: 30000
                });
                
                console.error("Site content type deleted successfully via alternative method");
                
                return {
                    content: [{
                        type: "text",
                        text: JSON.stringify({
                            success: true,
                            message: `Site content type with ID '${contentTypeId}' deleted successfully via alternative method.`
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
                
                console.error("Error in deleteSiteContentType tool:", errorMessage);
                
                return {
                    content: [{
                        type: "text",
                        text: `Error deleting site content type: ${errorMessage}`
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

        console.error("Error in deleteSiteContentType tool:", errorMessage);

        return {
            content: [{
                type: "text",
                text: `Error deleting site content type: ${errorMessage}`
            }],
            isError: true
        } as IToolResult;
    }
}

export default deleteSiteContentType;
