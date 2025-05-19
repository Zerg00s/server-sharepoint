// src/tools/getListContentType.ts
import request from 'request-promise';
import { ISharePointWebResponse, IToolResult } from '../interfaces';
import { getSharePointHeaders } from '../auth_factory';
import { SharePointConfig } from '../config';

/**
 * Parameters for getListContentType tool
 */
export interface GetListContentTypeParams {
    /**
     * URL of the SharePoint website
     */
    url: string;

    /**
     * Title of the SharePoint list
     */
    listTitle: string;

    /**
     * ID of the content type to retrieve
     */
    contentTypeId: string;
}

/**
 * Get a specific content type from a SharePoint list
 * @param params Parameters including the SharePoint site URL, list title, and content type ID
 * @param config SharePoint configuration
 * @returns List content type information
 */
export default async function getListContentType(
    params: GetListContentTypeParams,
    config: SharePointConfig
): Promise<IToolResult> {
    try {
        // Ensure the URL ends with a trailing slash
        const baseUrl = params.url.endsWith('/') ? params.url : `${params.url}/`;
        
        // Get headers with authentication token
        const headers = await getSharePointHeaders(baseUrl, config);
        
        // Clean the content type ID (remove any curly braces if present)
        const contentTypeId = params.contentTypeId.replace(/[{}]/g, '');
        
        // Encode the list title for the URL
        const encodedListTitle = encodeURIComponent(params.listTitle);
        
        // Build the request URL for the specific list content type
        const requestUrl = `${baseUrl}_api/web/lists/getbytitle('${encodedListTitle}')/contenttypes('${contentTypeId}')`;
        
        // Make the request to SharePoint REST API
        const response = await request({
            url: requestUrl,
            headers: headers,
            json: true,
            method: 'GET',
            timeout: 8000
        }) as ISharePointWebResponse;
        
        return {
            content: [{ 
                type: "text",
                text: JSON.stringify(response.d, null, 2) 
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

        console.error('Error in getListContentType:', errorMessage);
        
        return {
            content: [{ 
                type: "text",
                text: `Error fetching list content type: ${errorMessage}` 
            }],
            isError: true
        } as IToolResult;
    }
}