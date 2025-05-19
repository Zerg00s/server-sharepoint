// src/tools/getListContentTypes.ts
import request from 'request-promise';
import { ISharePointContentType, IContentType, IToolResult } from '../interfaces';
import { getSharePointHeaders } from '../auth_factory';
import { SharePointConfig } from '../config';

export interface GetListContentTypesParams {
    url: string;
    listTitle: string;
}

/**
 * Get all content types from a specific SharePoint list
 * @param params Parameters including the SharePoint site URL and list title
 * @param config SharePoint configuration
 * @returns Tool result with content types data
 */
export async function getListContentTypes(
    params: GetListContentTypesParams, 
    config: SharePointConfig
): Promise<IToolResult> {
    const { url, listTitle } = params;
    console.error("getListContentTypes tool called with URL:", url, "and list title:", listTitle);

    try {
        // Authenticate with SharePoint
        const headers = await getSharePointHeaders(url, config);
        console.error("Headers prepared:", headers);

        // Make request to SharePoint API to get list content types
        console.error(`Making request to SharePoint API for content types in list: ${listTitle}...`);
        const response = await request({
            url: `${url}/_api/web/lists/getByTitle('${encodeURIComponent(listTitle)}')/contenttypes`,
            headers: headers,
            json: true,
            method: 'GET',
            timeout: 15000
        });

        console.error(`SharePoint API response received with ${response.d.results.length} content types`);
        
        // Format the content type data for display
        const contentTypes: IContentType[] = response.d.results.map((contentType: ISharePointContentType) => {
            return {
                Id: contentType.Id.StringValue,
                Name: contentType.Name,
                Group: contentType.Group || 'No group',
                Description: contentType.Description || 'No description',
                Hidden: contentType.Hidden,
                ReadOnly: contentType.ReadOnly,
                Sealed: contentType.Sealed
            };
        });

        return {
            content: [{
                type: "text",
                text: JSON.stringify(contentTypes, null, 2)
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

        console.error("Error in getListContentTypes tool:", errorMessage);

        return {
            content: [{
                type: "text",
                text: `Error fetching content types: ${errorMessage}`
            }],
            isError: true
        } as IToolResult;
    }
}

export default getListContentTypes;
