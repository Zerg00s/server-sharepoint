// src/tools/getSiteContentTypes.ts
import request from 'request-promise';
import { ISharePointContentType, IContentType, IToolResult } from '../interfaces';
import { getSharePointHeaders } from '../auth_factory';
import { SharePointConfig } from '../config';

export interface GetSiteContentTypesParams {
    url: string;
}

/**
 * Get all content types from a SharePoint site
 * @param params Parameters including the SharePoint site URL
 * @param config SharePoint configuration
 * @returns Tool result with content types data
 */
export async function getSiteContentTypes(
    params: GetSiteContentTypesParams, 
    config: SharePointConfig
): Promise<IToolResult> {
    const { url } = params;
    console.error("getSiteContentTypes tool called with URL:", url);

    try {
        // Authenticate with SharePoint
        const headers = await getSharePointHeaders(url, config);
        console.error("Headers prepared:", headers);

        // Make request to SharePoint API to get site content types
        console.error(`Making request to SharePoint API for site content types...`);
        const response = await request({
            url: `${url}/_api/web/contenttypes`,
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

        console.error("Error in getSiteContentTypes tool:", errorMessage);

        return {
            content: [{
                type: "text",
                text: `Error fetching site content types: ${errorMessage}`
            }],
            isError: true
        } as IToolResult;
    }
}

export default getSiteContentTypes;
