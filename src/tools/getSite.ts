// src/tools/getSite.ts
import request from 'request-promise';
import { ISharePointWebResponse, IToolResult } from '../interfaces';
import { getSharePointHeaders } from '../auth';
import { SharePointConfig } from '../config';

export interface GetSiteParams {
    url: string;
}

/**
 * Get the details of a SharePoint website
 * @param params Parameters including the SharePoint site URL
 * @param config SharePoint configuration
 * @returns Tool result with site details
 */
export async function getSite(
    params: GetSiteParams, 
    config: SharePointConfig
): Promise<IToolResult> {
    const { url } = params;
    console.error("getSite tool called with URL:", url);

    try {
        // Authenticate with SharePoint
        const headers = await getSharePointHeaders(url, config);
        console.error("Headers prepared:", headers);

        // Make request to SharePoint API
        console.error("Making request to SharePoint API...");
        const response = await request({
            url: `${url}/_api/web`,
            headers: headers,
            json: true,
            method: 'GET',
            timeout: 8000
        }) as ISharePointWebResponse;

        console.error("SharePoint API response received");
        console.error("SharePoint site:", response.d);

        return {
            content: [{
                type: "text",
                text: `SharePoint site: ${response.d}`
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

        console.error("Error in getSite tool:", errorMessage);

        return {
            content: [{
                type: "text",
                text: `Error fetching site: ${errorMessage}`
            }],
            isError: true
        } as IToolResult;
    }
}

export default getSite;