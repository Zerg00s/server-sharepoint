// src/tools/getSiteFeatures.ts
import request from 'request-promise';
import { ISharePointWebResponse, IToolResult } from '../interfaces';
import { getSharePointHeaders } from '../auth_factory';
import { SharePointConfig } from '../config';

/**
 * Parameters for getSiteFeatures tool
 */
export interface GetSiteFeaturesParams {
    /**
     * URL of the SharePoint website
     */
    url: string;
}

/**
 * Get all features from a SharePoint site (web)
 * @param params Parameters including the SharePoint site URL
 * @param config SharePoint configuration
 * @returns Site features information
 */
export default async function getSiteFeatures(
    params: GetSiteFeaturesParams,
    config: SharePointConfig
): Promise<IToolResult> {
    try {
        // Ensure the URL ends with a trailing slash
        const baseUrl = params.url.endsWith('/') ? params.url : `${params.url}/`;
        
        // Get headers with authentication token
        const headers = await getSharePointHeaders(baseUrl, config);
        
        // Build the request URL for site features
        const requestUrl = `${baseUrl}_api/web/features`;
        
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
                text: JSON.stringify(response.d.results, null, 2) 
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

        console.error('Error in getSiteFeatures:', errorMessage);
        
        return {
            content: [{ 
                type: "text",
                text: `Error fetching site features: ${errorMessage}` 
            }],
            isError: true
        } as IToolResult;
    }
}
