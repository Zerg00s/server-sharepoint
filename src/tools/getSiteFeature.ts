// src/tools/getSiteFeature.ts
import request from 'request-promise';
import { ISharePointWebResponse, IToolResult } from '../interfaces';
import { getSharePointHeaders } from '../auth_factory';
import { SharePointConfig } from '../config';

/**
 * Parameters for getSiteFeature tool
 */
export interface GetSiteFeatureParams {
    /**
     * URL of the SharePoint website
     */
    url: string;
    
    /**
     * Feature ID (GUID) of the feature to retrieve
     */
    featureId: string;
}

/**
 * Get a specific feature from a SharePoint site (web) by feature ID
 * @param params Parameters including the SharePoint site URL and feature ID
 * @param config SharePoint configuration
 * @returns Site feature information
 */
export default async function getSiteFeature(
    params: GetSiteFeatureParams,
    config: SharePointConfig
): Promise<IToolResult> {
    try {
        // Ensure the URL ends with a trailing slash
        const baseUrl = params.url.endsWith('/') ? params.url : `${params.url}/`;
        
        // Get feature ID and remove any braces if present
        const featureId = params.featureId.replace(/[{}]/g, '');
        
        // Get headers with authentication token
        const headers = await getSharePointHeaders(baseUrl, config);
        
        // Build the request URL for the specific site feature
        const requestUrl = `${baseUrl}_api/web/features(guid'${featureId}')`;
        
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

        console.error('Error in getSiteFeature:', errorMessage);
        
        return {
            content: [{ 
                type: "text",
                text: `Error fetching site feature: ${errorMessage}` 
            }],
            isError: true
        } as IToolResult;
    }
}
