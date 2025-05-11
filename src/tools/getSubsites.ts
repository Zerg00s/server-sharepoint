// src/tools/getSubsites.ts
import request from 'request-promise';
import { IToolResult } from '../interfaces';
import { getSharePointHeaders } from '../auth';
import { SharePointConfig } from '../config';

export interface GetSubsitesParams {
    url: string;
}

/**
 * Get all subsites from a SharePoint site
 * @param params Parameters including site URL
 * @param config SharePoint configuration
 * @returns Tool result with subsites data
 */
export async function getSubsites(
    params: GetSubsitesParams,
    config: SharePointConfig
): Promise<IToolResult> {
    const { url } = params;
    console.error(`getSubsites tool called with URL: ${url}`);

    try {
        // Authenticate with SharePoint
        const headers = await getSharePointHeaders(url, config);
        console.error("Headers prepared with authentication");

        // Get all subsites
        console.error("Getting subsites...");
        const subsitesResponse = await request({
            url: `${url}/_api/web/webinfos`,
            headers: headers,
            json: true,
            method: 'GET',
            timeout: 30000
        });
        
        // Process the subsites
        const subsites = subsitesResponse.d.results;
        console.error(`Retrieved ${subsites.length} subsites`);
        
        // Format the subsites for display
        const formattedSubsites = subsites.map((site: any) => ({
            Title: site.Title,
            Url: site.ServerRelativeUrl,
            Description: site.Description || '',
            Created: site.Created,
            WebTemplate: site.WebTemplate,
            WebTemplateTitle: site.WebTemplateTitle,
            Id: site.Id
        }));

        return {
            content: [{
                type: "text",
                text: JSON.stringify({
                    parentSiteUrl: url,
                    subsitesCount: formattedSubsites.length,
                    subsites: formattedSubsites
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

        console.error("Error in getSubsites tool:", errorMessage);

        return {
            content: [{
                type: "text",
                text: `Error getting subsites: ${errorMessage}`
            }],
            isError: true
        } as IToolResult;
    }
}

export default getSubsites;
