// src/tools/updateSite.ts
import request from 'request-promise';
import { IToolResult } from '../interfaces';
import { getSharePointHeaders, getRequestDigest } from '../auth';
import { SharePointConfig } from '../config';

export interface UpdateSiteParams {
    url: string;
    siteData: {
        Title?: string;
        Description?: string;
        LogoUrl?: string;
        [key: string]: any; // Allow for any other site properties
    };
}

/**
 * Update a SharePoint site properties
 * @param params Parameters including site URL and update data
 * @param config SharePoint configuration
 * @returns Tool result with update status
 */
export async function updateSite(
    params: UpdateSiteParams,
    config: SharePointConfig
): Promise<IToolResult> {
    const { url, siteData } = params;
    console.error(`updateSite tool called with URL: ${url}`);

    try {
        // Validate input
        if (Object.keys(siteData).length === 0) {
            throw new Error("No site properties provided for update");
        }

        // Authenticate with SharePoint
        const headers = await getSharePointHeaders(url, config);
        console.error("Headers prepared with authentication");

        // Get request digest for POST operations
        headers['X-RequestDigest'] = await getRequestDigest(url, headers);
        headers['Content-Type'] = 'application/json;odata=verbose';
        headers['X-HTTP-Method'] = 'MERGE';
        headers['IF-MATCH'] = '*';
        console.error("Headers prepared with request digest for update operation");
        
        // First, verify the site exists
        console.error(`Verifying site exists...`);
        let originalSite;
        try {
            originalSite = await request({
                url: `${url}/_api/web`,
                headers: { ...headers, 'Content-Type': undefined, 'X-HTTP-Method': undefined, 'IF-MATCH': undefined },
                json: true,
                method: 'GET',
                timeout: 15000
            });
        } catch (error) {
            throw new Error(`Site at URL "${url}" not found`);
        }
        
        // Prepare the update data
        const updatePayload: any = {
            __metadata: { type: 'SP.Web' },
            ...siteData
        };
        
        console.error(`Updating site with payload: ${JSON.stringify(updatePayload)}`);
        
        // Update the site
        await request({
            url: `${url}/_api/web`,
            headers: headers,
            json: true,
            method: 'POST',
            body: updatePayload,
            timeout: 30000
        });
        
        // Get the updated site details
        const updatedSite = await request({
            url: `${url}/_api/web`,
            headers: { ...headers, 'Content-Type': undefined, 'X-HTTP-Method': undefined, 'IF-MATCH': undefined },
            json: true,
            method: 'GET',
            timeout: 15000
        });
        
        // Prepare the response with updated properties
        const updatedProperties: Record<string, any> = {};
        for (const key of Object.keys(siteData)) {
            if (updatedSite.d[key] !== undefined) {
                updatedProperties[key] = updatedSite.d[key];
            }
        }

        return {
            content: [{
                type: "text",
                text: JSON.stringify({
                    success: true,
                    message: `Site properties successfully updated`,
                    updatedSite: {
                        id: updatedSite.d.Id,
                        title: updatedSite.d.Title,
                        url: url,
                        updatedProperties: updatedProperties
                    }
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

        console.error("Error in updateSite tool:", errorMessage);

        return {
            content: [{
                type: "text",
                text: `Error updating site: ${errorMessage}`
            }],
            isError: true
        } as IToolResult;
    }
}

export default updateSite;
