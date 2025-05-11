// src/tools/deleteSubsite.ts
import request from 'request-promise';
import { IToolResult } from '../interfaces';
import { getSharePointHeaders, getRequestDigest } from '../auth_factory';
import { SharePointConfig } from '../config';

export interface DeleteSubsiteParams {
    url: string;
    confirmation?: string; // Optional confirmation string to prevent accidental deletion
}

/**
 * Delete a SharePoint subsite
 * @param params Parameters including site URL and confirmation
 * @param config SharePoint configuration
 * @returns Tool result with deletion status
 */
export async function deleteSubsite(
    params: DeleteSubsiteParams,
    config: SharePointConfig
): Promise<IToolResult> {
    const { url, confirmation } = params;
    console.error(`deleteSubsite tool called with URL: ${url}`);

    try {
        // Extract site title from URL for confirmation
        const urlParts = url.split('/');
        let siteTitle = urlParts[urlParts.length - 1];
        if (siteTitle === '') {
            siteTitle = urlParts[urlParts.length - 2];
        }
        
        // Check confirmation string to prevent accidental deletion
        if (!confirmation || confirmation !== siteTitle) {
            throw new Error(`To delete the subsite, please provide a confirmation parameter that matches exactly the site name '${siteTitle}' from the URL`);
        }

        // Authenticate with SharePoint
        const headers = await getSharePointHeaders(url, config);
        console.error("Headers prepared with authentication");

        // Get request digest for DELETE operations
        headers['X-RequestDigest'] = await getRequestDigest(url, headers);
        headers['X-HTTP-Method'] = 'DELETE';
        headers['IF-MATCH'] = '*';
        console.error("Headers prepared with request digest for delete operation");
        
        // First, verify the site exists and get its details
        console.error(`Verifying site exists...`);
        let siteDetails;
        try {
            siteDetails = await request({
                url: `${url}/_api/web`,
                headers: { ...headers, 'X-HTTP-Method': undefined, 'IF-MATCH': undefined },
                json: true,
                method: 'GET',
                timeout: 15000
            });
        } catch (error) {
            throw new Error(`Site at URL "${url}" not found`);
        }
        
        // Store the site details for response
        const siteInfo = {
            id: siteDetails.d.Id,
            title: siteDetails.d.Title,
            url: url,
            created: siteDetails.d.Created
        };
        
        // Delete the site
        console.error(`Deleting site at URL "${url}"...`);
        await request({
            url: `${url}/_api/web`,
            headers: headers,
            method: 'POST', // POST with DELETE X-HTTP-Method header
            body: '',
            timeout: 60000 // Longer timeout for site deletion
        });
        
        return {
            content: [{
                type: "text",
                text: JSON.stringify({
                    success: true,
                    message: `Site "${siteInfo.title}" successfully deleted`,
                    deletedSite: siteInfo
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

        console.error("Error in deleteSubsite tool:", errorMessage);

        return {
            content: [{
                type: "text",
                text: `Error deleting subsite: ${errorMessage}`
            }],
            isError: true
        } as IToolResult;
    }
}

export default deleteSubsite;

