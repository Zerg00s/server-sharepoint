// src/tools/deleteModernPage.ts
import request from 'request-promise';
import { IToolResult } from '../interfaces';
import { getSharePointHeaders, getRequestDigest } from '../auth_factory';
import { SharePointConfig } from '../config';

export interface DeleteModernPageParams {
    url: string;
    pageId: number;
    confirmation: string; // Must match the page title to confirm deletion
}

/**
 * Delete a modern page from SharePoint
 * @param params Parameters including site URL, page ID, and confirmation
 * @param config SharePoint configuration
 * @returns Tool result with deletion status
 */
export async function deleteModernPage(
    params: DeleteModernPageParams,
    config: SharePointConfig
): Promise<IToolResult> {
    const { url, pageId, confirmation } = params;
    console.error(`deleteModernPage tool called with URL: ${url}, Page ID: ${pageId}`);

    try {
        // Authenticate with SharePoint
        const headers = await getSharePointHeaders(url, config);
        console.error("Headers prepared with authentication");

        // Get request digest for DELETE operations
        headers['X-RequestDigest'] = await getRequestDigest(url, headers);
        console.error("Headers prepared with request digest");
        
        // First, get current page details
        console.error(`Getting current page details for page ID ${pageId}...`);
        let pageDetails;
        try {
            const getPageResponse = await request({
                url: `${url}/_api/sitepages/pages(${pageId})`,
                method: 'GET',
                headers: { ...headers, 'Content-Type': undefined },
                json: true,
                timeout: 30000
            });
            
            pageDetails = getPageResponse.d;
        } catch (error) {
            throw new Error(`Page with ID ${pageId} not found`);
        }
        
        // Check if page exists and get its title
        if (!pageDetails) {
            throw new Error(`Page with ID ${pageId} not found`);
        }
        
        const pageTitle = pageDetails.Title;
        const pageUrl = pageDetails.Url || pageDetails.AbsoluteUrl;
        const fileName = pageDetails.FileName;
        
        console.error(`Found page "${pageTitle}" with ID ${pageId}`);
        
        // Check confirmation matches page title
        if (confirmation !== pageTitle) {
            throw new Error(`Confirmation "${confirmation}" does not match page title "${pageTitle}". Deletion aborted.`);
        }
        
        // Prepare headers for DELETE operation
        headers['X-HTTP-Method'] = 'DELETE';
        headers['IF-MATCH'] = '*';
        headers['Content-Type'] = 'application/json;odata=verbose';
        
        // Delete the page
        console.error(`Deleting page "${pageTitle}" with ID ${pageId}...`);
        await request({
            url: `${url}/_api/sitepages/pages(${pageId})`,
            method: 'POST', // Using POST with X-HTTP-Method: DELETE
            headers: headers,
            body: '',
            timeout: 30000
        });
        
        console.error(`Page "${pageTitle}" successfully deleted`);
        
        return {
            content: [{
                type: "text",
                text: JSON.stringify({
                    success: true,
                    message: `Modern page "${pageTitle}" successfully deleted`,
                    deletedPage: {
                        id: pageId,
                        title: pageTitle,
                        url: pageUrl,
                        fileName: fileName
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

        console.error("Error in deleteModernPage tool:", errorMessage);

        return {
            content: [{
                type: "text",
                text: `Error deleting modern page: ${errorMessage}`
            }],
            isError: true
        } as IToolResult;
    }
}

export default deleteModernPage;
