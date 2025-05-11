// src/tools/deleteListView.ts
import request from 'request-promise';
import { IToolResult } from '../interfaces';
import { getSharePointHeaders, getRequestDigest } from '../auth';
import { SharePointConfig } from '../config';

export interface DeleteListViewParams {
    url: string;
    listTitle: string;
    viewTitle: string;
}

/**
 * Delete a view from a SharePoint list
 * @param params Parameters including site URL, list title, and view title
 * @param config SharePoint configuration
 * @returns Tool result with deletion status
 */
export async function deleteListView(
    params: DeleteListViewParams,
    config: SharePointConfig
): Promise<IToolResult> {
    const { url, listTitle, viewTitle } = params;
    console.error(`deleteListView tool called with URL: ${url}, List Title: ${listTitle}, View Title: ${viewTitle}`);

    try {
        // Authenticate with SharePoint
        const headers = await getSharePointHeaders(url, config);
        console.error("Headers prepared with authentication");

        // Get request digest for DELETE operations
        headers['X-RequestDigest'] = await getRequestDigest(url, headers);
        headers['X-HTTP-Method'] = 'DELETE';
        headers['IF-MATCH'] = '*';
        console.error("Headers prepared with request digest for delete operation");

        // Encode the list title and view title to handle special characters
        const encodedListTitle = encodeURIComponent(listTitle);
        const encodedViewTitle = encodeURIComponent(viewTitle);
        
        // First, verify the list and view exist
        console.error(`Verifying list "${listTitle}" and view "${viewTitle}" exist...`);
        let viewDetails;
        try {
            viewDetails = await request({
                url: `${url}/_api/web/lists/getByTitle('${encodedListTitle}')/views/getByTitle('${encodedViewTitle}')`,
                headers: { ...headers, 'X-HTTP-Method': undefined, 'IF-MATCH': undefined },
                json: true,
                method: 'GET',
                timeout: 15000
            });
        } catch (error) {
            throw new Error(`View "${viewTitle}" not found in list "${listTitle}"`);
        }
        
        // Check if it's the default view - we can't delete the default view
        if (viewDetails.d.DefaultView) {
            throw new Error(`Cannot delete the default view "${viewTitle}". Please set another view as default first.`);
        }
        
        // Store the view ID and details for response
        const viewId = viewDetails.d.Id;
        const viewInfo = {
            id: viewId,
            title: viewDetails.d.Title,
            url: `${url}${viewDetails.d.ServerRelativeUrl}`
        };
        
        // Delete the view
        console.error(`Deleting view "${viewTitle}" with ID ${viewId}...`);
        await request({
            url: `${url}/_api/web/lists/getByTitle('${encodedListTitle}')/views/getByTitle('${encodedViewTitle}')`,
            headers: headers,
            method: 'POST', // POST with DELETE X-HTTP-Method header
            body: '',
            timeout: 20000
        });
        
        return {
            content: [{
                type: "text",
                text: JSON.stringify({
                    success: true,
                    message: `View "${viewTitle}" successfully deleted from list "${listTitle}"`,
                    deletedView: viewInfo
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

        console.error("Error in deleteListView tool:", errorMessage);

        return {
            content: [{
                type: "text",
                text: `Error deleting list view: ${errorMessage}`
            }],
            isError: true
        } as IToolResult;
    }
}

export default deleteListView;
