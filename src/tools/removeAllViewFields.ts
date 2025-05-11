// src/tools/removeAllViewFields.ts
import request from 'request-promise';
import { IToolResult } from '../interfaces';
import { getSharePointHeaders, getRequestDigest } from '../auth';
import { SharePointConfig } from '../config';

export interface RemoveAllViewFieldsParams {
    url: string;
    listTitle: string;
    viewTitle: string;
}

/**
 * Remove all fields from a SharePoint list view
 * @param params Parameters including site URL, list title, and view title
 * @param config SharePoint configuration
 * @returns Tool result with removal status
 */
export async function removeAllViewFields(
    params: RemoveAllViewFieldsParams,
    config: SharePointConfig
): Promise<IToolResult> {
    const { url, listTitle, viewTitle } = params;
    console.error(`removeAllViewFields tool called with URL: ${url}, List Title: ${listTitle}, View Title: ${viewTitle}`);

    try {
        // Authenticate with SharePoint
        const headers = await getSharePointHeaders(url, config);
        console.error("Headers prepared with authentication");

        // Get request digest for POST operations
        headers['X-RequestDigest'] = await getRequestDigest(url, headers);
        headers['Content-Type'] = 'application/json;odata=verbose';
        console.error("Headers prepared with request digest");

        // Encode the list title and view title to handle special characters
        const encodedListTitle = encodeURIComponent(listTitle);
        const encodedViewTitle = encodeURIComponent(viewTitle);
        
        // Verify the view exists
        console.error(`Verifying view "${viewTitle}" exists...`);
        let viewDetails;
        try {
            viewDetails = await request({
                url: `${url}/_api/web/lists/getByTitle('${encodedListTitle}')/views/getByTitle('${encodedViewTitle}')`,
                headers: { ...headers, 'Content-Type': undefined },
                json: true,
                method: 'GET',
                timeout: 15000
            });
        } catch (error) {
            throw new Error(`View "${viewTitle}" not found in list "${listTitle}"`);
        }
        
        // Get the current view fields first to know how many we're removing
        console.error(`Getting current fields for view "${viewTitle}"...`);
        let fieldCount = 0;
        try {
            const viewFieldsResponse = await request({
                url: `${url}/_api/web/lists/getByTitle('${encodedListTitle}')/views/getByTitle('${encodedViewTitle}')/viewFields`,
                headers: { ...headers, 'Content-Type': undefined },
                json: true,
                method: 'GET',
                timeout: 20000
            });
            
            fieldCount = (viewFieldsResponse.d.Items.results || []).length;
        } catch (error) {
            console.error(`Warning: Could not get current field count: ${error instanceof Error ? error.message : String(error)}`);
        }
        
        // Remove all fields from the view
        console.error(`Removing all fields from view "${viewTitle}"...`);
        await request({
            url: `${url}/_api/web/lists/getByTitle('${encodedListTitle}')/views/getByTitle('${encodedViewTitle}')/viewFields/removeAllViewFields`,
            headers: headers,
            json: true,
            method: 'POST',
            body: {},
            timeout: 20000
        });
        
        return {
            content: [{
                type: "text",
                text: JSON.stringify({
                    success: true,
                    message: `All fields successfully removed from view "${viewTitle}" in list "${listTitle}"`,
                    listTitle: listTitle,
                    viewTitle: viewTitle,
                    viewId: viewDetails.d.Id,
                    removedFieldCount: fieldCount
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

        console.error("Error in removeAllViewFields tool:", errorMessage);

        return {
            content: [{
                type: "text",
                text: `Error removing all view fields: ${errorMessage}`
            }],
            isError: true
        } as IToolResult;
    }
}

export default removeAllViewFields;