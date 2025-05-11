// src/tools/removeViewField.ts
import request from 'request-promise';
import { IToolResult } from '../interfaces';
import { getSharePointHeaders, getRequestDigest } from '../auth_factory';
import { SharePointConfig } from '../config';

export interface RemoveViewFieldParams {
    url: string;
    listTitle: string;
    viewTitle: string;
    fieldName: string; // Internal name of the field to remove
}

/**
 * Remove a field from a SharePoint list view
 * @param params Parameters including site URL, list title, view title, and field name
 * @param config SharePoint configuration
 * @returns Tool result with removal status
 */
export async function removeViewField(
    params: RemoveViewFieldParams,
    config: SharePointConfig
): Promise<IToolResult> {
    const { url, listTitle, viewTitle, fieldName } = params;
    console.error(`removeViewField tool called with URL: ${url}, List Title: ${listTitle}, View Title: ${viewTitle}, Field: ${fieldName}`);

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
        
        // Remove the field from the view
        console.error(`Removing field "${fieldName}" from view "${viewTitle}"...`);
        await request({
            url: `${url}/_api/web/lists/getByTitle('${encodedListTitle}')/views/getByTitle('${encodedViewTitle}')/viewFields/removeViewField`,
            headers: headers,
            json: true,
            method: 'POST',
            body: { "strField": fieldName },
            timeout: 20000
        });
        
        return {
            content: [{
                type: "text",
                text: JSON.stringify({
                    success: true,
                    message: `Field "${fieldName}" successfully removed from view "${viewTitle}" in list "${listTitle}"`,
                    listTitle: listTitle,
                    viewTitle: viewTitle,
                    viewId: viewDetails.d.Id,
                    removedField: fieldName
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

        console.error("Error in removeViewField tool:", errorMessage);

        return {
            content: [{
                type: "text",
                text: `Error removing view field: ${errorMessage}`
            }],
            isError: true
        } as IToolResult;
    }
}

export default removeViewField;
