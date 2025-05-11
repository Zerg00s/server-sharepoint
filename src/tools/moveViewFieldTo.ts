// src/tools/moveViewFieldTo.ts
import request from 'request-promise';
import { IToolResult } from '../interfaces';
import { getSharePointHeaders, getRequestDigest } from '../auth';
import { SharePointConfig } from '../config';

export interface MoveViewFieldToParams {
    url: string;
    listTitle: string;
    viewTitle: string;
    fieldName: string; // Internal name of the field to move
    index: number; // New position index (0-based)
}

/**
 * Move a field to a specific position in a SharePoint list view
 * @param params Parameters including site URL, list title, view title, field name, and position index
 * @param config SharePoint configuration
 * @returns Tool result with move status
 */
export async function moveViewFieldTo(
    params: MoveViewFieldToParams,
    config: SharePointConfig
): Promise<IToolResult> {
    const { url, listTitle, viewTitle, fieldName, index } = params;
    console.error(`moveViewFieldTo tool called with URL: ${url}, List Title: ${listTitle}, View Title: ${viewTitle}, Field: ${fieldName}, Index: ${index}`);

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
        
        // Get the current view fields to verify field exists
        console.error(`Verifying field "${fieldName}" exists in view "${viewTitle}"...`);
        let currentFields;
        try {
            const viewFieldsResponse = await request({
                url: `${url}/_api/web/lists/getByTitle('${encodedListTitle}')/views/getByTitle('${encodedViewTitle}')/viewFields`,
                headers: { ...headers, 'Content-Type': undefined },
                json: true,
                method: 'GET',
                timeout: 20000
            });
            
            currentFields = viewFieldsResponse.d.Items.results || [];
            if (!currentFields.includes(fieldName)) {
                throw new Error(`Field "${fieldName}" not found in view "${viewTitle}"`);
            }
        } catch (error) {
            if (error instanceof Error && error.message.includes("not found in view")) {
                throw error; // Re-throw our custom error
            }
            throw new Error(`Error verifying field in view: ${error instanceof Error ? error.message : String(error)}`);
        }
        
        // Move the field to the specified position
        console.error(`Moving field "${fieldName}" to position ${index} in view "${viewTitle}"...`);
        await request({
            url: `${url}/_api/web/lists/getByTitle('${encodedListTitle}')/views/getByTitle('${encodedViewTitle}')/viewFields/moveViewFieldTo`,
            headers: headers,
            json: true,
            method: 'POST',
            body: {
                "field": fieldName,
                "index": index
            },
            timeout: 20000
        });
        
        return {
            content: [{
                type: "text",
                text: JSON.stringify({
                    success: true,
                    message: `Field "${fieldName}" successfully moved to position ${index} in view "${viewTitle}" in list "${listTitle}"`,
                    listTitle: listTitle,
                    viewTitle: viewTitle,
                    viewId: viewDetails.d.Id,
                    movedField: fieldName,
                    newPosition: index
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

        console.error("Error in moveViewFieldTo tool:", errorMessage);

        return {
            content: [{
                type: "text",
                text: `Error moving view field: ${errorMessage}`
            }],
            isError: true
        } as IToolResult;
    }
}

export default moveViewFieldTo;