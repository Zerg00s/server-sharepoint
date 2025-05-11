// src/tools/getViewFields.ts
import request from 'request-promise';
import { IToolResult } from '../interfaces';
import { getSharePointHeaders, getRequestDigest } from '../auth_factory';
import { SharePointConfig } from '../config';

export interface GetViewFieldsParams {
    url: string;
    listTitle: string;
    viewTitle: string;
}

/**
 * Get fields from a specific SharePoint list view
 * @param params Parameters including site URL, list title, and view title
 * @param config SharePoint configuration
 * @returns Tool result with view fields data
 */
export async function getViewFields(
    params: GetViewFieldsParams,
    config: SharePointConfig
): Promise<IToolResult> {
    const { url, listTitle, viewTitle } = params;
    console.error(`getViewFields tool called with URL: ${url}, List Title: ${listTitle}, View Title: ${viewTitle}`);

    try {
        // Authenticate with SharePoint
        const headers = await getSharePointHeaders(url, config);
        console.error("Headers prepared with authentication");

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
        
        // Get the view fields
        console.error(`Getting fields for view "${viewTitle}"...`);
        const viewFieldsResponse = await request({
            url: `${url}/_api/web/lists/getByTitle('${encodedListTitle}')/views/getByTitle('${encodedViewTitle}')/viewFields`,
            headers: headers,
            json: true,
            method: 'GET',
            timeout: 20000
        });
        
        const viewFields = viewFieldsResponse.d.Items.results || [];
        
        return {
            content: [{
                type: "text",
                text: JSON.stringify({
                    success: true,
                    listTitle: listTitle,
                    viewTitle: viewTitle,
                    viewId: viewDetails.d.Id,
                    fieldCount: viewFields.length,
                    fields: viewFields
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

        console.error("Error in getViewFields tool:", errorMessage);

        return {
            content: [{
                type: "text",
                text: `Error getting view fields: ${errorMessage}`
            }],
            isError: true
        } as IToolResult;
    }
}

export default getViewFields;
