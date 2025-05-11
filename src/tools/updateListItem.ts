// src/tools/updateListItem.ts
import request from 'request-promise';
import { IToolResult } from '../interfaces';
import { getSharePointHeaders, getRequestDigest } from '../auth';
import { SharePointConfig } from '../config';

export interface UpdateListItemParams {
    url: string;
    listTitle: string;
    itemId: number;
    itemData: Record<string, any>;
}

/**
 * Update an item in a SharePoint list
 * @param params Parameters including site URL, list title, item ID, and item data to update
 * @param config SharePoint configuration
 * @returns Tool result with update status
 */
export async function updateListItem(
    params: UpdateListItemParams,
    config: SharePointConfig
): Promise<IToolResult> {
    const { url, listTitle, itemId, itemData } = params;
    console.error(`updateListItem tool called with URL: ${url}, List Title: ${listTitle}, Item ID: ${itemId}`);

    try {
        // Authenticate with SharePoint
        const headers = await getSharePointHeaders(url, config);
        console.error("Headers prepared with authentication");

        // Get request digest for POST operations
        headers['X-RequestDigest'] = await getRequestDigest(url, headers);
        headers['Content-Type'] = 'application/json;odata=verbose';
        headers['IF-MATCH'] = '*';
        headers['X-HTTP-Method'] = 'MERGE';
        console.error("Headers prepared with request digest for update operation");

        // Encode the list title to handle special characters
        const encodedListTitle = encodeURIComponent(listTitle);
        
        // Get list item entity type name
        console.error(`Getting list schema for "${listTitle}"...`);
        const listResponse = await request({
            url: `${url}/_api/web/lists/getByTitle('${encodedListTitle}')`,
            headers: { ...headers, 'Content-Type': undefined, 'X-HTTP-Method': undefined, 'IF-MATCH': undefined },
            json: true,
            method: 'GET',
            timeout: 15000
        });
        
        // Verify the item exists
        console.error(`Verifying item ID ${itemId} exists...`);
        try {
            await request({
                url: `${url}/_api/web/lists/getByTitle('${encodedListTitle}')/items(${itemId})`,
                headers: { ...headers, 'Content-Type': undefined, 'X-HTTP-Method': undefined, 'IF-MATCH': undefined },
                json: true,
                method: 'GET',
                timeout: 15000
            });
        } catch (error) {
            throw new Error(`Item with ID ${itemId} not found in list "${listTitle}"`);
        }
        
        // Prepare the update data
        const updatePayload: any = {
            __metadata: { type: listResponse.d.ListItemEntityTypeFullName },
            ...itemData
        };
        
        console.error(`Updating item with payload: ${JSON.stringify(updatePayload)}`);
        
        // Update the item
        await request({
            url: `${url}/_api/web/lists/getByTitle('${encodedListTitle}')/items(${itemId})`,
            headers: headers,
            json: true,
            method: 'POST',
            body: updatePayload,
            timeout: 20000
        });
        
        // Get the updated item to return its new state
        const updatedItem = await request({
            url: `${url}/_api/web/lists/getByTitle('${encodedListTitle}')/items(${itemId})`,
            headers: { ...headers, 'Content-Type': undefined, 'X-HTTP-Method': undefined, 'IF-MATCH': undefined },
            json: true,
            method: 'GET',
            timeout: 15000
        });

        return {
            content: [{
                type: "text",
                text: JSON.stringify({
                    success: true,
                    message: `Item with ID ${itemId} successfully updated in list "${listTitle}"`,
                    updatedItem: updatedItem.d
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

        console.error("Error in updateListItem tool:", errorMessage);

        return {
            content: [{
                type: "text",
                text: `Error updating list item: ${errorMessage}`
            }],
            isError: true
        } as IToolResult;
    }
}

export default updateListItem;
