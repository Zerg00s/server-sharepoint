// src/tools/deleteListItem.ts
import request from 'request-promise';
import { IToolResult } from '../interfaces';
import { getSharePointHeaders, getRequestDigest } from '../auth_factory';
import { SharePointConfig } from '../config';

export interface DeleteListItemParams {
    url: string;
    listTitle: string;
    itemId: number;
}

/**
 * Delete an item from a SharePoint list
 * @param params Parameters including site URL, list title, and item ID
 * @param config SharePoint configuration
 * @returns Tool result with deletion status
 */
export async function deleteListItem(
    params: DeleteListItemParams,
    config: SharePointConfig
): Promise<IToolResult> {
    const { url, listTitle, itemId } = params;
    console.error(`deleteListItem tool called with URL: ${url}, List Title: ${listTitle}, Item ID: ${itemId}`);

    try {
        // Authenticate with SharePoint
        const headers = await getSharePointHeaders(url, config);
        console.error("Headers prepared with authentication");

        // Get request digest for DELETE operations
        headers['X-RequestDigest'] = await getRequestDigest(url, headers);
        headers['IF-MATCH'] = '*';
        headers['X-HTTP-Method'] = 'DELETE';
        console.error("Headers prepared with request digest for delete operation");

        // Encode the list title to handle special characters
        const encodedListTitle = encodeURIComponent(listTitle);
        
        // First, verify the item exists
        console.error(`Verifying item ID ${itemId} exists...`);
        let itemExists = true;
        let itemDetails = null;
        
        try {
            const itemResponse = await request({
                url: `${url}/_api/web/lists/getByTitle('${encodedListTitle}')/items(${itemId})`,
                headers: { ...headers, 'X-HTTP-Method': undefined, 'IF-MATCH': undefined },
                json: true,
                method: 'GET',
                timeout: 15000
            });
            
            itemDetails = {
                id: itemResponse.d.ID,
                title: itemResponse.d.Title || '(No Title)',
                created: itemResponse.d.Created
            };
        } catch (error) {
            itemExists = false;
            throw new Error(`Item with ID ${itemId} not found in list "${listTitle}"`);
        }
        
        // Delete the item
        console.error(`Deleting item ID ${itemId}...`);
        await request({
            url: `${url}/_api/web/lists/getByTitle('${encodedListTitle}')/items(${itemId})`,
            headers: headers,
            json: true,
            method: 'POST',
            body: '',
            timeout: 20000
        });
        
        return {
            content: [{
                type: "text",
                text: JSON.stringify({
                    success: true,
                    message: `Item with ID ${itemId} successfully deleted from list "${listTitle}"`,
                    deletedItem: itemDetails
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

        console.error("Error in deleteListItem tool:", errorMessage);

        return {
            content: [{
                type: "text",
                text: `Error deleting list item: ${errorMessage}`
            }],
            isError: true
        } as IToolResult;
    }
}

export default deleteListItem;

