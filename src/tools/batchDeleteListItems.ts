// src/tools/batchDeleteListItems.ts
import axios from 'axios';
import { IToolResult } from '../interfaces';
import { getSharePointHeaders } from '../auth_factory';
import { SharePointConfig } from '../config';

export interface BatchDeleteListItemsParams {
    url: string;
    listTitle: string;
    itemIds: number[];
}

/**
 * Delete multiple items from a SharePoint list using individual requests
 * @param params Parameters including site URL, list title, and array of item IDs to delete
 * @param config SharePoint configuration
 * @returns Tool result with deletion status
 */
export async function batchDeleteListItems(
    params: BatchDeleteListItemsParams,
    config: SharePointConfig
): Promise<IToolResult> {
    const { url, listTitle, itemIds } = params;
    console.error(`batchDeleteListItems tool called with URL: ${url}, List Title: ${listTitle}, Items Count: ${itemIds.length}`);

    if (!itemIds || itemIds.length === 0) {
        return {
            content: [{
                type: "text",
                text: "Error: No item IDs provided for batch deletion"
            }],
            isError: true
        } as IToolResult;
    }

    try {
        // Authenticate with SharePoint
        const headers = await getSharePointHeaders(url, config);
        console.error("Headers prepared with authentication");
        
        // Get request digest for POST operations
        console.error("Getting request digest...");
        const digestUrl = `${url}/_api/contextinfo`;
        
        const digestResponse = await axios({
            url: digestUrl,
            method: 'POST',
            headers: { 
                ...headers, 
                'Accept': 'application/json;odata=verbose' 
            },
            timeout: 15000
        });
        
        const requestDigest = digestResponse.data.d.GetContextWebInformation.FormDigestValue;
        console.error("Request digest obtained successfully");
        
        // First get details of items to be deleted
        console.error(`Fetching details of ${itemIds.length} items to be deleted...`);
        
        const itemsToDelete = [];
        for (const itemId of itemIds) {
            try {
                const getItemUrl = `${url}/_api/web/lists/getByTitle('${encodeURIComponent(listTitle)}')/items(${itemId})`;
                
                const itemResponse = await axios({
                    url: getItemUrl,
                    method: 'GET',
                    headers: {
                        ...headers,
                        'Accept': 'application/json;odata=verbose'
                    },
                    timeout: 15000
                });
                
                itemsToDelete.push({
                    id: itemResponse.data.d.ID,
                    title: itemResponse.data.d.Title || '(No Title)',
                    created: itemResponse.data.d.Created
                });
            } catch (error) {
                console.error(`Error getting item ${itemId}: Item may not exist`);
                // Continue with other items
            }
        }
        
        if (itemsToDelete.length === 0) {
            return {
                content: [{
                    type: "text",
                    text: `Error: None of the specified items were found in list "${listTitle}"`
                }],
                isError: true
            } as IToolResult;
        }
        
        // Delete items using individual requests (not a true batch but more reliable)
        console.error(`Deleting ${itemsToDelete.length} items individually...`);
        
        const deletedItems = [];
        for (const item of itemsToDelete) {
            const deleteUrl = `${url}/_api/web/lists/getByTitle('${encodeURIComponent(listTitle)}')/items(${item.id})`;
            
            try {
                await axios({
                    url: deleteUrl,
                    method: 'POST',
                    headers: {
                        ...headers,
                        'Accept': 'application/json;odata=verbose',
                        'X-RequestDigest': requestDigest,
                        'IF-MATCH': '*',
                        'X-HTTP-Method': 'DELETE'
                    },
                    timeout: 20000
                });
                
                console.error(`Deleted item ${item.id} successfully`);
                deletedItems.push(item);
            } catch (itemError) {
                console.error(`Error deleting item ${item.id}:`, itemError);
                // Continue with other items
            }
        }
        
        return {
            content: [{
                type: "text",
                text: JSON.stringify({
                    success: true,
                    message: `Successfully deleted ${deletedItems.length} out of ${itemsToDelete.length} items from list "${listTitle}"`,
                    deletedItems: deletedItems
                }, null, 2)
            }]
        } as IToolResult;
    } catch (error: unknown) {
        // Type-safe error handling
        let errorMessage: string;

        if (error instanceof Error) {
            errorMessage = error.message;
            console.error(error.stack);
        } else if (typeof error === 'string') {
            errorMessage = error;
        } else {
            errorMessage = "Unknown error occurred";
        }

        console.error("Error in batchDeleteListItems tool:", errorMessage);

        return {
            content: [{
                type: "text",
                text: `Error deleting items in batch: ${errorMessage}`
            }],
            isError: true
        } as IToolResult;
    }
}

export default batchDeleteListItems;
