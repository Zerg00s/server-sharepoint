// src/tools/batchUpdateListItems.ts
import axios from 'axios';
import { IToolResult } from '../interfaces';
import { getSharePointHeaders } from '../auth_factory';
import { SharePointConfig } from '../config';

export interface BatchUpdateItem {
    id: number;
    data: Record<string, any>;
}

export interface BatchUpdateListItemsParams {
    url: string;
    listTitle: string;
    items: BatchUpdateItem[];
}

/**
 * Update multiple items in a SharePoint list using individual requests
 * @param params Parameters including site URL, list title, and array of items with IDs and data to update
 * @param config SharePoint configuration
 * @returns Tool result with update status
 */
export async function batchUpdateListItems(
    params: BatchUpdateListItemsParams,
    config: SharePointConfig
): Promise<IToolResult> {
    const { url, listTitle, items } = params;
    console.error(`batchUpdateListItems tool called with URL: ${url}, List Title: ${listTitle}, Items Count: ${items.length}`);

    if (!items || items.length === 0) {
        return {
            content: [{
                type: "text",
                text: "Error: No items provided for batch update"
            }],
            isError: true
        } as IToolResult;
    }

    try {
        // Authenticate with SharePoint
        const headers = await getSharePointHeaders(url, config);
        console.error("Headers prepared with authentication");

        // Get list metadata to find the entity type name
        console.error(`Getting list schema for "${listTitle}"...`);
        const listUrl = `${url}/_api/web/lists/getByTitle('${encodeURIComponent(listTitle)}')`;
        
        const listResponse = await axios({
            url: listUrl,
            headers: { 
                ...headers, 
                'Accept': 'application/json;odata=verbose' 
            },
            method: 'GET',
            timeout: 15000
        });
        
        // Get entity type name from list response
        const entityTypeFullName = listResponse.data.d.ListItemEntityTypeFullName;
        console.error(`Entity type name: ${entityTypeFullName}`);
        
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
        
        // Update items using individual requests (not a true batch but more reliable)
        console.error(`Updating ${items.length} items individually...`);
        
        const updatedItems = [];
        for (const item of items) {
            const updateUrl = `${url}/_api/web/lists/getByTitle('${encodeURIComponent(listTitle)}')/items(${item.id})`;
            
            // Prepare the item data with metadata
            const updatePayload = {
                __metadata: { type: entityTypeFullName },
                ...item.data
            };
            
            try {
                await axios({
                    url: updateUrl,
                    method: 'POST',
                    headers: {
                        ...headers,
                        'Accept': 'application/json;odata=verbose',
                        'Content-Type': 'application/json;odata=verbose',
                        'X-RequestDigest': requestDigest,
                        'IF-MATCH': '*',
                        'X-HTTP-Method': 'MERGE'
                    },
                    data: updatePayload,
                    timeout: 20000
                });
                
                console.error(`Updated item ${item.id} successfully`);
                
                // Get the updated item
                const getItemResponse = await axios({
                    url: updateUrl,
                    method: 'GET',
                    headers: {
                        ...headers,
                        'Accept': 'application/json;odata=verbose'
                    },
                    timeout: 15000
                });
                
                updatedItems.push({
                    id: getItemResponse.data.d.ID,
                    title: getItemResponse.data.d.Title || '(No Title)',
                    modified: getItemResponse.data.d.Modified
                });
            } catch (itemError) {
                console.error(`Error updating item ${item.id}:`, itemError);
                // Continue with other items
            }
        }
        
        return {
            content: [{
                type: "text",
                text: JSON.stringify({
                    success: true,
                    message: `Successfully updated ${updatedItems.length} out of ${items.length} items in list "${listTitle}"`,
                    updatedItems: updatedItems
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

        console.error("Error in batchUpdateListItems tool:", errorMessage);

        return {
            content: [{
                type: "text",
                text: `Error updating items in batch: ${errorMessage}`
            }],
            isError: true
        } as IToolResult;
    }
}

export default batchUpdateListItems;
