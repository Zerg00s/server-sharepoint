// src/tools/batchCreateListItems.ts
import axios from 'axios';
import { IToolResult } from '../interfaces';
import { getSharePointHeaders } from '../auth_factory';
import { SharePointConfig } from '../config';

export interface BatchCreateListItemsParams {
    url: string;
    listTitle: string;
    items: Record<string, any>[];
}

/**
 * Create multiple items in a SharePoint list using individual requests
 * @param params Parameters including site URL, list title, and array of item data objects
 * @param config SharePoint configuration
 * @returns Tool result with creation status and new item IDs
 */
export async function batchCreateListItems(
    params: BatchCreateListItemsParams,
    config: SharePointConfig
): Promise<IToolResult> {
    const { url, listTitle, items } = params;
    console.error(`batchCreateListItems tool called with URL: ${url}, List Title: ${listTitle}, Items Count: ${items.length}`);

    if (!items || items.length === 0) {
        return {
            content: [{
                type: "text",
                text: "Error: No items provided for batch creation"
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
        
        // Create items using individual requests (not a true batch but more reliable)
        console.error(`Creating ${items.length} items individually...`);
        
        const createdItems = [];
        for (const item of items) {
            const createUrl = `${url}/_api/web/lists/getByTitle('${encodeURIComponent(listTitle)}')/items`;
            
            // Prepare the item data with metadata
            const createPayload = {
                __metadata: { type: entityTypeFullName },
                ...item
            };
            
            try {
                const createResponse = await axios({
                    url: createUrl,
                    method: 'POST',
                    headers: {
                        ...headers,
                        'Accept': 'application/json;odata=verbose',
                        'Content-Type': 'application/json;odata=verbose',
                        'X-RequestDigest': requestDigest
                    },
                    data: createPayload,
                    timeout: 20000
                });
                
                console.error(`Created item ${createResponse.data.d.ID} successfully`);
                createdItems.push({
                    id: createResponse.data.d.ID,
                    title: createResponse.data.d.Title || '(No Title)',
                    created: createResponse.data.d.Created
                });
            } catch (itemError) {
                console.error(`Error creating item:`, itemError);
                // Continue with other items
            }
        }
        
        return {
            content: [{
                type: "text",
                text: JSON.stringify({
                    success: true,
                    message: `Successfully created ${createdItems.length} out of ${items.length} items in list "${listTitle}"`,
                    createdItems: createdItems
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

        console.error("Error in batchCreateListItems tool:", errorMessage);

        return {
            content: [{
                type: "text",
                text: `Error creating items in batch: ${errorMessage}`
            }],
            isError: true
        } as IToolResult;
    }
}

export default batchCreateListItems;
