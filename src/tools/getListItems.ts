// src/tools/getListItems.ts
import request from 'request-promise';
import { ISharePointListItem, IFormattedListItem, IToolResult } from '../interfaces';
import { getSharePointHeaders, getRequestDigest } from '../auth_factory';
import { SharePointConfig } from '../config';

export interface GetListItemsParams {
    url: string;
    listTitle: string;
}

/**
 * Get all items from a specific SharePoint list
 * @param params Parameters including site URL and list title
 * @param config SharePoint configuration
 * @returns Tool result with list items data
 */
export async function getListItems(
    params: GetListItemsParams,

    config: SharePointConfig
): Promise<IToolResult> {
    const { url, listTitle } = params;
    console.error(`getListItems tool called with URL: ${url}, List Title: ${listTitle}`);

    try {
        // Authenticate with SharePoint
        const headers = await getSharePointHeaders(url, config);
        console.error("Headers prepared:", headers);

        // Encode the list title to handle special characters
        const encodedListTitle = encodeURIComponent(listTitle);
        
        // First, get the list to validate it exists
        console.error(`Getting list details for "${listTitle}"...`);
        const listResponse = await request({
            url: `${url}/_api/web/lists/getByTitle('${encodedListTitle}')`,
            headers: headers,
            json: true,
            method: 'GET',
            timeout: 10000
        });
        
        console.error(`List found: ${listResponse.d.Title}, ID: ${listResponse.d.Id}`);
        
        // Now get all items from the list
        console.error(`Retrieving items from list "${listTitle}"...`);
        const itemsResponse = await request({
            url: `${url}/_api/web/lists/getByTitle('${encodedListTitle}')/items?$top=5000`,
            headers: headers,
            json: true,
            method: 'GET',
            timeout: 20000
        });
        
        const items: ISharePointListItem[] = itemsResponse.d.results;
        console.error(`Retrieved ${items.length} items from list "${listTitle}"`);
        
        // Extract field names by looking at the first item (if exists)
        let fieldNames: string[] = [];
        if (items.length > 0) {
            fieldNames = Object.keys(items[0])
                .filter(key => !key.startsWith('__') && 
                               !['AttachmentFiles', 'Attachments', 'FirstUniqueAncestorSecurableObject',
                                 'RoleAssignments', 'ContentType', 'FieldValuesAsHtml', 'FieldValuesAsText', 
                                 'FieldValuesForEdit', 'File', 'Folder', 'ParentList'].includes(key));
        }
        
        console.error(`Fields available: ${fieldNames.join(', ')}`);
        
        // Format items for nicer display - only include relevant fields
        const formattedItems: IFormattedListItem[] = items.map((item: ISharePointListItem) => {
            const formattedItem: IFormattedListItem = {};
            fieldNames.forEach(field => {
                formattedItem[field] = item[field];
            });
            return formattedItem;
        });

        return {
            content: [{
                type: "text",
                text: JSON.stringify({
                    listTitle: listTitle,
                    totalItems: items.length,
                    fields: fieldNames,
                    items: formattedItems
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

        console.error("Error in getListItems tool:", errorMessage);

        return {
            content: [{
                type: "text",
                text: `Error fetching list items: ${errorMessage}`
            }],
            isError: true
        } as IToolResult;
    }
}

export default getListItems;
