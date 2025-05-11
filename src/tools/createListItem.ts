// src/tools/createListItem.ts
import request from 'request-promise';
import { IToolResult } from '../interfaces';
import { getSharePointHeaders, getRequestDigest } from '../auth_factory';
import { SharePointConfig } from '../config';

export interface CreateListItemParams {
    url: string;
    listTitle: string;
    itemData: Record<string, any>;
}

/**
 * Create a new item in a SharePoint list
 * @param params Parameters including site URL, list title, and item data
 * @param config SharePoint configuration
 * @returns Tool result with creation status and new item ID
 */
export async function createListItem(
    params: CreateListItemParams,
    config: SharePointConfig
): Promise<IToolResult> {
    const { url, listTitle, itemData } = params;
    console.error(`createListItem tool called with URL: ${url}, List Title: ${listTitle}`);

    try {
        // Authenticate with SharePoint
        const headers = await getSharePointHeaders(url, config);
        console.error("Headers prepared with authentication");

        // Get request digest for POST operations
        headers['X-RequestDigest'] = await getRequestDigest(url, headers);
        headers['Content-Type'] = 'application/json;odata=verbose';
        console.error("Headers prepared with request digest");

        // Encode the list title to handle special characters
        const encodedListTitle = encodeURIComponent(listTitle);
        
        // Get list item entity type name
        console.error(`Getting list schema for "${listTitle}"...`);
        const listResponse = await request({
            url: `${url}/_api/web/lists/getByTitle('${encodedListTitle}')`,
            headers: { ...headers, 'Content-Type': undefined },
            json: true,
            method: 'GET',
            timeout: 15000
        });
        
        // Prepare the item data
        const createPayload: any = {
            __metadata: { type: listResponse.d.ListItemEntityTypeFullName },
            ...itemData
        };
        
        console.error(`Creating item with payload: ${JSON.stringify(createPayload)}`);
        
        // Create the item
        const createResponse = await request({
            url: `${url}/_api/web/lists/getByTitle('${encodedListTitle}')/items`,
            headers: headers,
            json: true,
            method: 'POST',
            body: createPayload,
            timeout: 20000
        });
        
        return {
            content: [{
                type: "text",
                text: JSON.stringify({
                    success: true,
                    message: `Item successfully created in list "${listTitle}"`,
                    newItem: {
                        id: createResponse.d.ID,
                        title: createResponse.d.Title || '(No Title)',
                        created: createResponse.d.Created,
                        url: `${url}/Lists/${listTitle}/DispForm.aspx?ID=${createResponse.d.ID}`
                    }
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

        console.error("Error in createListItem tool:", errorMessage);

        return {
            content: [{
                type: "text",
                text: `Error creating list item: ${errorMessage}`
            }],
            isError: true
        } as IToolResult;
    }
}

export default createListItem;

