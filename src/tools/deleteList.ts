// src/tools/deleteList.ts
import request from 'request-promise';
import { IToolResult } from '../interfaces';
import { getSharePointHeaders, getRequestDigest } from '../auth';
import { SharePointConfig } from '../config';

export interface DeleteListParams {
    url: string;
    listTitle: string;
    confirmation?: string; // Optional confirmation string to prevent accidental deletion
}

/**
 * Delete a SharePoint list
 * @param params Parameters including site URL and list title
 * @param config SharePoint configuration
 * @returns Tool result with deletion status
 */
export async function deleteList(
    params: DeleteListParams,
    config: SharePointConfig
): Promise<IToolResult> {
    const { url, listTitle, confirmation } = params;
    console.error(`deleteList tool called with URL: ${url}, List Title: ${listTitle}`);

    try {
        // Check confirmation string to prevent accidental deletion
        if (!confirmation || confirmation !== listTitle) {
            throw new Error(`To delete the list, please provide a confirmation parameter that matches exactly the list title '${listTitle}'`);
        }

        // Authenticate with SharePoint
        const headers = await getSharePointHeaders(url, config);
        console.error("Headers prepared with authentication");

        // Get request digest for DELETE operations
        headers['X-RequestDigest'] = await getRequestDigest(url, headers);
        headers['X-HTTP-Method'] = 'DELETE';
        headers['IF-MATCH'] = '*';
        console.error("Headers prepared with request digest for delete operation");

        // Encode the list title to handle special characters
        const encodedListTitle = encodeURIComponent(listTitle);
        
        // First, verify the list exists and get its details
        console.error(`Verifying list "${listTitle}" exists...`);
        let listDetails;
        try {
            listDetails = await request({
                url: `${url}/_api/web/lists/getByTitle('${encodedListTitle}')`,
                headers: { ...headers, 'X-HTTP-Method': undefined, 'IF-MATCH': undefined },
                json: true,
                method: 'GET',
                timeout: 15000
            });
        } catch (error) {
            throw new Error(`List "${listTitle}" not found`);
        }
        
        // Store the list details for response
        const listInfo = {
            id: listDetails.d.Id,
            title: listDetails.d.Title,
            itemCount: listDetails.d.ItemCount,
            templateType: listDetails.d.BaseTemplate
        };
        
        // Delete the list
        console.error(`Deleting list "${listTitle}" with ID ${listInfo.id}...`);
        await request({
            url: `${url}/_api/web/lists/getByTitle('${encodedListTitle}')`,
            headers: headers,
            method: 'POST', // POST with DELETE X-HTTP-Method header
            body: '',
            timeout: 30000
        });
        
        return {
            content: [{
                type: "text",
                text: JSON.stringify({
                    success: true,
                    message: `List "${listTitle}" successfully deleted from site`,
                    deletedList: listInfo
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

        console.error("Error in deleteList tool:", errorMessage);

        return {
            content: [{
                type: "text",
                text: `Error deleting list: ${errorMessage}`
            }],
            isError: true
        } as IToolResult;
    }
}

export default deleteList;
