// src/tools/updateList.ts
import request from 'request-promise';
import { IToolResult } from '../interfaces';
import { getSharePointHeaders, getRequestDigest } from '../auth_factory';
import { SharePointConfig } from '../config';

export interface UpdateListParams {
    url: string;
    listTitle: string;
    updateData: {
        Title?: string;
        Description?: string;
        EnableVersioning?: boolean;
        EnableMinorVersions?: boolean;
        EnableModeration?: boolean;
        DraftVersionVisibility?: number;   // 0=Reader, 1=Author, 2=Approver
        ContentTypesEnabled?: boolean;
        Hidden?: boolean;
        Ordered?: boolean;
    };
}

/**
 * Update an existing SharePoint list's properties
 * @param params Parameters including site URL, list title, and update data
 * @param config SharePoint configuration
 * @returns Tool result with update status
 */
export async function updateList(
    params: UpdateListParams,
    config: SharePointConfig
): Promise<IToolResult> {
    const { url, listTitle, updateData } = params;
    console.error(`updateList tool called with URL: ${url}, List Title: ${listTitle}`);

    try {
        // Authenticate with SharePoint
        const headers = await getSharePointHeaders(url, config);
        console.error("Headers prepared with authentication");

        // Get request digest for POST operations
        const requestDigest = await getRequestDigest(url, headers);
        console.error("Request digest obtained");

        // Encode the list title to handle special characters
        const encodedListTitle = encodeURIComponent(listTitle);
        
        // First, get the list to ensure it exists and to get its entity type
        console.error(`Getting list schema for "${listTitle}"...`);
        const listResponse = await request({
            url: `${url}/_api/web/lists/getByTitle('${encodedListTitle}')`,
            headers: { ...headers },
            json: true,
            method: 'GET',
            timeout: 15000
        });
        
        const entityTypeFullName = listResponse.d.__metadata.type;
        console.error(`List found with entity type: ${entityTypeFullName}`);
        
        // Prepare the update data with metadata
        const updatePayload = {
            __metadata: { type: entityTypeFullName },
            ...updateData
        };
        
        // Set headers for update
        const updateHeaders = {
            ...headers,
            'Content-Type': 'application/json;odata=verbose',
            'Accept': 'application/json;odata=verbose',
            'X-RequestDigest': requestDigest,
            'X-HTTP-Method': 'MERGE',
            'IF-MATCH': '*'  // Optimistic concurrency
        };
        
        // Send the update request
        console.error("Sending update request...");
        await request({
            url: `${url}/_api/web/lists/getByTitle('${encodedListTitle}')`,
            headers: updateHeaders,
            body: updatePayload,
            json: true,
            method: 'POST',
            timeout: 30000
        });
        
        console.error(`List "${listTitle}" updated successfully.`);
        
        // Get the updated list to return the new properties
        const updatedList = await request({
            url: `${url}/_api/web/lists/getByTitle('${encodedListTitle}')`,
            headers: { ...headers },
            json: true,
            method: 'GET',
            timeout: 15000
        });
        
        // Extract the relevant properties for the response
        const listProps = {
            title: updatedList.d.Title,
            description: updatedList.d.Description,
            itemCount: updatedList.d.ItemCount,
            lastModified: updatedList.d.LastItemModifiedDate,
            enableVersioning: updatedList.d.EnableVersioning,
            enableMinorVersions: updatedList.d.EnableMinorVersions,
            enableModeration: updatedList.d.EnableModeration,
            draftVersionVisibility: updatedList.d.DraftVersionVisibility,
            contentTypesEnabled: updatedList.d.ContentTypesEnabled,
            hidden: updatedList.d.Hidden,
            ordered: updatedList.d.Ordered,
            created: updatedList.d.Created
        };
        
        return {
            content: [{
                type: "text",
                text: JSON.stringify({
                    success: true,
                    message: `List "${listTitle}" updated successfully`,
                    list: listProps
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

        console.error("Error in updateList tool:", errorMessage);

        return {
            content: [{
                type: "text",
                text: `Error updating list "${listTitle}": ${errorMessage}`
            }],
            isError: true
        } as IToolResult;
    }
}

export default updateList;
