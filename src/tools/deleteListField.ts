// src/tools/deleteListField.ts
import request from 'request-promise';
import { IToolResult } from '../interfaces';
import { getSharePointHeaders, getRequestDigest } from '../auth';
import { SharePointConfig } from '../config';

export interface DeleteListFieldParams {
    url: string;
    listTitle: string;
    fieldInternalName: string;
    confirmation?: string; // Optional confirmation string to prevent accidental deletion
}

/**
 * Delete a field (column) from a SharePoint list
 * @param params Parameters including site URL, list title, and field internal name
 * @param config SharePoint configuration
 * @returns Tool result with deletion status
 */
export async function deleteListField(
    params: DeleteListFieldParams,
    config: SharePointConfig
): Promise<IToolResult> {
    const { url, listTitle, fieldInternalName, confirmation } = params;
    console.error(`deleteListField tool called with URL: ${url}, List Title: ${listTitle}, Field: ${fieldInternalName}`);

    try {
        // Check confirmation string to prevent accidental deletion of important fields
        if (!confirmation || confirmation !== fieldInternalName) {
            throw new Error(`To delete the field, please provide a confirmation parameter that matches exactly the field internal name '${fieldInternalName}'`);
        }

        // Authenticate with SharePoint
        const headers = await getSharePointHeaders(url, config);
        console.error("Headers prepared with authentication");

        // Get request digest for DELETE operations
        headers['X-RequestDigest'] = await getRequestDigest(url, headers);
        headers['X-HTTP-Method'] = 'DELETE';
        headers['IF-MATCH'] = '*';
        console.error("Headers prepared with request digest for delete operation");

        // Encode the list title and field name to handle special characters
        const encodedListTitle = encodeURIComponent(listTitle);
        const encodedFieldName = encodeURIComponent(fieldInternalName);
        
        // First, verify the field exists and get its details
        console.error(`Verifying field "${fieldInternalName}" exists...`);
        let fieldDetails;
        try {
            fieldDetails = await request({
                url: `${url}/_api/web/lists/getByTitle('${encodedListTitle}')/fields/getByInternalNameOrTitle('${encodedFieldName}')`,
                headers: { ...headers, 'X-HTTP-Method': undefined, 'IF-MATCH': undefined },
                json: true,
                method: 'GET',
                timeout: 15000
            });
        } catch (error) {
            throw new Error(`Field "${fieldInternalName}" not found in list "${listTitle}"`);
        }
        
        // Check if it's a system field that shouldn't be deleted
        if (fieldDetails.d.ReadOnlyField) {
            throw new Error(`Field "${fieldInternalName}" is a read-only field and cannot be deleted`);
        }
        
        // Store the field details for response
        const fieldInfo = {
            internalName: fieldDetails.d.InternalName,
            title: fieldDetails.d.Title,
            type: fieldDetails.d.TypeAsString || fieldDetails.d.TypeDisplayName
        };
        
        // Delete the field
        console.error(`Deleting field "${fieldInternalName}"...`);
        await request({
            url: `${url}/_api/web/lists/getByTitle('${encodedListTitle}')/fields/getByInternalNameOrTitle('${encodedFieldName}')`,
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
                    message: `Field "${fieldInternalName}" successfully deleted from list "${listTitle}"`,
                    deletedField: fieldInfo
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

        console.error("Error in deleteListField tool:", errorMessage);

        return {
            content: [{
                type: "text",
                text: `Error deleting list field: ${errorMessage}`
            }],
            isError: true
        } as IToolResult;
    }
}

export default deleteListField;
