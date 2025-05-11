// src/tools/updateListField.ts
import request from 'request-promise';
import { IToolResult, IFieldUpdateData } from '../interfaces';
import { getSharePointHeaders, getRequestDigest } from '../auth_factory';
import { SharePointConfig } from '../config';

export interface UpdateListFieldParams {
    url: string;
    listTitle: string;
    fieldInternalName: string;
    updateData: IFieldUpdateData;
}

/**
 * Update a field (column) in a SharePoint list
 * @param params Parameters including site URL, list title, field internal name, and update data
 * @param config SharePoint configuration
 * @returns Tool result with update result
 */
export async function updateListField(
    params: UpdateListFieldParams,
    config: SharePointConfig
): Promise<IToolResult> {
    const { url, listTitle, fieldInternalName, updateData } = params;
    console.error(`updateListField tool called with URL: ${url}, List Title: ${listTitle}, Field: ${fieldInternalName}`);

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

        // Encode the list title and field name to handle special characters
        const encodedListTitle = encodeURIComponent(listTitle);
        const encodedFieldName = encodeURIComponent(fieldInternalName);
        
        // First, verify the field exists
        console.error(`Verifying field "${fieldInternalName}" in list "${listTitle}"...`);
        try {
            await request({
                url: `${url}/_api/web/lists/getByTitle('${encodedListTitle}')/fields/getByInternalNameOrTitle('${encodedFieldName}')`,
                headers: { ...headers, 'Content-Type': undefined, 'X-HTTP-Method': undefined, 'IF-MATCH': undefined },
                json: true,
                method: 'GET',
                timeout: 15000
            });
        } catch (error) {
            throw new Error(`Field "${fieldInternalName}" not found in list "${listTitle}"`);
        }
        
        // Prepare the update data for the field
        const updatePayload: any = {
            __metadata: { type: 'SP.Field' }
        };
        
        // Add update properties from the updateData
        Object.entries(updateData).forEach(([key, value]) => {
            // Special handling for choices
            if (key === 'Choices' && Array.isArray(value)) {
                updatePayload.Choices = { results: value };
            } else {
                updatePayload[key] = value;
            }
        });
        
        console.error(`Updating field with payload: ${JSON.stringify(updatePayload)}`);
        
        // Update the field
        await request({
            url: `${url}/_api/web/lists/getByTitle('${encodedListTitle}')/fields/getByInternalNameOrTitle('${encodedFieldName}')`,
            headers: headers,
            json: true,
            method: 'POST',
            body: updatePayload,
            timeout: 20000
        });
        
        // Get the updated field to return its new state
        const updatedField = await request({
            url: `${url}/_api/web/lists/getByTitle('${encodedListTitle}')/fields/getByInternalNameOrTitle('${encodedFieldName}')`,
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
                    message: `Field "${fieldInternalName}" successfully updated in list "${listTitle}"`,
                    updatedField: {
                        InternalName: updatedField.d.InternalName,
                        Title: updatedField.d.Title,
                        Type: updatedField.d.TypeAsString || updatedField.d.TypeDisplayName,
                        Description: updatedField.d.Description || '',
                        Required: updatedField.d.Required || false,
                        CustomFormatter: updatedField.d.CustomFormatter || null
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

        console.error("Error in updateListField tool:", errorMessage);

        return {
            content: [{
                type: "text",
                text: `Error updating list field: ${errorMessage}`
            }],
            isError: true
        } as IToolResult;
    }
}

export default updateListField;

