// src/tools/getListFields.ts
import request from 'request-promise';
import { ISharePointField, IToolResult } from '../interfaces';
import { getSharePointHeaders, getRequestDigest } from '../auth_factory';
import { SharePointConfig } from '../config';

export interface GetListFieldsParams {
    url: string;
    listTitle: string;
}

/**
 * Get detailed information about fields (columns) in a SharePoint list
 * @param params Parameters including site URL and list title
 * @param config SharePoint configuration
 * @returns Tool result with list fields data
 */
export async function getListFields(
    params: GetListFieldsParams,
    config: SharePointConfig
): Promise<IToolResult> {
    const { url, listTitle } = params;
    console.error(`getListFields tool called with URL: ${url}, List Title: ${listTitle}`);

    try {
        // Authenticate with SharePoint
        const headers = await getSharePointHeaders(url, config);
        console.error("Headers prepared with authentication");

        // Encode the list title to handle special characters
        const encodedListTitle = encodeURIComponent(listTitle);
        
        // Get fields for the specified list
        console.error(`Getting fields for list "${listTitle}"...`);
        const fieldsResponse = await request({
            url: `${url}/_api/web/lists/getByTitle('${encodedListTitle}')/fields?$filter=Hidden eq false`,
            headers: { ...headers, 'Content-Type': undefined },
            json: true,
            method: 'GET',
            timeout: 30000
        });
        
        // Process fields to show relevant information
        const formattedFields = fieldsResponse.d.results.map((field: ISharePointField) => {
            const formattedField: any = {
                InternalName: field.InternalName,
                Title: field.Title,
                Type: field.TypeAsString || field.TypeDisplayName || 'Unknown',
                ReadOnly: field.ReadOnlyField,
                Required: field.Required || false,
                Description: field.Description || '',
                Group: field.Group || ''
            };
            
            // Add choice options if available
            if (field.Choices && field.Choices.results && field.Choices.results.length > 0) {
                formattedField.Choices = field.Choices.results;
            }
            
            // Add lookup information if it's a lookup field
            if (field.TypeAsString?.toLowerCase().includes('lookup') && field.LookupList) {
                formattedField.LookupList = field.LookupList;
                formattedField.LookupField = field.LookupField;
            }
            
            // Add default value if it exists
            if (field.DefaultValue) {
                formattedField.DefaultValue = field.DefaultValue;
            }
            
            // Add custom formatter if it exists
            if (field.CustomFormatter) {
                formattedField.CustomFormatter = field.CustomFormatter;
            }
            
            return formattedField;
        });

        return {
            content: [{
                type: "text",
                text: JSON.stringify({
                    listTitle: listTitle,
                    totalFields: formattedFields.length,
                    fields: formattedFields
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

        console.error("Error in getListFields tool:", errorMessage);

        return {
            content: [{
                type: "text",
                text: `Error fetching list fields: ${errorMessage}`
            }],
            isError: true
        } as IToolResult;
    }
}

export default getListFields;

