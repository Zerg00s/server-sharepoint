// src/tools/createListView.ts
import request from 'request-promise';
import { IToolResult, ICreateListViewData, ISharePointField } from '../interfaces';
import { getSharePointHeaders, getRequestDigest } from '../auth_factory';
import { SharePointConfig } from '../config';

export interface CreateListViewParams {
    url: string;
    listTitle: string;
    viewData: ICreateListViewData;
}

/**
 * Create a new view for a SharePoint list
 * @param params Parameters including site URL, list title, and view data
 * @param config SharePoint configuration
 * @returns Tool result with creation status and new view info
 */
export async function createListView(
    params: CreateListViewParams,
    config: SharePointConfig
): Promise<IToolResult> {
    const { url, listTitle, viewData } = params;
    console.error(`createListView tool called with URL: ${url}, List Title: ${listTitle}, View: ${viewData.Title}`);

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
        
        // First, get the list to validate it exists
        console.error(`Getting list details for "${listTitle}"...`);
        const listResponse = await request({
            url: `${url}/_api/web/lists/getByTitle('${encodedListTitle}')`,
            headers: { ...headers, 'Content-Type': undefined },
            json: true,
            method: 'GET',
            timeout: 15000
        });
        
        // Get fields to validate that the view fields exist
        if (viewData.ViewFields && viewData.ViewFields.length > 0) {
            console.error(`Validating view fields...`);
            const fieldsResponse = await request({
                url: `${url}/_api/web/lists/getByTitle('${encodedListTitle}')/fields?$select=InternalName,Title`,
                headers: { ...headers, 'Content-Type': undefined },
                json: true,
                method: 'GET',
                timeout: 30000
            });
            
            const availableFields = fieldsResponse.d.results.map((field: ISharePointField) => field.InternalName);
            
            // Check if all requested view fields exist in the list
            const invalidFields = viewData.ViewFields.filter((field: string) => !availableFields.includes(field));
            
            if (invalidFields.length > 0) {
                throw new Error(`Some view fields do not exist in the list: ${invalidFields.join(', ')}`);
            }
        }
        
        // Use the ViewCreationInformation approach as recommended in SharePoint REST API docs
        const viewCreationPayload: any = {
            "parameters": {
                "__metadata": { "type": "SP.ViewCreationInformation" },
                "Title": viewData.Title,
                "RowLimit": viewData.RowLimit || 30,
                "PersonalView": viewData.PersonalView || false
            }
        };
        
        // Add ViewFields if provided
        if (viewData.ViewFields && viewData.ViewFields.length > 0) {
            viewCreationPayload.parameters["ViewFields"] = {
                "__metadata": {
                    "type": "Collection(Edm.String)"
                },
                "results": viewData.ViewFields
            };
        }
        
        // Add ViewQuery (CAML query) if provided
        if (viewData.ViewQuery) {
            viewCreationPayload.parameters["Query"] = viewData.ViewQuery;
        }

        // Create the view using the Add method
        console.error(`Creating view using Add method with payload: ${JSON.stringify(viewCreationPayload)}`);
        
        const createResponse = await request({
            url: `${url}/_api/web/lists/getByTitle('${encodedListTitle}')/Views/Add`,
            headers: headers,
            json: true,
            method: 'POST',
            body: viewCreationPayload,
            timeout: 30000
        });
        
        // Set as default view if requested (direct approach)
        if (viewData.SetAsDefaultView) {
            console.error(`Setting view "${viewData.Title}" as default...`);
            
            try {
                // Update the DefaultView property on the view itself
                const defaultViewPayload = {
                    __metadata: { type: 'SP.View' },
                    DefaultView: true
                };
                
                await request({
                    url: `${url}/_api/web/lists/getByTitle('${encodedListTitle}')/views/getByTitle('${viewData.Title}')`,
                    headers: { 
                        ...headers, 
                        'X-HTTP-Method': 'MERGE',
                        'IF-MATCH': '*'
                    },
                    json: true,
                    method: 'POST',
                    body: defaultViewPayload,
                    timeout: 20000
                });
                
                console.error(`View "${viewData.Title}" set as default successfully`);
            } catch (setDefaultError) {
                console.error(`Warning: Could not set view as default: ${setDefaultError instanceof Error ? setDefaultError.message : String(setDefaultError)}`);
                // Continue execution, don't throw an error as the view was still created successfully
            }
        }
        
        return {
            content: [{
                type: "text",
                text: JSON.stringify({
                    success: true,
                    message: `View "${viewData.Title}" successfully created for list "${listTitle}"`,
                    newView: {
                        id: createResponse.d.Id,
                        title: createResponse.d.Title,
                        url: `${url}${createResponse.d.ServerRelativeUrl}`,
                        isDefault: createResponse.d.DefaultView,
                        personalView: createResponse.d.PersonalView,
                        rowLimit: createResponse.d.RowLimit
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

        console.error("Error in createListView tool:", errorMessage);

        return {
            content: [{
                type: "text",
                text: `Error creating list view: ${errorMessage}`
            }],
            isError: true
        } as IToolResult;
    }
}

export default createListView;

