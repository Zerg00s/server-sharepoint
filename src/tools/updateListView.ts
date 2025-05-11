// src/tools/updateListView.ts
import request from 'request-promise';
import { IToolResult, IListViewData, ISharePointField } from '../interfaces';
import { getSharePointHeaders, getRequestDigest } from '../auth_factory';
import { SharePointConfig } from '../config';

export interface UpdateListViewParams {
    url: string;
    listTitle: string;
    viewTitle: string;
    updateData: IListViewData;
    appendFields?: boolean; // Optional flag to append fields instead of replacing them
}

/**
 * Update an existing view for a SharePoint list
 * @param params Parameters including site URL, list title, view title, and update data
 * @param config SharePoint configuration
 * @returns Tool result with update status
 */
export async function updateListView(
    params: UpdateListViewParams,
    config: SharePointConfig
): Promise<IToolResult> {
    const { url, listTitle, viewTitle, updateData, appendFields = false } = params;
    console.error(`updateListView tool called with URL: ${url}, List Title: ${listTitle}, View: ${viewTitle}, Append Fields: ${appendFields}`);

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

        // Encode the list title and view title to handle special characters
        const encodedListTitle = encodeURIComponent(listTitle);
        const encodedViewTitle = encodeURIComponent(viewTitle);
        
        // First, verify the list and view exist
        console.error(`Verifying list "${listTitle}" and view "${viewTitle}" exist...`);
        let viewDetails;
        try {
            viewDetails = await request({
                url: `${url}/_api/web/lists/getByTitle('${encodedListTitle}')/views/getByTitle('${encodedViewTitle}')`,
                headers: { ...headers, 'Content-Type': undefined, 'X-HTTP-Method': undefined, 'IF-MATCH': undefined },
                json: true,
                method: 'GET',
                timeout: 15000
            });
        } catch (error) {
            throw new Error(`View "${viewTitle}" not found in list "${listTitle}"`);
        }
        
        // Store the view ID for later use
        const viewId = viewDetails.d.Id;
        
        // If ViewFields are specified, validate that they exist in the list
        if (updateData.ViewFields && updateData.ViewFields.length > 0) {
            console.error(`Validating view fields...`);
            const fieldsResponse = await request({
                url: `${url}/_api/web/lists/getByTitle('${encodedListTitle}')/fields?$select=InternalName,Title`,
                headers: { ...headers, 'Content-Type': undefined, 'X-HTTP-Method': undefined, 'IF-MATCH': undefined },
                json: true,
                method: 'GET',
                timeout: 30000
            });
            
            const availableFields = fieldsResponse.d.results.map((field: ISharePointField) => field.InternalName);
            
            // Check if all requested view fields exist in the list
            const invalidFields = updateData.ViewFields.filter((field: string) => !availableFields.includes(field));
            
            if (invalidFields.length > 0) {
                throw new Error(`Some view fields do not exist in the list: ${invalidFields.join(', ')}`);
            }
        }
        
        // Handle setting view as default directly on the view if specified
        if (updateData.SetAsDefaultView !== undefined) {
            console.error(`Setting DefaultView property to ${updateData.SetAsDefaultView}...`);
            
            const defaultViewPayload = {
                __metadata: { type: 'SP.View' },
                DefaultView: updateData.SetAsDefaultView
            };
            
            try {
                await request({
                    url: `${url}/_api/web/lists/getByTitle('${encodedListTitle}')/views/getByTitle('${encodedViewTitle}')`,
                    headers: headers,
                    json: true,
                    method: 'POST',
                    body: defaultViewPayload,
                    timeout: 20000
                });
                
                console.error(`DefaultView property set successfully`);
            } catch (error) {
                console.error(`Warning: Error setting DefaultView property: ${error instanceof Error ? error.message : String(error)}`);
                // Continue with other updates even if this one fails
            }
        }
        
        // Prepare the update data for other view properties
        const updatePayload: any = {
            __metadata: { type: 'SP.View' }
        };
        
        // Add update properties from the updateData, excluding specially handled ones
        Object.entries(updateData).forEach(([key, value]) => {
            // Skip specially handled properties
            if (key !== 'ViewFields' && key !== 'Title' && key !== 'SetAsDefaultView') {
                updatePayload[key] = value;
            }
        });
        
        // Only proceed with the update if there are properties to update
        if (Object.keys(updatePayload).length > 1) { // > 1 because __metadata is always there
            console.error(`Updating view properties with payload: ${JSON.stringify(updatePayload)}`);
            
            // Update the view properties
            await request({
                url: `${url}/_api/web/lists/getByTitle('${encodedListTitle}')/views/getByTitle('${encodedViewTitle}')`,
                headers: headers,
                json: true,
                method: 'POST',
                body: updatePayload,
                timeout: 20000
            });
        }
        
        // If ViewFields are specified, update them
        if (updateData.ViewFields && updateData.ViewFields.length > 0) {
            if (appendFields) {
                // Use the addViewField endpoint to add fields individually without replacing existing ones
                console.error(`Appending fields to view: ${updateData.ViewFields.join(', ')}`);
                
                for (const field of updateData.ViewFields) {
                    try {
                        console.error(`Adding field "${field}" to view...`);
                        
                        await request({
                            url: `${url}/_api/web/lists/getByTitle('${encodedListTitle}')/Views('${viewId}')/ViewFields/addViewField`,
                            headers: { 
                                ...headers,
                                'X-HTTP-Method': undefined, // We want a POST, not MERGE
                                'IF-MATCH': undefined
                            },
                            json: true,
                            method: 'POST',
                            body: { "strField": field },
                            timeout: 10000
                        });
                    } catch (fieldError) {
                        // Field might already be in the view, or other error - log but continue
                        console.error(`Warning: Could not add field "${field}": ${fieldError instanceof Error ? fieldError.message : String(fieldError)}`);
                    }
                }
            } else {
                // Replace all fields in one operation
                console.error(`Replacing view fields with: ${updateData.ViewFields.join(', ')}`);
                
                const viewFieldsUpdatePayload = {
                    __metadata: { type: 'SP.View' },
                    ViewFields: {
                        __metadata: { type: 'Collection(Edm.String)' },
                        results: updateData.ViewFields
                    }
                };
                
                await request({
                    url: `${url}/_api/web/lists/getByTitle('${encodedListTitle}')/views/getByTitle('${encodedViewTitle}')`,
                    headers: headers,
                    json: true,
                    method: 'POST',
                    body: viewFieldsUpdatePayload,
                    timeout: 30000
                });
            }
        }
        
        // If Title is being changed, update the title in a separate request
        if (updateData.Title && updateData.Title !== viewTitle) {
            console.error(`Updating view title from "${viewTitle}" to "${updateData.Title}"`);
            
            const titleUpdatePayload = {
                __metadata: { type: 'SP.View' },
                Title: updateData.Title
            };
            
            await request({
                url: `${url}/_api/web/lists/getByTitle('${encodedListTitle}')/views/getByTitle('${encodedViewTitle}')`,
                headers: headers,
                json: true,
                method: 'POST',
                body: titleUpdatePayload,
                timeout: 20000
            });
        }
        
        // Get the updated view title (in case it was changed)
        const updatedViewTitle = updateData.Title || viewTitle;
        const encodedUpdatedViewTitle = encodeURIComponent(updatedViewTitle);
        
        // Get the updated view to return its new state
        const updatedView = await request({
            url: `${url}/_api/web/lists/getByTitle('${encodedListTitle}')/views/getByTitle('${encodedUpdatedViewTitle}')`,
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
                    message: `View "${viewTitle}" successfully updated in list "${listTitle}"`,
                    updatedView: {
                        id: updatedView.d.Id,
                        title: updatedView.d.Title,
                        url: `${url}${updatedView.d.ServerRelativeUrl}`,
                        isDefault: updatedView.d.DefaultView,
                        personalView: updatedView.d.PersonalView,
                        rowLimit: updatedView.d.RowLimit
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

        console.error("Error in updateListView tool:", errorMessage);

        return {
            content: [{
                type: "text",
                text: `Error updating list view: ${errorMessage}`
            }],
            isError: true
        } as IToolResult;
    }
}

export default updateListView;

