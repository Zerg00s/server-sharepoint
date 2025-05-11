// src/tools/getListViews.ts
import request from 'request-promise';
import { IToolResult } from '../interfaces';
import { getSharePointHeaders } from '../auth';
import { SharePointConfig } from '../config';

export interface GetListViewsParams {
    url: string;
    listTitle: string;
    includeFields?: boolean; // Whether to include the fields for each view
    includeHidden?: boolean; // Whether to include hidden views
}

/**
 * Get all views from a SharePoint list
 * @param params Parameters including site URL and list title
 * @param config SharePoint configuration
 * @returns Tool result with list views data
 */
export async function getListViews(
    params: GetListViewsParams,
    config: SharePointConfig
): Promise<IToolResult> {
    const { url, listTitle, includeFields = false, includeHidden = false } = params;
    console.error(`getListViews tool called with URL: ${url}, List Title: ${listTitle}, Include Fields: ${includeFields}, Include Hidden: ${includeHidden}`);

    try {
        // Authenticate with SharePoint
        const headers = await getSharePointHeaders(url, config);
        console.error("Headers prepared with authentication");

        // Encode the list title to handle special characters
        const encodedListTitle = encodeURIComponent(listTitle);
        
        // First, verify the list exists
        console.error(`Verifying list "${listTitle}" exists...`);
        try {
            await request({
                url: `${url}/_api/web/lists/getByTitle('${encodedListTitle}')`,
                headers: headers,
                json: true,
                method: 'GET',
                timeout: 15000
            });
        } catch (error) {
            throw new Error(`List "${listTitle}" not found`);
        }

        // Get all views for the list
        console.error(`Getting views for list "${listTitle}"...`);
        let viewsUrl = `${url}/_api/web/lists/getByTitle('${encodedListTitle}')/views`;
        
        // Add parameter to exclude hidden views if requested
        if (!includeHidden) {
            viewsUrl += `?$filter=Hidden eq false`;
        }
        
        const viewsResponse = await request({
            url: viewsUrl,
            headers: headers,
            json: true,
            method: 'GET',
            timeout: 30000
        });

        const views = viewsResponse.d.results;
        console.error(`Retrieved ${views.length} views from list "${listTitle}"`);
        
        // Process and format the views
        const formattedViews = [];
        
        for (const view of views) {
            const formattedView: any = {
                Id: view.Id,
                Title: view.Title,
                DefaultView: view.DefaultView,
                PersonalView: view.PersonalView,
                ViewType: view.ViewType,
                RowLimit: view.RowLimit,
                Paged: view.Paged,
                Hidden: view.Hidden,
                ServerRelativeUrl: view.ServerRelativeUrl,
                ViewQuery: view.ViewQuery || ''
            };
            
            // Get view fields if requested
            if (includeFields) {
                console.error(`Getting fields for view "${view.Title}"...`);
                try {
                    const viewFieldsResponse = await request({
                        url: `${url}/_api/web/lists/getByTitle('${encodedListTitle}')/views(guid'${view.Id}')/ViewFields`,
                        headers: headers,
                        json: true,
                        method: 'GET',
                        timeout: 20000
                    });
                    
                    formattedView.ViewFields = viewFieldsResponse.d.Items.results || [];
                } catch (fieldsError) {
                    console.error(`Error getting fields for view "${view.Title}": ${fieldsError instanceof Error ? fieldsError.message : String(fieldsError)}`);
                    formattedView.ViewFields = [];
                }
            }
            
            formattedViews.push(formattedView);
        }

        return {
            content: [{
                type: "text",
                text: JSON.stringify({
                    listTitle: listTitle,
                    url: url,
                    totalViews: formattedViews.length,
                    views: formattedViews
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

        console.error("Error in getListViews tool:", errorMessage);

        return {
            content: [{
                type: "text",
                text: `Error fetching list views: ${errorMessage}`
            }],
            isError: true
        } as IToolResult;
    }
}

export default getListViews;
