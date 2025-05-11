// src/tools/createList.ts
import request from 'request-promise';
import { IToolResult, IListCreationData } from '../interfaces';
import { getSharePointHeaders, getRequestDigest } from '../auth';
import { SharePointConfig } from '../config';

export interface CreateListParams {
    url: string;
    listData: IListCreationData & {
        Url?: string; // Add custom URL property
    };
}

/**
 * Create a new SharePoint list
 * @param params Parameters including site URL and list creation data
 * @param config SharePoint configuration
 * @returns Tool result with creation status and new list info
 */
export async function createList(
    params: CreateListParams,
    config: SharePointConfig
): Promise<IToolResult> {
    const { url, listData } = params;
    console.error(`createList tool called with URL: ${url}, List Title: ${listData.Title}`);

    try {
        // Authenticate with SharePoint
        const headers = await getSharePointHeaders(url, config);
        console.error("Headers prepared with authentication");

        // Get request digest for POST operations
        headers['X-RequestDigest'] = await getRequestDigest(url, headers);
        headers['Content-Type'] = 'application/json;odata=verbose';
        console.error("Headers prepared with request digest");
        
        // Set default template type if not provided
        const templateType = listData.TemplateType || 100; // 100 is generic list, 101 is document library
        
        // Prepare the list creation data
        const createPayload: any = {
            __metadata: { type: 'SP.List' },
            BaseTemplate: templateType,
            Title: listData.Title,
            Description: listData.Description || '',
            AllowContentTypes: listData.AllowContentTypes || false,
            ContentTypesEnabled: listData.ContentTypesEnabled || false,
            EnableVersioning: listData.EnableVersioning || false,
            EnableMinorVersions: listData.EnableMinorVersions || false,
            EnableModeration: listData.EnableModeration || false
        };
        
        // Handle custom URL if provided
        if (listData.Url) {
            console.error(`Using custom URL: ${listData.Url}`);
            // Use CustomSchemaXml to specify a clean URL
            createPayload.CustomSchemaXml = `<List xmlns:ows="Microsoft SharePoint" Title="${listData.Title}" FolderCreation="FALSE" Direction="$Resources:Direction;" Url="${listData.Url}" BaseType="${templateType}" />`;
        }
        
        // Add any other properties from listData to the payload
        Object.keys(listData).forEach(key => {
            if (!createPayload[key] && key !== 'TemplateType' && key !== 'Url') {
                createPayload[key] = listData[key];
            }
        });
        
        console.error(`Creating list with payload: ${JSON.stringify(createPayload)}`);
        
        // Create the list
        const createResponse = await request({
            url: `${url}/_api/web/lists`,
            headers: headers,
            json: true,
            method: 'POST',
            body: createPayload,
            timeout: 30000
        });
        
        return {
            content: [{
                type: "text",
                text: JSON.stringify({
                    success: true,
                    message: `List "${listData.Title}" successfully created`,
                    newList: {
                        id: createResponse.d.Id,
                        title: createResponse.d.Title,
                        url: `${url}${createResponse.d.RootFolder.ServerRelativeUrl}`,
                        internalUrl: createResponse.d.RootFolder.ServerRelativeUrl,
                        templateType: createResponse.d.BaseTemplate,
                        itemCount: 0,
                        created: createResponse.d.Created
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

        console.error("Error in createList tool:", errorMessage);

        return {
            content: [{
                type: "text",
                text: `Error creating list: ${errorMessage}`
            }],
            isError: true
        } as IToolResult;
    }
}

export default createList;