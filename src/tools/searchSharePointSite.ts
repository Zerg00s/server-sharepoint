// src/tools/searchSharePointSite.ts
import axios from 'axios';
import { IToolResult } from '../interfaces';
import { getSharePointHeaders, getRequestDigest } from '../auth_factory';
import { SharePointConfig } from '../config';

export interface SearchSharePointSiteParams {
    url: string;
    query: string;
    rowLimit?: number;
    startRow?: number;
    selectProperties?: string[];
    sourceid?: string;
}

interface ISearchResult {
    Title?: string;
    Path?: string;
    Size?: number;
    Created?: string;
    LastModifiedTime?: string;
    FileType?: string;
    Author?: string;
    ContentClass?: string;
    SiteName?: string;
    HitHighlightedSummary?: string;
    [key: string]: any;
}

/**
 * Search within a SharePoint site using KQL query
 * @param params Parameters including the site URL and KQL query
 * @param config SharePoint configuration
 * @returns Tool result with search results
 */
export async function searchSharePointSite(
    params: SearchSharePointSiteParams,
    config: SharePointConfig
): Promise<IToolResult> {
    const { url, query, rowLimit = 50, startRow = 0, selectProperties = [], sourceid } = params;
    console.error(`searchSharePointSite tool called with URL: ${url}, Query: ${query}`);

    try {
        // Authenticate with SharePoint
        const headers = await getSharePointHeaders(url, config);
        
        // Get request digest for POST request
        const requestDigest = await getRequestDigest(url, headers);
        
        // Prepare select properties array
        const defaultSelectProperties = [
            "Title",
            "Path",
            "Size",
            "Created",
            "LastModifiedTime",
            "FileType",
            "Author",
            "ContentClass",
            "SiteName",
            "HitHighlightedSummary"
        ];
        
        // Use provided select properties or default ones
        const properties = selectProperties.length > 0 ? selectProperties : defaultSelectProperties;
        
        // Prepare search query request
        const searchPostUrl = `${url}/_api/search/postquery`;
        const searchPayload = {
            request: {
                Querytext: query,
                RowLimit: rowLimit,
                StartRow: startRow,
                SelectProperties: { results: properties },
                SourceId: sourceid || undefined,
                TrimDuplicates: true
            }
        };
        
        console.error("Making search request to SharePoint API...");
        
        // Execute search request
        const response = await axios({
            url: searchPostUrl,
            method: 'POST',
            headers: {
                ...headers,
                'Accept': 'application/json;odata=verbose',
                'Content-Type': 'application/json;odata=verbose',
                'X-RequestDigest': requestDigest
            },
            data: JSON.stringify(searchPayload),
            timeout: 30000
        });
        
        // Process search results
        const searchResults = response.data?.d?.postquery?.PrimaryQueryResult?.RelevantResults?.Table?.Rows?.results || [];
        console.error(`Received ${searchResults.length} search results`);
        
        // Format search results
        const formattedResults: ISearchResult[] = searchResults.map((row: any) => {
            const result: ISearchResult = {};
            const cells = row.Cells.results;
            
            // Process each cell and add to result object
            cells.forEach((cell: any) => {
                if (cell.Key && cell.Value !== null) {
                    result[cell.Key] = cell.Value;
                }
            });
            
            return result;
        });
        
        // Return search results
        return {
            content: [{
                type: "text",
                text: JSON.stringify({
                    query,
                    totalRows: response.data?.d?.postquery?.PrimaryQueryResult?.RelevantResults?.TotalRows || 0,
                    totalRowsIncludingDuplicates: response.data?.d?.postquery?.PrimaryQueryResult?.RelevantResults?.TotalRowsIncludingDuplicates || 0,
                    results: formattedResults
                }, null, 2)
            }]
        } as IToolResult;
        
    } catch (error: unknown) {
        // Error handling
        let errorMessage: string;
        
        if (error instanceof Error) {
            errorMessage = error.message;
            console.error("Error stack:", error.stack);
        } else if (typeof error === 'string') {
            errorMessage = error;
        } else {
            errorMessage = "Unknown error occurred";
        }
        
        console.error("Error in searchSharePointSite tool:", errorMessage);
        
        return {
            content: [{
                type: "text",
                text: `Error searching SharePoint site: ${errorMessage}`
            }],
            isError: true
        } as IToolResult;
    }
}

export default searchSharePointSite;