// src/tools/getModernPages.ts
import request from 'request-promise';
import { IToolResult } from '../interfaces';
import { getSharePointHeaders } from '../auth_factory';
import { SharePointConfig } from '../config';

export interface GetModernPagesParams {
    url: string;
    pageTitle?: string;  // Optional - filter by page title
    limit?: number;      // Maximum number of pages to return
}

/**
 * Get modern pages from a SharePoint site
 * @param params Parameters including site URL and optional filters
 * @param config SharePoint configuration
 * @returns Tool result with pages data
 */
export async function getModernPages(
    params: GetModernPagesParams,
    config: SharePointConfig
): Promise<IToolResult> {
    const { url, pageTitle, limit = 50 } = params;
    console.error(`getModernPages tool called with URL: ${url}, Title Filter: ${pageTitle || 'None'}, Limit: ${limit}`);

    try {
        // Authenticate with SharePoint
        const headers = await getSharePointHeaders(url, config);
        console.error("Headers prepared with authentication");
        
        // Construct the query URL
        let queryUrl = `${url}/_api/sitepages/pages`;
        
        // Add filter if page title is specified
        if (pageTitle) {
            // Handle special characters in the title for OData filtering
            const encodedTitle = encodeURIComponent(pageTitle);
            queryUrl += `/GetByTitle('${encodedTitle}')`;
            console.error(`Getting specific page with title: ${pageTitle}`);
        } else {
            // Add limit if getting multiple pages
            queryUrl += `?$top=${limit}`;
            
            // Add orderby to get newest pages first
            queryUrl += `&$orderby=Created desc`;
            
            console.error(`Getting up to ${limit} pages, ordered by creation date`);
        }
        
        // Make the request to get pages
        const response = await request({
            url: queryUrl,
            method: 'GET',
            headers: { ...headers, 'Content-Type': undefined },
            json: true,
            timeout: 30000
        });
        
        // Process the response based on whether we're getting a single page or multiple pages
        let pages = [];
        
        if (pageTitle) {
            // Single page response
            if (response.d) {
                pages = [{
                    id: response.d.Id,
                    title: response.d.Title,
                    description: response.d.Description || '',
                    url: response.d.Url || response.d.AbsoluteUrl,
                    fileName: response.d.FileName,
                    pageLayoutType: response.d.PageLayoutType,
                    promotedState: response.d.PromotedState,
                    created: response.d.Created,
                    modified: response.d.Modified,
                    firstPublished: response.d.FirstPublished,
                    bannerImageUrl: response.d.BannerImageUrl || null,
                    topicHeader: response.d.TopicHeader || null,
                    publishingStatus: response.d.PublishingStatus
                }];
            }
        } else {
            // Multiple pages response
            if (response.d && response.d.results) {
                pages = response.d.results.map((page: any) => ({
                    id: page.Id,
                    title: page.Title,
                    description: page.Description || '',
                    url: page.Url || page.AbsoluteUrl,
                    fileName: page.FileName,
                    pageLayoutType: page.PageLayoutType,
                    promotedState: page.PromotedState,
                    created: page.Created,
                    modified: page.Modified,
                    firstPublished: page.FirstPublished,
                    bannerImageUrl: page.BannerImageUrl || null,
                    topicHeader: page.TopicHeader || null,
                    publishingStatus: page.PublishingStatus
                }));
            }
        }
        
        return {
            content: [{
                type: "text",
                text: JSON.stringify({
                    success: true,
                    totalPages: pages.length,
                    pages: pages
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

        console.error("Error in getModernPages tool:", errorMessage);

        return {
            content: [{
                type: "text",
                text: `Error retrieving modern pages: ${errorMessage}`
            }],
            isError: true
        } as IToolResult;
    }
}

export default getModernPages;
