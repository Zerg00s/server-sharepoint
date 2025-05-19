// src/tools/getModernPage.ts
import request from 'request-promise';
import { IToolResult } from '../interfaces';
import { getSharePointHeaders } from '../auth_factory';
import { SharePointConfig } from '../config';

export interface GetModernPageParams {
    url: string;
    pageId: number;          // ID of the page to retrieve
}

/**
 * Get a specific modern page by ID from a SharePoint site
 * @param params Parameters including site URL and page ID
 * @param config SharePoint configuration
 * @returns Tool result with page data
 */
export async function getModernPage(
    params: GetModernPageParams,
    config: SharePointConfig
): Promise<IToolResult> {
    const { url, pageId } = params;
    console.error(`getModernPage tool called with URL: ${url}, Page ID: ${pageId}`);

    try {
        // Authenticate with SharePoint
        const headers = await getSharePointHeaders(url, config);
        console.error("Headers prepared with authentication");
        
        // Construct the query URL for the specific page
        const queryUrl = `${url}/_api/sitepages/pages(${pageId})`;
        console.error(`Getting page with ID: ${pageId}`);
        
        // Make the request to get the page details
        const response = await request({
            url: queryUrl,
            method: 'GET',
            headers: { ...headers, 'Content-Type': undefined },
            json: true,
            timeout: 30000
        });
        
        if (!response.d) {
            throw new Error(`Page with ID ${pageId} not found`);
        }
        
        // Extract canvas content if available
        let canvasContent = null;
        if (response.d.CanvasContent1) {
            try {
                // Try to parse the canvas content, which might contain web part information
                canvasContent = JSON.parse(response.d.CanvasContent1);
                console.error("Successfully parsed Canvas Content");
            } catch (e) {
                console.error("Failed to parse Canvas Content:", e);
                canvasContent = response.d.CanvasContent1; // Keep as string
            }
        }
        
        // Build page details
        const page = {
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
            publishingStatus: response.d.PublishingStatus,
            canvasContent: canvasContent
        };
        
        // Final response structure
        const responseData = {
            success: true,
            page: page
        };
        
        return {
            content: [{
                type: "text",
                text: JSON.stringify(responseData, null, 2)
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

        console.error("Error in getModernPage tool:", errorMessage);

        return {
            content: [{
                type: "text",
                text: `Error retrieving modern page: ${errorMessage}`
            }],
            isError: true
        } as IToolResult;
    }
}

export default getModernPage;
