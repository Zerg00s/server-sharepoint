// src/tools/createModernPage.ts
import request from 'request-promise';
import { IToolResult } from '../interfaces';
import { getSharePointHeaders, getRequestDigest } from '../auth_factory';
import { SharePointConfig } from '../config';

export interface CreateModernPageParams {
    url: string;
    title: string;
    fileName?: string;        // Optional filename (e.g., "sample.aspx")
    pageLayoutType?: string;  // Article, Home, SingleWebPartAppPage, etc.
    description?: string;
    thumbnailUrl?: string;
    promotedState?: number;   // 0=Not promoted, 1=Promoted, 2=Promoted to news
    publishPage?: boolean;    // Whether to publish the page after creation
    content?: string;         // Optional HTML content for the page
}

/**
 * Create a modern page in SharePoint
 * @param params Parameters including site URL, page title, and other page properties
 * @param config SharePoint configuration
 * @returns Tool result with page creation status
 */
export async function createModernPage(
    params: CreateModernPageParams,
    config: SharePointConfig
): Promise<IToolResult> {
    const { 
        url, 
        title, 
        fileName,
        pageLayoutType = "Article", 
        description = "", 
        thumbnailUrl = "",
        promotedState = 0,
        publishPage = true,
        content = ""
    } = params;
    
    console.error(`createModernPage tool called with URL: ${url}, Title: ${title}, FileName: ${fileName || 'auto-generated'}`);

    try {
        // Authenticate with SharePoint
        const headers = await getSharePointHeaders(url, config);
        console.error("Headers prepared with authentication");

        // Get request digest for POST operations
        headers['X-RequestDigest'] = await getRequestDigest(url, headers);
        headers['Content-Type'] = 'application/json;odata=verbose';
        console.error("Headers prepared with request digest");
        
        // Prepare the filename
        let requestedName = "";
        if (fileName) {
            // Keep the filename exactly as provided - DO NOT modify it
            requestedName = fileName;
        }
        
        // Prepare the payload for creating a modern page
        const createPagePayload: {
            __metadata: { type: string };
            Title: string;
            PageLayoutType: string;
            Description: string;
            PromotedState: number;
            Name?: string;      // Filename with or without extension
            BannerImageUrl?: string;
        } = {
            __metadata: { type: 'SP.Publishing.SitePage' },
            Title: title,
            PageLayoutType: pageLayoutType,
            Description: description,
            PromotedState: promotedState
        };
        
        // Only set Name if fileName is explicitly provided
        if (fileName) {
            createPagePayload.Name = requestedName;
            console.error(`Setting Name property to exactly: "${requestedName}"`);
        }
        
        // If thumbnailUrl is provided, add it to the payload
        if (thumbnailUrl) {
            createPagePayload.BannerImageUrl = thumbnailUrl;
        }
        
        // Create the page
        console.error(`Creating modern page with title "${title}"${fileName ? ` and filename "${fileName}"` : ''}...`);
        const createPageResponse = await request({
            url: `${url}/_api/sitepages/pages`,
            method: 'POST',
            headers: headers,
            json: true,
            body: createPagePayload,
            timeout: 30000
        });
        
        // Get the page ID and URL for the response
        const pageId = createPageResponse.d.Id;
        let pageUrl = createPageResponse.d.Url || createPageResponse.d.AbsoluteUrl;
        const createdFileName = createPageResponse.d.FileName || '';
        
        console.error(`Page created with ID: ${pageId}, URL: ${pageUrl}, FileName: ${createdFileName}`);
        
        // If content is provided, add a text web part with the content
        if (content) {
            console.error(`Adding content as text web part to page...`);
            try {
                const webPartData = {
                    dataVersion: "1.4",
                    description: "Text block",
                    title: "Text",
                    serverProcessedContent: {
                        htmlStrings: {
                            html: content
                        },
                        searchablePlainTexts: {},
                        imageSources: {},
                        links: {}
                    },
                    properties: {
                        isInEditMode: false,
                        layoutType: "Normal",
                        position: {
                            zoneIndex: 1,
                            sectionIndex: 1,
                            controlIndex: 1
                        }
                    }
                };
                
                await request({
                    url: `${url}/_api/sitepages/pages(${pageId})/AddWebPart`,
                    method: 'POST',
                    headers: headers,
                    json: true,
                    body: {
                        webPartDataAsJson: JSON.stringify(webPartData),
                        controlType: 3 // Text web part type
                    },
                    timeout: 20000
                });
                
                console.error(`Content added successfully to page`);
            } catch (webPartError) {
                console.error(`Warning: Could not add content to page: ${webPartError instanceof Error ? webPartError.message : String(webPartError)}`);
                // Continue with page creation even if content addition fails
            }
        }
        
        // Save the page changes
        console.error(`Saving page changes...`);
        try {
            await request({
                url: `${url}/_api/sitepages/pages(${pageId})/SavePage`,
                method: 'POST',
                headers: headers,
                json: true,
                timeout: 20000
            });
        } catch (saveError) {
            console.error(`Warning: Could not save page changes: ${saveError instanceof Error ? saveError.message : String(saveError)}`);
            // Continue with page creation even if save fails
        }
        
        // Optionally publish the page if requested
        if (publishPage) {
            console.error(`Publishing page with ID ${pageId}...`);
            try {
                await request({
                    url: `${url}/_api/sitepages/pages(${pageId})/Publish`,
                    method: 'POST',
                    headers: headers,
                    json: true,
                    timeout: 20000
                });
                console.error(`Page published successfully`);
            } catch (publishError) {
                console.error(`Warning: Could not publish page: ${publishError instanceof Error ? publishError.message : String(publishError)}`);
                // Continue with page creation even if publishing fails
            }
        }
        
        // Get the full updated page details for the response
        const pageDetails = {
            id: pageId,
            title: title,
            url: pageUrl,
            fileName: createdFileName,
            requestedFileName: fileName || null,
            layoutType: pageLayoutType,
            description: description,
            promotedState: promotedState,
            isPublished: publishPage
        };
        
        return {
            content: [{
                type: "text",
                text: JSON.stringify({
                    success: true,
                    message: `Modern page "${title}" successfully created${fileName ? ` with filename request "${fileName}"` : ''}${publishPage ? ' and published' : ''}`,
                    page: pageDetails
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

        console.error("Error in createModernPage tool:", errorMessage);

        return {
            content: [{
                type: "text",
                text: `Error creating modern page: ${errorMessage}`
            }],
            isError: true
        } as IToolResult;
    }
}

export default createModernPage;
