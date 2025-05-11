// src/tools/getGlobalNavigationLinks.ts
import request from 'request-promise';
import { IToolResult } from '../interfaces';
import { getSharePointHeaders } from '../auth';
import { SharePointConfig } from '../config';

export interface GetGlobalNavigationLinksParams {
    url: string;
}

/**
 * Get global navigation links from a SharePoint site using MenuState API
 * @param params Parameters including site URL
 * @param config SharePoint configuration
 * @returns Tool result with global navigation links data
 */
export async function getGlobalNavigationLinks(
    params: GetGlobalNavigationLinksParams,
    config: SharePointConfig
): Promise<IToolResult> {
    const { url } = params;
    console.error(`getGlobalNavigationLinks tool called with URL: ${url}`);

    try {
        // Authenticate with SharePoint
        const headers = await getSharePointHeaders(url, config);
        console.error("Headers prepared with authentication");

        // Try an alternative approach first - get web navigation nodes directly
        console.error("Trying direct navigation nodes approach for global navigation...");
        let useDirectApi = true;
        let menuStateResponse;
        
        try {
            const topNavUrl = `${url}/_api/web/Navigation/TopNavigationBar`;
            console.error(`Requesting navigation from: ${topNavUrl}`);
            
            menuStateResponse = await request({
                url: topNavUrl,
                headers: headers,
                json: true,
                method: 'GET',
                timeout: 30000
            });
            
            // Check if we got a valid response
            if (!menuStateResponse.d || !menuStateResponse.d.results) {
                console.error("Direct API did not return expected structure, falling back to MenuState API");
                useDirectApi = false;
            }
        } catch (directApiError) {
            console.error("Direct API approach failed:", directApiError);
            useDirectApi = false;
        }
        
        // Fall back to original MenuState approach if direct approach failed
        if (!useDirectApi) {
            console.error("Getting global navigation links using MenuState API...");
            menuStateResponse = await request({
                url: `${url}/_api/navigation/MenuState?mapProviderName='GlobalNavigationSwitchableProvider'`,
                headers: headers,
                json: true,
                method: 'GET',
                timeout: 30000
            });
        }
        
        let formattedNodes = [];
        let menuState: any = null;
        let accessMethod = useDirectApi ? "direct" : "menustate";
        
        if (useDirectApi) {
            // Process the direct navigation nodes response
            const navNodes = menuStateResponse.d.results || [];
            console.error(`Retrieved ${navNodes.length} global navigation links using direct API`);
            
            // Format the navigation nodes for display
            formattedNodes = navNodes.map((node: any) => ({
                Key: node.Id?.toString() || "",
                Title: node.Title || "",
                Url: node.Url || "",
                IsExternal: node.IsExternal || false,
                HasChildren: false, // Direct API doesn't provide children info
                ParentKey: null,
                Children: [] // Direct API doesn't provide children
            }));
        } else {
            // Process the menu state to extract navigation nodes
            menuState = menuStateResponse.MenuState;
            const navNodes = menuState.Nodes || [];
            console.error(`Retrieved ${navNodes.length} global navigation links using MenuState API`);
            
            // Format the navigation nodes for display
            formattedNodes = navNodes
                .filter((node: any) => !node.IsDeleted) // Filter out deleted nodes
                .map((node: any) => ({
                    Key: node.Key,
                    Title: node.Title,
                    Url: node.SimpleUrl,
                    IsExternal: node.FriendlyUrlSegment === '' && !node.SimpleUrl.startsWith(menuState.SPSitePrefix),
                    HasChildren: (node.Nodes && node.Nodes.length > 0) || false,
                    ParentKey: node.ParentKey || null,
                    Children: (node.Nodes && node.Nodes.length > 0) ? 
                        node.Nodes
                            .filter((childNode: any) => !childNode.IsDeleted)
                            .map((childNode: any) => ({
                                Key: childNode.Key,
                                Title: childNode.Title,
                                Url: childNode.SimpleUrl,
                                IsExternal: childNode.FriendlyUrlSegment === '' && !childNode.SimpleUrl.startsWith(menuState.SPSitePrefix),
                                ParentKey: node.Key
                            })) : []
                }));
        }

        const responseData: any = {
            siteUrl: url,
            navLinksCount: formattedNodes.length,
            accessMethod: accessMethod,
            navLinks: formattedNodes
        };
        
        // Include menuState info if available
        if (menuState) {
            responseData.menuState = {
                version: menuState.Version,
                startingNodeKey: menuState.StartingNodeKey,
                startingNodeTitle: menuState.StartingNodeTitle,
                sitePrefix: menuState.SPSitePrefix,
                webPrefix: menuState.SPWebPrefix
            };
        }
        
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

        console.error("Error in getGlobalNavigationLinks tool:", errorMessage);

        return {
            content: [{
                type: "text",
                text: `Error getting global navigation links: ${errorMessage}`
            }],
            isError: true
        } as IToolResult;
    }
}

export default getGlobalNavigationLinks;
