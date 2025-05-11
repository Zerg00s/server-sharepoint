// src/tools/getQuickNavigationLinks.ts
import request from 'request-promise';
import { IToolResult } from '../interfaces';
import { getSharePointHeaders } from '../auth';
import { SharePointConfig } from '../config';

export interface GetQuickNavigationLinksParams {
    url: string;
}

/**
 * Get quick navigation links (left navigation) from a SharePoint site using MenuState API
 * @param params Parameters including site URL
 * @param config SharePoint configuration
 * @returns Tool result with quick navigation links data
 */
export async function getQuickNavigationLinks(
    params: GetQuickNavigationLinksParams,
    config: SharePointConfig
): Promise<IToolResult> {
    const { url } = params;
    console.error(`getQuickNavigationLinks tool called with URL: ${url}`);

    try {
        // Log configuration credentials (masked for security)
        console.error("SharePoint Configuration Check:");
        console.error(`Client ID: ${config.clientId ? '✓ Present' : '✗ Missing'}`);
        console.error(`Client Secret: ${config.clientSecret ? '✓ Present' : '✗ Missing'}`);
        console.error(`Tenant ID: ${config.tenantId ? '✓ Present' : '✗ Missing'}`);

        // Authenticate with SharePoint
        console.error("Attempting to authenticate with SharePoint...");
        const headers = await getSharePointHeaders(url, config);
        console.error("Headers prepared with authentication: ", JSON.stringify({
            ...headers,
            // Mask any auth tokens for security
            Authorization: headers.Authorization ? "Bearer [MASKED]" : undefined,
            "X-RequestDigest": headers["X-RequestDigest"] ? "[MASKED]" : undefined
        }));

        // Try an alternative approach first - get web navigation nodes directly
        console.error("Trying direct navigation nodes approach...");
        
        try {
            const quickLaunchUrl = `${url}/_api/web/Navigation/QuickLaunch`;
            console.error(`Requesting navigation from: ${quickLaunchUrl}`);
            
            const navResponse = await request({
                url: quickLaunchUrl,
                headers: headers,
                json: true,
                method: 'GET',
                timeout: 30000,
                resolveWithFullResponse: true,
                simple: false // Don't throw on non-2xx responses
            });
            
            // Check response status
            console.error(`Response status code: ${navResponse.statusCode}`);
            
            if (navResponse.statusCode >= 400) {
                // If this approach fails, fall back to the original MenuState approach
                console.error("Direct approach failed, falling back to MenuState API...");
                
                const menuStateUrl = `${url}/_api/navigation/MenuState?mapProviderName='QuickLaunch'`;
                console.error(`Requesting menu state from: ${menuStateUrl}`);
                
                const menuStateResponse = await request({
                    url: menuStateUrl,
                    headers: headers,
                    json: true,
                    method: 'GET',
                    timeout: 30000,
                    resolveWithFullResponse: true,
                    simple: false
                });
                
                if (menuStateResponse.statusCode >= 400) {
                    throw new Error(`HTTP Error ${menuStateResponse.statusCode}: ${JSON.stringify(menuStateResponse.body)}`);
                }
                
                // Process the menu state to extract navigation nodes (original approach)
                const menuState = menuStateResponse.body.MenuState;
                if (!menuState) {
                    throw new Error("MenuState not found in response: " + JSON.stringify(menuStateResponse.body));
                }
                
                const navNodes = menuState.Nodes || [];
                console.error(`Retrieved ${navNodes.length} quick navigation links using MenuState API`);
                
                // Format the navigation nodes for display
                const formattedNodes = navNodes
                    .filter((node: any) => !node.IsDeleted) // Filter out deleted nodes
                    .map((node: any) => ({
                        Key: node.Key,
                        Title: node.Title,
                        Url: node.SimpleUrl,
                        IsExternal: node.FriendlyUrlSegment === '' && 
                                   menuState.SPSitePrefix && 
                                   !node.SimpleUrl.startsWith(menuState.SPSitePrefix),
                        HasChildren: (node.Nodes && node.Nodes.length > 0) || false,
                        ParentKey: node.ParentKey || null,
                        Children: (node.Nodes && node.Nodes.length > 0) ? 
                            node.Nodes
                                .filter((childNode: any) => !childNode.IsDeleted)
                                .map((childNode: any) => ({
                                    Key: childNode.Key,
                                    Title: childNode.Title,
                                    Url: childNode.SimpleUrl,
                                    IsExternal: childNode.FriendlyUrlSegment === '' && 
                                             menuState.SPSitePrefix &&
                                             !childNode.SimpleUrl.startsWith(menuState.SPSitePrefix),
                                    ParentKey: node.Key
                                })) : []
                    }));
                    
                return {
                    content: [{
                        type: "text",
                        text: JSON.stringify({
                            siteUrl: url,
                            navLinksCount: formattedNodes.length,
                            menuState: {
                                version: menuState.Version,
                                startingNodeKey: menuState.StartingNodeKey,
                                startingNodeTitle: menuState.StartingNodeTitle,
                                sitePrefix: menuState.SPSitePrefix,
                                webPrefix: menuState.SPWebPrefix
                            },
                            navLinks: formattedNodes
                        }, null, 2)
                    }]
                } as IToolResult;
            }
            
            // Process the direct navigation nodes response
            const navNodes = navResponse.body.d?.results || [];
            console.error(`Retrieved ${navNodes.length} quick navigation links using direct API`);
            
            // Format the navigation nodes for display
            const formattedNodes = navNodes.map((node: any) => ({
                Key: node.Id?.toString() || "",
                Title: node.Title || "",
                Url: node.Url || "",
                IsExternal: node.IsExternal || false,
                HasChildren: false, // Direct API doesn't provide children info
                ParentKey: null,
                Children: [] // Direct API doesn't provide children
            }));

            return {
                content: [{
                    type: "text",
                    text: JSON.stringify({
                        siteUrl: url,
                        navLinksCount: formattedNodes.length,
                        accessMethod: "direct",
                        navLinks: formattedNodes
                    }, null, 2)
                }]
            } as IToolResult;
        } catch (requestError: any) {
            // Handle request specific errors
            console.error("Error during request:", requestError);
            
            // Try to get more information from the error
            let detailedError = "";
            if (requestError.response) {
                detailedError = `Status: ${requestError.response.statusCode}, ` +
                                `Headers: ${JSON.stringify(requestError.response.headers)}, ` +
                                `Body: ${JSON.stringify(requestError.response.body)}`;
            } else if (requestError.error) {
                detailedError = `${JSON.stringify(requestError.error)}`;
            }
            
            // Try a third approach - see if we can get left navigation using Web ClientObject Model
            try {
                console.error("Trying ClientObject Model approach...");
                const clientObjectUrl = `${url}/_api/web/Navigation`;
                console.error(`Requesting navigation from: ${clientObjectUrl}`);
                
                const coResponse = await request({
                    url: clientObjectUrl,
                    headers: headers,
                    json: true,
                    method: 'GET',
                    timeout: 30000
                });
                
                // If we have a successful response, try to extract any usable navigation info
                if (coResponse && coResponse.d) {
                    const navData = {
                        siteUrl: url,
                        navLinksCount: 0,
                        accessMethod: "clientobject",
                        warning: "Limited navigation data available. Using root navigation properties.",
                        navigationInfo: coResponse.d
                    };
                    
                    return {
                        content: [{
                            type: "text",
                            text: JSON.stringify(navData, null, 2)
                        }]
                    } as IToolResult;
                }
            } catch (finalError) {
                console.error("All navigation retrieval approaches failed:", finalError);
            }
            
            throw new Error(`API request failed: ${requestError.message || requestError}. Details: ${detailedError}`);
        }
    } catch (error: unknown) {
        // Type-safe error handling with more detailed information
        let errorMessage: string;

        if (error instanceof Error) {
            errorMessage = error.message;
            console.error("Error stack:", error.stack);
        } else if (typeof error === 'string') {
            errorMessage = error;
        } else {
            errorMessage = "Unknown error occurred";
            console.error("Unknown error type:", error);
        }

        console.error("Error in getQuickNavigationLinks tool:", errorMessage);

        return {
            content: [{
                type: "text",
                text: `Error getting quick navigation links: ${errorMessage}\n\nTry the following to troubleshoot:\n1. Verify SharePoint App permissions\n2. Check if the SharePoint API endpoints are accessible\n3. Ensure the app credentials are correct and not expired\n4. Try using getSite first to test basic connectivity`
            }],
            isError: true
        } as IToolResult;
    }
}

export default getQuickNavigationLinks;
