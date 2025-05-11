// src/tools/addNavigationLink.ts
import request from 'request-promise';
import { IToolResult } from '../interfaces';
import { getSharePointHeaders, getRequestDigest } from '../auth_factory';
import { SharePointConfig } from '../config';

export interface AddNavigationLinkParams {
    url: string;
    linkData: {
        Title: string;
        Url: string;
        IsExternal?: boolean;
        ParentKey?: string; // Key of the parent node (if adding a child link)
    };
    navigationType: 'Global' | 'Quick'; // Global = top navigation, Quick = left navigation
}

/**
 * Add a navigation link to a SharePoint site using MenuState/SaveMenuState API
 * @param params Parameters including site URL, link data, and navigation type
 * @param config SharePoint configuration
 * @returns Tool result with creation status
 */
export async function addNavigationLink(
    params: AddNavigationLinkParams,
    config: SharePointConfig
): Promise<IToolResult> {
    const { url, linkData, navigationType } = params;
    console.error(`addNavigationLink tool called with URL: ${url}, Navigation Type: ${navigationType}`);

    try {
        // Validate required parameters
        if (!linkData.Title) {
            throw new Error("Link Title is required");
        }
        
        if (!linkData.Url) {
            throw new Error("Link URL is required");
        }

        // Authenticate with SharePoint
        const headers = await getSharePointHeaders(url, config);
        console.error("Headers prepared with authentication");

        // Get request digest for POST operations
        headers['X-RequestDigest'] = await getRequestDigest(url, headers);
        headers['Content-Type'] = 'application/json;odata=verbose';
        console.error("Headers prepared with request digest");
        
        // Try the direct approach first (using Navigation API)
        try {
            console.error(`Trying direct API approach for adding ${navigationType.toLowerCase()} navigation link...`);
            
            // Determine the endpoint based on navigation type
            const navigationApiUrl = navigationType === 'Global' 
                ? `${url}/_api/web/Navigation/TopNavigationBar`
                : `${url}/_api/web/Navigation/QuickLaunch`;
            
            // Prepare the request payload
            const createLinkPayload = {
                __metadata: { type: 'SP.NavigationNode' },
                Title: linkData.Title,
                Url: linkData.Url,
                IsExternal: linkData.IsExternal === true
            };
            
            // If parent ID is provided, we need to use a different endpoint
            if (linkData.ParentKey) {
                console.error(`Adding as child node to parent with key: ${linkData.ParentKey}`);
                
                // Convert ParentKey string to number (SharePoint Navigation API expects ID as number)
                const parentId = parseInt(linkData.ParentKey, 10);
                if (isNaN(parentId)) {
                    throw new Error("Parent key must be a valid integer ID");
                }
                
                // Add to specific parent node
                await request({
                    url: `${url}/_api/web/Navigation/GetNodeById(${parentId})/Children`,
                    headers: headers,
                    json: true,
                    method: 'POST',
                    body: createLinkPayload,
                    timeout: 30000
                });
                
                console.error(`Successfully added child navigation link using direct API`);
                
                return {
                    content: [{
                        type: "text",
                        text: JSON.stringify({
                            success: true,
                            message: `Navigation link "${linkData.Title}" successfully created in ${navigationType.toLowerCase()} navigation`,
                            newLink: {
                                title: linkData.Title,
                                url: linkData.Url,
                                isExternal: linkData.IsExternal === true,
                                parentId: parentId,
                                navigationType: navigationType,
                                accessMethod: "direct"
                            }
                        }, null, 2)
                    }]
                } as IToolResult;
            }
            
            // Add to top-level navigation
            const response = await request({
                url: navigationApiUrl,
                headers: headers,
                json: true,
                method: 'POST',
                body: createLinkPayload,
                timeout: 30000
            });
            
            console.error(`Successfully added navigation link using direct API`);
            
            return {
                content: [{
                    type: "text",
                    text: JSON.stringify({
                        success: true,
                        message: `Navigation link "${linkData.Title}" successfully created in ${navigationType.toLowerCase()} navigation`,
                        newLink: {
                            id: response.d?.Id,
                            title: linkData.Title,
                            url: linkData.Url,
                            isExternal: linkData.IsExternal === true,
                            navigationType: navigationType,
                            accessMethod: "direct"
                        }
                    }, null, 2)
                }]
            } as IToolResult;
        }
        catch (directApiError) {
            console.error("Direct API approach failed, falling back to MenuState API...", directApiError);
            // Continue with MenuState approach
        }
        
        // Fall back to MenuState approach
        const providerName = navigationType === 'Global' ? 'GlobalNavigationSwitchableProvider' : 'QuickLaunch';
        
        // First, get the current menu state
        console.error(`Getting current ${navigationType.toLowerCase()} navigation menu state...`);
        const menuStateResponse = await request({
            url: `${url}/_api/navigation/MenuState?mapProviderName='${providerName}'`,
            headers: { ...headers, 'Content-Type': undefined },
            json: true,
            method: 'GET',
            timeout: 30000
        });
        
        const menuState = menuStateResponse.MenuState;
        console.error(`Retrieved menu state with ${menuState.Nodes.length} nodes`);
        
        // Generate a unique key for the new node
        // Find the highest existing key and increment it
        let highestKey = 1000; // Starting value
        
        const findHighestKey = (nodes: any[]) => {
            for (const node of nodes || []) {
                if (node.Key) {
                    const keyNum = parseInt(node.Key, 10);
                    if (!isNaN(keyNum) && keyNum > highestKey) {
                        highestKey = keyNum;
                    }
                }
                // Check child nodes recursively
                if (node.Nodes && node.Nodes.length > 0) {
                    findHighestKey(node.Nodes);
                }
            }
        };
        
        findHighestKey(menuState.Nodes);
        const newKey = (highestKey + 1).toString();
        console.error(`Generated new key: ${newKey} for new navigation link`);
        
        // Create the new node
        const newNode = {
            NodeType: 0,
            Key: newKey,
            Title: linkData.Title,
            SimpleUrl: linkData.Url,
            FriendlyUrlSegment: "",
            IsDeleted: false
        };
        
        // If a parent key is provided, add it as a child node
        if (linkData.ParentKey) {
            console.error(`Adding as child node to parent with key: ${linkData.ParentKey}`);
            
            const findAndAddToParent = (nodes: any[]): boolean => {
                for (let i = 0; i < nodes.length; i++) {
                    if (nodes[i].Key === linkData.ParentKey) {
                        // Found the parent node, add the new node as a child
                        if (!nodes[i].Nodes) {
                            nodes[i].Nodes = [];
                        }
                        // Add ParentKey explicitly and with type safety
                        const nodeWithParent = {
                            ...newNode,
                            ParentKey: linkData.ParentKey
                        };
                        nodes[i].Nodes.push(nodeWithParent);
                        return true;
                    }
                    
                    // Check child nodes recursively
                    if (nodes[i].Nodes && nodes[i].Nodes.length > 0) {
                        if (findAndAddToParent(nodes[i].Nodes)) {
                            return true;
                        }
                    }
                }
                return false;
            };
            
            const parentFound = findAndAddToParent(menuState.Nodes);
            if (!parentFound) {
                throw new Error(`Parent node with key ${linkData.ParentKey} not found`);
            }
        } else {
            // Add as a top-level node
            menuState.Nodes.push(newNode);
        }
        
        // Update the version timestamp
        menuState.Version = new Date().toISOString();
        
        // Save the updated menu state
        console.error(`Saving updated menu state with new navigation link...`);
        await request({
            url: `${url}/_api/navigation/SaveMenuState`,
            headers: headers,
            json: true,
            method: 'POST',
            body: { menuState: menuState },
            timeout: 30000
        });
        
        return {
            content: [{
                type: "text",
                text: JSON.stringify({
                    success: true,
                    message: `Navigation link "${linkData.Title}" successfully created in ${navigationType.toLowerCase()} navigation`,
                    newLink: {
                        key: newKey,
                        title: linkData.Title,
                        url: linkData.Url,
                        parentKey: linkData.ParentKey || null,
                        navigationType: navigationType,
                        accessMethod: "menustate"
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

        console.error("Error in addNavigationLink tool:", errorMessage);

        return {
            content: [{
                type: "text",
                text: `Error adding navigation link: ${errorMessage}`
            }],
            isError: true
        } as IToolResult;
    }
}

export default addNavigationLink;

