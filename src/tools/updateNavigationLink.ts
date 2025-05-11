// src/tools/updateNavigationLink.ts
import request from 'request-promise';
import { IToolResult } from '../interfaces';
import { getSharePointHeaders, getRequestDigest } from '../auth';
import { SharePointConfig } from '../config';

export interface UpdateNavigationLinkParams {
    url: string;
    linkKey: string;
    navigationType: 'Global' | 'Quick'; // Global = top navigation, Quick = left navigation
    updateData: {
        Title?: string;
        Url?: string;
        IsExternal?: boolean;
    };
}

// Interface for navigation node data
interface NavigationNodeData {
    Key: string;
    Title: string;
    SimpleUrl: string;
    IsDeleted?: boolean;
    Nodes?: NavigationNodeData[];
    [key: string]: any; // For any other properties
}

/**
 * Update a navigation link in a SharePoint site using MenuState/SaveMenuState API
 * @param params Parameters including site URL, link key, navigation type, and update data
 * @param config SharePoint configuration
 * @returns Tool result with update status
 */
export async function updateNavigationLink(
    params: UpdateNavigationLinkParams,
    config: SharePointConfig
): Promise<IToolResult> {
    const { url, linkKey, navigationType, updateData } = params;
    console.error(`updateNavigationLink tool called with URL: ${url}, Link Key: ${linkKey}, Navigation Type: ${navigationType}`);

    try {
        // Validate input
        if (Object.keys(updateData).length === 0) {
            throw new Error("No link properties provided for update");
        }

        // Authenticate with SharePoint
        const headers = await getSharePointHeaders(url, config);
        console.error("Headers prepared with authentication");

        // Get request digest for POST operations
        headers['X-RequestDigest'] = await getRequestDigest(url, headers);
        headers['Content-Type'] = 'application/json;odata=verbose';
        console.error("Headers prepared with request digest");
        
        // Try the direct approach first
        try {
            console.error(`Trying direct API approach for updating ${navigationType.toLowerCase()} navigation link...`);
            
            // Convert linkKey string to number (SharePoint Navigation API expects ID as number)
            const nodeId = parseInt(linkKey, 10);
            if (isNaN(nodeId)) {
                throw new Error("Link key must be a valid integer ID for direct API");
            }
            
            // Prepare the update data
            const updatePayload: any = {
                __metadata: { type: 'SP.NavigationNode' }
            };
            
            if (updateData.Title !== undefined) {
                updatePayload.Title = updateData.Title;
            }
            
            if (updateData.Url !== undefined) {
                updatePayload.Url = updateData.Url;
            }
            
            if (updateData.IsExternal !== undefined) {
                updatePayload.IsExternal = updateData.IsExternal;
            }
            
            // Update using standard API
            console.error(`Updating navigation node with ID ${nodeId}...`);
            await request({
                url: `${url}/_api/web/Navigation/GetNodeById(${nodeId})`,
                headers: {
                    ...headers,
                    'X-HTTP-Method': 'MERGE',
                    'IF-MATCH': '*'
                },
                json: true,
                method: 'POST',
                body: updatePayload,
                timeout: 30000
            });
            
            console.error("Successfully updated navigation link using direct API");
            
            // Get the node to return updated information
            const updatedNode = await request({
                url: `${url}/_api/web/Navigation/GetNodeById(${nodeId})`,
                headers: { ...headers, 'Content-Type': undefined },
                json: true,
                method: 'GET',
                timeout: 15000
            });
            
            return {
                content: [{
                    type: "text",
                    text: JSON.stringify({
                        success: true,
                        message: `Navigation link successfully updated in ${navigationType.toLowerCase()} navigation`,
                        updatedLink: {
                            id: updatedNode.d.Id,
                            title: updatedNode.d.Title,
                            url: updatedNode.d.Url,
                            isExternal: updatedNode.d.IsExternal,
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
        
        // Get the current menu state
        console.error(`Getting current ${navigationType.toLowerCase()} navigation menu state...`);
        const menuStateResponse = await request({
            url: `${url}/_api/navigation/MenuState?mapProviderName='${providerName}'`,
            headers: { ...headers, 'Content-Type': undefined },
            json: true,
            method: 'GET',
            timeout: 30000
        });
        
        // Type-safe handling of menuState
        const menuState = menuStateResponse.MenuState;
        
        // Initialize the nodes with a fallback to an empty array
        // This ensures we don't have undefined issues
        const menuNodes: NavigationNodeData[] = Array.isArray(menuState.Nodes) ? 
            menuState.Nodes.map((node: any) => node as NavigationNodeData) : [];
            
        console.error(`Retrieved menu state with ${menuNodes.length} nodes`);
        
        // Find the navigation node to update
        let nodeFound = false;
        let originalNodeData = {
            Key: '',
            Title: '',
            SimpleUrl: ''
        } as NavigationNodeData;
        
        const findAndUpdateNode = (nodeArray: NavigationNodeData[]): boolean => {
            for (let i = 0; i < nodeArray.length; i++) {
                if (nodeArray[i].Key === linkKey) {
                    // Found the node to update
                    originalNodeData = { ...nodeArray[i] } as NavigationNodeData;
                    
                    // Update properties
                    if (updateData.Title !== undefined) {
                        nodeArray[i].Title = updateData.Title;
                    }
                    
                    if (updateData.Url !== undefined) {
                        nodeArray[i].SimpleUrl = updateData.Url;
                    }
                    
                    return true;
                }
                
                // Check child nodes recursively if they exist
                const childNodes = nodeArray[i].Nodes;
                if (childNodes && Array.isArray(childNodes) && childNodes.length > 0) {
                    // Cast child nodes to the proper type
                    const typedChildNodes = childNodes as NavigationNodeData[];
                    if (findAndUpdateNode(typedChildNodes)) {
                        return true;
                    }
                }
            }
            return false;
        };
        
        nodeFound = findAndUpdateNode(menuNodes);
        
        if (!nodeFound) {
            throw new Error(`Navigation link with key ${linkKey} not found in ${navigationType.toLowerCase()} navigation`);
        }
        
        // Update the version timestamp
        menuState.Version = new Date().toISOString();
        
        // Save the updated menu state
        console.error(`Saving updated menu state with modified navigation link...`);
        await request({
            url: `${url}/_api/navigation/SaveMenuState`,
            headers: headers,
            json: true,
            method: 'POST',
            body: { menuState: menuState },
            timeout: 30000
        });
        
        // Explicitly create safe strings to avoid TS issues
        const nodeKey = String(originalNodeData.Key);
        const nodeTitle = String(originalNodeData.Title);
        const nodeUrl = String(originalNodeData.SimpleUrl);
        
        // Use the extracted values for response
        const updatedTitle = updateData.Title !== undefined ? updateData.Title : nodeTitle;
        const updatedUrl = updateData.Url !== undefined ? updateData.Url : nodeUrl;
        
        return {
            content: [{
                type: "text",
                text: JSON.stringify({
                    success: true,
                    message: `Navigation link successfully updated in ${navigationType.toLowerCase()} navigation`,
                    originalLink: {
                        key: nodeKey,
                        title: nodeTitle,
                        url: nodeUrl
                    },
                    updatedLink: {
                        key: linkKey,
                        title: updatedTitle,
                        url: updatedUrl,
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

        console.error("Error in updateNavigationLink tool:", errorMessage);

        return {
            content: [{
                type: "text",
                text: `Error updating navigation link: ${errorMessage}`
            }],
            isError: true
        } as IToolResult;
    }
}

export default updateNavigationLink;
