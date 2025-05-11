// src/tools/deleteNavigationLink.ts
import request from 'request-promise';
import { IToolResult } from '../interfaces';
import { getSharePointHeaders, getRequestDigest } from '../auth_factory';
import { SharePointConfig } from '../config';

export interface DeleteNavigationLinkParams {
    url: string;
    linkKey: string;
    navigationType: 'Global' | 'Quick'; // Global = top navigation, Quick = left navigation
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
 * Delete a navigation link from a SharePoint site using MenuState/SaveMenuState API
 * @param params Parameters including site URL, link key, and navigation type
 * @param config SharePoint configuration
 * @returns Tool result with deletion status
 */
export async function deleteNavigationLink(
    params: DeleteNavigationLinkParams,
    config: SharePointConfig
): Promise<IToolResult> {
    const { url, linkKey, navigationType } = params;
    console.error(`deleteNavigationLink tool called with URL: ${url}, Link Key: ${linkKey}, Navigation Type: ${navigationType}`);

    try {
        // Authenticate with SharePoint
        const headers = await getSharePointHeaders(url, config);
        console.error("Headers prepared with authentication");

        // Get request digest for POST operations
        headers['X-RequestDigest'] = await getRequestDigest(url, headers);
        headers['Content-Type'] = 'application/json;odata=verbose';
        console.error("Headers prepared with request digest");
        
        // Try the direct approach first
        try {
            console.error(`Trying direct API approach for deleting ${navigationType.toLowerCase()} navigation link...`);
            
            // Convert linkKey string to number (SharePoint Navigation API expects ID as number)
            const nodeId = parseInt(linkKey, 10);
            if (isNaN(nodeId)) {
                throw new Error("Link key must be a valid integer ID for direct API");
            }
            
            // Get the node details first for the response
            let nodeDetails;
            try {
                nodeDetails = await request({
                    url: `${url}/_api/web/Navigation/GetNodeById(${nodeId})`,
                    headers: { ...headers, 'Content-Type': undefined },
                    json: true,
                    method: 'GET',
                    timeout: 15000
                });
            }
            catch (detailsError) {
                console.error(`Could not get node details: ${detailsError instanceof Error ? detailsError.message : String(detailsError)}`);
                // Continue with deletion anyway
            }
            
            // Delete the navigation node
            console.error(`Deleting navigation node with ID ${nodeId}...`);
            await request({
                url: `${url}/_api/web/Navigation/GetNodeById(${nodeId})`,
                headers: {
                    ...headers,
                    'X-HTTP-Method': 'DELETE',
                    'IF-MATCH': '*'
                },
                method: 'POST',
                timeout: 30000
            });
            
            console.error("Successfully deleted navigation link using direct API");
            
            // Prepare node info for the response
            const nodeInfo = {
                key: linkKey,
                title: nodeDetails?.d?.Title || '(unknown)',
                url: nodeDetails?.d?.Url || '',
                navigationType: navigationType,
                accessMethod: "direct"
            };
            
            return {
                content: [{
                    type: "text",
                    text: JSON.stringify({
                        success: true,
                        message: `Navigation link "${nodeInfo.title}" successfully deleted from ${navigationType.toLowerCase()} navigation`,
                        deletedLink: nodeInfo
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
        
        // Find the navigation node to delete (mark as deleted)
        let nodeFound = false;
        let deletedNodeData = {
            Key: '',
            Title: '',
            SimpleUrl: ''
        } as NavigationNodeData;
        
        const findAndMarkAsDeleted = (nodeArray: NavigationNodeData[]): boolean => {
            for (let i = 0; i < nodeArray.length; i++) {
                if (nodeArray[i].Key === linkKey) {
                    // Found the node to mark as deleted
                    deletedNodeData = { ...nodeArray[i] } as NavigationNodeData;
                    
                    // Mark as deleted instead of removing
                    nodeArray[i].IsDeleted = true;
                    
                    return true;
                }
                
                // Check child nodes recursively if they exist
                const childNodes = nodeArray[i].Nodes;
                if (childNodes && Array.isArray(childNodes) && childNodes.length > 0) {
                    // Cast child nodes to the proper type
                    const typedChildNodes = childNodes as NavigationNodeData[];
                    if (findAndMarkAsDeleted(typedChildNodes)) {
                        return true;
                    }
                }
            }
            return false;
        };
        
        nodeFound = findAndMarkAsDeleted(menuNodes);
        
        if (!nodeFound) {
            throw new Error(`Navigation link with key ${linkKey} not found in ${navigationType.toLowerCase()} navigation`);
        }
        
        // Update the version timestamp
        menuState.Version = new Date().toISOString();
        
        // Save the updated menu state
        console.error(`Saving updated menu state with deleted navigation link...`);
        await request({
            url: `${url}/_api/navigation/SaveMenuState`,
            headers: headers,
            json: true,
            method: 'POST',
            body: { menuState: menuState },
            timeout: 30000
        });
        
        // Explicitly create safe strings to avoid TS issues
        const nodeTitle = String(deletedNodeData.Title);
        const nodeKey = String(deletedNodeData.Key);
        const nodeUrl = String(deletedNodeData.SimpleUrl);
                
        return {
            content: [{
                type: "text",
                text: JSON.stringify({
                    success: true,
                    message: `Navigation link "${nodeTitle}" successfully deleted from ${navigationType.toLowerCase()} navigation`,
                    deletedLink: {
                        key: nodeKey,
                        title: nodeTitle,
                        url: nodeUrl,
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

        console.error("Error in deleteNavigationLink tool:", errorMessage);

        return {
            content: [{
                type: "text",
                text: `Error deleting navigation link: ${errorMessage}`
            }],
            isError: true
        } as IToolResult;
    }
}

export default deleteNavigationLink;

