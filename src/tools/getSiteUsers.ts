// src/tools/getSiteUsers.ts
import request from 'request-promise';
import { IToolResult, ISharePointUser } from '../interfaces';
import { getSharePointHeaders } from '../auth_factory';
import { SharePointConfig } from '../config';

export interface GetSiteUsersParams {
    url: string;
    role?: "All" | "Owners" | "Members" | "Visitors";
}

/**
 * Get users from a SharePoint site
 * @param params Parameters including site URL and optional role filter (role filter is ignored in this implementation)
 * @param config SharePoint configuration
 * @returns Tool result with users data
 */
export async function getSiteUsers(
    params: GetSiteUsersParams,
    config: SharePointConfig
): Promise<IToolResult> {
    const { url } = params;
    console.log(`getSiteUsers tool called with URL: ${url}`);

    try {
        // Authenticate with SharePoint
        const headers = await getSharePointHeaders(url, config);
        console.log("Headers prepared with authentication");

        // Get all users from the site with a simple direct API call
        console.log(`Getting users for site...`);
        
        const usersResponse = await request({
            url: `${url}/_api/web/SiteUsers?$select=Id,Title,Email,LoginName,IsSiteAdmin`,
            headers: headers,
            json: true,
            method: 'GET',
            timeout: 30000
        });
        
        // Process all users, filtering out system accounts
        const allUsers = usersResponse.d?.results || [];
        console.log(`Retrieved ${allUsers.length} total users`);
        
        const filteredUsers = allUsers.filter((user: ISharePointUser) => {
            // Skip system accounts
            return !(
                user.LoginName && (
                    user.LoginName.includes('SP_FARM') || 
                    user.LoginName.includes('SHAREPOINT\\') ||
                    user.LoginName.includes('NT AUTHORITY\\') ||
                    user.LoginName.includes('APP@SharePoint')
                )
            );
        }).map((user: ISharePointUser) => {
            // Format user data consistently
            return {
                Id: user.Id,
                Title: user.Title || '',
                Email: user.Email || '',
                LoginName: user.LoginName || '',
                IsSiteAdmin: user.IsSiteAdmin || false
            };
        });
        
        console.log(`Filtered to ${filteredUsers.length} non-system users`);
        
        return {
            content: [{
                type: "text",
                text: JSON.stringify({
                    siteUrl: url,
                    users: {
                        total: filteredUsers.length,
                        items: filteredUsers
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

        console.error("Error in getSiteUsers tool:", errorMessage);

        return {
            content: [{
                type: "text",
                text: `Error getting site users: ${errorMessage}`
            }],
            isError: true
        } as IToolResult;
    }
}

export default getSiteUsers;