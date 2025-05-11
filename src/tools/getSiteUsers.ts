// src/tools/getSiteUsers.ts
import request from 'request-promise';
import { IToolResult, ISharePointUser } from '../interfaces';
import { getSharePointHeaders, getRequestDigest } from '../auth_factory';
import { SharePointConfig } from '../config';

export interface GetSiteUsersParams {
    url: string;
    role?: "All" | "Owners" | "Members" | "Visitors";
}

/**
 * Get users from a SharePoint site, optionally filtered by role
 * @param params Parameters including site URL and optional role filter
 * @param config SharePoint configuration
 * @returns Tool result with users data
 */
export async function getSiteUsers(
    params: GetSiteUsersParams,
    config: SharePointConfig
): Promise<IToolResult> {
    const { url, role = "All" } = params;
    console.error(`getSiteUsers tool called with URL: ${url}, Role Filter: ${role}`);

    try {
        // Authenticate with SharePoint
        const headers = await getSharePointHeaders(url, config);
        console.error("Headers prepared with authentication");

        // Get all users from the site
        console.error(`Getting users for site...`);
        const usersResponse = await request({
            url: `${url}/_api/web/SiteUsers?$select=Id,Title,Email,LoginName,IsSiteAdmin`,
            headers: headers,
            json: true,
            method: 'GET',
            timeout: 30000
        });
        
        // Get site groups to filter by role if needed
        const groupsResponse = await request({
            url: `${url}/_api/web/AssociatedOwnerGroup`,
            headers: headers,
            json: true,
            method: 'GET',
            timeout: 15000
        });
        
        const ownersGroupId = groupsResponse.d.Id;
        
        // Get Members group
        const membersGroupResponse = await request({
            url: `${url}/_api/web/AssociatedMemberGroup`,
            headers: headers,
            json: true,
            method: 'GET',
            timeout: 15000
        });
        
        const membersGroupId = membersGroupResponse.d.Id;
        
        // Get Visitors group
        const visitorsGroupResponse = await request({
            url: `${url}/_api/web/AssociatedVisitorGroup`,
            headers: headers,
            json: true,
            method: 'GET',
            timeout: 15000
        });
        
        const visitorsGroupId = visitorsGroupResponse.d.Id;
        
        // Get group members for Owners
        const ownersResponse = await request({
            url: `${url}/_api/web/GetGroupById(${ownersGroupId})/Users`,
            headers: headers,
            json: true,
            method: 'GET',
            timeout: 30000
        });
        
        // Get group members for Members
        const membersResponse = await request({
            url: `${url}/_api/web/GetGroupById(${membersGroupId})/Users`,
            headers: headers,
            json: true,
            method: 'GET',
            timeout: 30000
        });
        
        // Get group members for Visitors
        const visitorsResponse = await request({
            url: `${url}/_api/web/GetGroupById(${visitorsGroupId})/Users`,
            headers: headers,
            json: true,
            method: 'GET',
            timeout: 30000
        });
        
        // Create sets of user IDs by role for efficient lookup
        const ownerIds = new Set(ownersResponse.d.results.map((user: ISharePointUser) => user.Id));
        const memberIds = new Set(membersResponse.d.results.map((user: ISharePointUser) => user.Id));
        const visitorIds = new Set(visitorsResponse.d.results.map((user: ISharePointUser) => user.Id));
        
        // Process all users and categorize by role
        const allUsers = usersResponse.d.results;
        
        const owners: ISharePointUser[] = [];
        const members: ISharePointUser[] = [];
        const visitors: ISharePointUser[] = [];
        const others: ISharePointUser[] = [];
        
        allUsers.forEach((user: ISharePointUser) => {
            // Skip system accounts
            if (user.LoginName && (
                user.LoginName.includes('SP_FARM') || 
                user.LoginName.includes('SHAREPOINT\\') ||
                user.LoginName.includes('NT AUTHORITY\\') ||
                user.LoginName.includes('APP@SharePoint')
            )) {
                return;
            }
            
            const formattedUser = {
                Id: user.Id,
                Title: user.Title,
                Email: user.Email || '',
                LoginName: user.LoginName,
                IsSiteAdmin: user.IsSiteAdmin || false
            };
            
            if (ownerIds.has(user.Id)) {
                owners.push(formattedUser);
            } else if (memberIds.has(user.Id)) {
                members.push(formattedUser);
            } else if (visitorIds.has(user.Id)) {
                visitors.push(formattedUser);
            } else {
                others.push(formattedUser);
            }
        });
        
        // Filter results based on requested role
        let resultUsers: ISharePointUser[] = [];
        
        switch (role) {
            case "Owners":
                resultUsers = owners;
                break;
            case "Members":
                resultUsers = members;
                break;
            case "Visitors":
                resultUsers = visitors;
                break;
            case "All":
            default:
                resultUsers = [...owners, ...members, ...visitors, ...others];
                break;
        }
        
        return {
            content: [{
                type: "text",
                text: JSON.stringify({
                    siteUrl: url,
                    roleFilter: role,
                    users: {
                        total: resultUsers.length,
                        owners: role === "All" ? owners.length : undefined,
                        members: role === "All" ? members.length : undefined,
                        visitors: role === "All" ? visitors.length : undefined,
                        others: role === "All" ? others.length : undefined,
                        items: resultUsers
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

