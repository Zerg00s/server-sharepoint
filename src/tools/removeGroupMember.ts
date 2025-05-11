// src/tools/removeGroupMember.ts
import request from 'request-promise';
import { IToolResult, ISharePointGroupMember } from '../interfaces';
import { getSharePointHeaders, getRequestDigest } from '../auth_factory';
import { SharePointConfig } from '../config';

export interface RemoveGroupMemberParams {
    url: string;
    groupId: number;
    loginName: string;
}

/**
 * Remove a user from a SharePoint group
 * @param params Parameters including site URL, group ID, and user login name
 * @param config SharePoint configuration
 * @returns Tool result with removal status
 */
export async function removeGroupMember(
    params: RemoveGroupMemberParams,
    config: SharePointConfig
): Promise<IToolResult> {
    const { url, groupId, loginName } = params;
    console.error(`removeGroupMember tool called with URL: ${url}, Group ID: ${groupId}, Login: ${loginName}`);

    try {
        // Authenticate with SharePoint
        const headers = await getSharePointHeaders(url, config);
        console.error("Headers prepared with authentication");

        // Get request digest for operations
        headers['X-RequestDigest'] = await getRequestDigest(url, headers);
        console.error("Headers prepared with request digest");

        // First, verify the group exists
        console.error(`Verifying group ID ${groupId} exists...`);
        let groupDetails;
        try {
            groupDetails = await request({
                url: `${url}/_api/web/SiteGroups/GetById(${groupId})`,
                headers: { ...headers, 'Content-Type': undefined },
                json: true,
                method: 'GET',
                timeout: 15000
            });
        } catch (error) {
            throw new Error(`Group with ID ${groupId} not found`);
        }
        
        // Get group members to verify the user exists in the group
        console.error(`Getting members of group "${groupDetails.d.Title}"...`);
        const membersResponse = await request({
            url: `${url}/_api/web/SiteGroups/GetById(${groupId})/Users`,
            headers: { ...headers, 'Content-Type': undefined },
            json: true,
            method: 'GET',
            timeout: 30000
        });
        
        // Find the user in the group
        const members = membersResponse.d.results as ISharePointGroupMember[];
        const memberToRemove = members.find(member => member.LoginName === loginName);
        
        if (!memberToRemove) {
            throw new Error(`User with login name "${loginName}" not found in group "${groupDetails.d.Title}"`);
        }
        
        // Encode the login name to handle special characters
        const encodedLoginName = encodeURIComponent(loginName);
        
        // Remove the user from the group
        console.error(`Removing user "${memberToRemove.Title}" from group "${groupDetails.d.Title}"...`);
        await request({
            url: `${url}/_api/web/SiteGroups/GetById(${groupId})/Users/RemoveByLoginName(@v)?@v='${encodedLoginName}'`,
            headers: { ...headers, 'X-HTTP-Method': 'POST' },
            method: 'POST',
            timeout: 30000
        });
        
        return {
            content: [{
                type: "text",
                text: JSON.stringify({
                    success: true,
                    message: `User "${memberToRemove.Title}" successfully removed from group "${groupDetails.d.Title}"`,
                    removedUser: {
                        id: memberToRemove.Id,
                        title: memberToRemove.Title,
                        email: memberToRemove.Email || '',
                        loginName: memberToRemove.LoginName
                    },
                    group: {
                        id: groupId,
                        title: groupDetails.d.Title
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

        console.error("Error in removeGroupMember tool:", errorMessage);

        return {
            content: [{
                type: "text",
                text: `Error removing group member: ${errorMessage}`
            }],
            isError: true
        } as IToolResult;
    }
}

export default removeGroupMember;

