// src/tools/getGroupMembers.ts
import request from 'request-promise';
import { IToolResult, ISharePointGroupMember } from '../interfaces';
import { getSharePointHeaders, getRequestDigest } from '../auth_factory';
import { SharePointConfig } from '../config';

export interface GetGroupMembersParams {
    url: string;
    groupId: number;
}

/**
 * Get members of a specific SharePoint group
 * @param params Parameters including site URL and group ID
 * @param config SharePoint configuration
 * @returns Tool result with group members data
 */
export async function getGroupMembers(
    params: GetGroupMembersParams,
    config: SharePointConfig
): Promise<IToolResult> {
    const { url, groupId } = params;
    console.error(`getGroupMembers tool called with URL: ${url}, Group ID: ${groupId}`);

    try {
        // Authenticate with SharePoint
        const headers = await getSharePointHeaders(url, config);
        console.error("Headers prepared with authentication");

        // First, verify the group exists and get its details
        console.error(`Verifying group ID ${groupId} exists...`);
        let groupDetails;
        try {
            groupDetails = await request({
                url: `${url}/_api/web/SiteGroups/GetById(${groupId})`,
                headers: headers,
                json: true,
                method: 'GET',
                timeout: 15000
            });
        } catch (error) {
            throw new Error(`Group with ID ${groupId} not found`);
        }
        
        // Get all members of the group
        console.error(`Getting members of group "${groupDetails.d.Title}"...`);
        const membersResponse = await request({
            url: `${url}/_api/web/SiteGroups/GetById(${groupId})/Users?$select=Id,Title,Email,LoginName,IsSiteAdmin`,
            headers: headers,
            json: true,
            method: 'GET',
            timeout: 30000
        });
        
        // Process group members
        const members = membersResponse.d.results as ISharePointGroupMember[];
        const formattedMembers = members.map(member => ({
            Id: member.Id,
            Title: member.Title,
            Email: member.Email || '',
            LoginName: member.LoginName,
            IsSiteAdmin: member.IsSiteAdmin || false
        }));

        return {
            content: [{
                type: "text",
                text: JSON.stringify({
                    groupId: groupId,
                    groupTitle: groupDetails.d.Title,
                    groupDescription: groupDetails.d.Description || '',
                    memberCount: formattedMembers.length,
                    members: formattedMembers
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

        console.error("Error in getGroupMembers tool:", errorMessage);

        return {
            content: [{
                type: "text",
                text: `Error getting group members: ${errorMessage}`
            }],
            isError: true
        } as IToolResult;
    }
}

export default getGroupMembers;

