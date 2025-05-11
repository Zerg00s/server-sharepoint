// src/tools/getSiteGroups.ts
import request from 'request-promise';
import { IToolResult, ISharePointGroup } from '../interfaces';
import { getSharePointHeaders } from '../auth';
import { SharePointConfig } from '../config';

export interface GetSiteGroupsParams {
    url: string;
}

/**
 * Get all SharePoint groups for a site
 * @param params Parameters including site URL
 * @param config SharePoint configuration
 * @returns Tool result with groups data
 */
export async function getSiteGroups(
    params: GetSiteGroupsParams,
    config: SharePointConfig
): Promise<IToolResult> {
    const { url } = params;
    console.error(`getSiteGroups tool called with URL: ${url}`);

    try {
        // Authenticate with SharePoint
        const headers = await getSharePointHeaders(url, config);
        console.error("Headers prepared with authentication");

        // Get all groups for the site
        console.error(`Getting groups for site...`);
        const groupsResponse = await request({
            url: `${url}/_api/web/SiteGroups?$select=Id,Title,Description,OwnerTitle,AllowMembersEditMembership,OnlyAllowMembersViewMembership`,
            headers: headers,
            json: true,
            method: 'GET',
            timeout: 30000
        });
        
        // Get associated groups to mark them
        const ownerGroupResponse = await request({
            url: `${url}/_api/web/AssociatedOwnerGroup?$select=Id`,
            headers: headers,
            json: true,
            method: 'GET',
            timeout: 15000
        });
        
        const memberGroupResponse = await request({
            url: `${url}/_api/web/AssociatedMemberGroup?$select=Id`,
            headers: headers,
            json: true,
            method: 'GET',
            timeout: 15000
        });
        
        const visitorGroupResponse = await request({
            url: `${url}/_api/web/AssociatedVisitorGroup?$select=Id`,
            headers: headers,
            json: true,
            method: 'GET',
            timeout: 15000
        });
        
        const ownerGroupId = ownerGroupResponse.d.Id;
        const memberGroupId = memberGroupResponse.d.Id;
        const visitorGroupId = visitorGroupResponse.d.Id;
        
        // Process groups with additional information
        const formattedGroups = groupsResponse.d.results.map((group: ISharePointGroup) => {
            let groupType = "Custom";
            
            if (group.Id === ownerGroupId) {
                groupType = "Owners";
            } else if (group.Id === memberGroupId) {
                groupType = "Members";
            } else if (group.Id === visitorGroupId) {
                groupType = "Visitors";
            }
            
            return {
                Id: group.Id,
                Title: group.Title,
                Description: group.Description || '',
                OwnerTitle: group.OwnerTitle || '',
                GroupType: groupType,
                AllowMembersEditMembership: group.AllowMembersEditMembership || false,
                OnlyAllowMembersViewMembership: group.OnlyAllowMembersViewMembership || false
            };
        });

        return {
            content: [{
                type: "text",
                text: JSON.stringify({
                    siteUrl: url,
                    groups: {
                        total: formattedGroups.length,
                        items: formattedGroups
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

        console.error("Error in getSiteGroups tool:", errorMessage);

        return {
            content: [{
                type: "text",
                text: `Error getting site groups: ${errorMessage}`
            }],
            isError: true
        } as IToolResult;
    }
}

export default getSiteGroups;
