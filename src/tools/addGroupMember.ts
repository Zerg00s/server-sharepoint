// src/tools/addGroupMember.ts
import request from 'request-promise';
import { IToolResult } from '../interfaces';
import { getSharePointHeaders, getRequestDigest } from '../auth';
import { SharePointConfig } from '../config';

export interface AddGroupMemberParams {
    url: string;
    groupId: number;
    loginName: string; // User principal name or login name (e.g. "i:0#.f|membership|user@domain.com")
}

/**
 * Add a user to a SharePoint group
 * @param params Parameters including site URL, group ID, and user login name
 * @param config SharePoint configuration
 * @returns Tool result with addition status
 */
export async function addGroupMember(
    params: AddGroupMemberParams,
    config: SharePointConfig
): Promise<IToolResult> {
    const { url, groupId, loginName } = params;
    console.error(`addGroupMember tool called with URL: ${url}, Group ID: ${groupId}, Login: ${loginName}`);

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
        
        // Encode the login name to handle special characters
        const encodedLoginName = encodeURIComponent(loginName);
        
        // Add the user to the group
        console.error(`Adding user "${loginName}" to group "${groupDetails.d.Title}"...`);
        try {
            // Use the direct approach with the correct endpoint format
            await request({
                url: `${url}/_api/web/sitegroups(${groupId})/users`,
                headers: { 
                    ...headers, 
                    'Content-Type': 'application/json',
                    'accept': 'application/json;odata.metadata=none'
                },
                json: true,
                method: 'POST',
                body: {
                    'LoginName': loginName
                },
                timeout: 30000
            });
        } catch (error) {
            // Check if the error is because the user is already a member
            const errorMessage = error instanceof Error ? error.message : String(error);
            if (errorMessage.includes("already exists") || errorMessage.includes("already a member")) {
                return {
                    content: [{
                        type: "text",
                        text: JSON.stringify({
                            success: false,
                            message: `User "${loginName}" is already a member of group "${groupDetails.d.Title}"`,
                            group: {
                                id: groupId,
                                title: groupDetails.d.Title
                            }
                        }, null, 2)
                    }]
                } as IToolResult;
            } else {
                throw error;
            }
        }
        
        // Get the newly added user to include in the response
        console.error(`Getting user information...`);
        let addedUserDetails: {
            id?: number;
            title: string;
            email?: string; 
            loginName: string;
        } = {
            title: loginName, // Default if we can't get the actual user details
            loginName: loginName
        };
        
        try {
            // Try to get the user information using the login name
            const userResponse = await request({
                url: `${url}/_api/web/SiteUsers/GetByLoginName(@v)?@v='${encodedLoginName}'`,
                headers: { ...headers, 'Content-Type': undefined },
                json: true,
                method: 'GET',
                timeout: 15000
            });
            
            if (userResponse && userResponse.d) {
                addedUserDetails = {
                    id: userResponse.d.Id,
                    title: userResponse.d.Title,
                    email: userResponse.d.Email || '',
                    loginName: userResponse.d.LoginName
                };
            }
        } catch (userError) {
            console.error(`Warning: Could not retrieve added user details: ${userError instanceof Error ? userError.message : String(userError)}`);
            // Continue with basic user information
        }
        
        return {
            content: [{
                type: "text",
                text: JSON.stringify({
                    success: true,
                    message: `User successfully added to group "${groupDetails.d.Title}"`,
                    addedUser: addedUserDetails,
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

        console.error("Error in addGroupMember tool:", errorMessage);

        return {
            content: [{
                type: "text",
                text: `Error adding group member: ${errorMessage}`
            }],
            isError: true
        } as IToolResult;
    }
}

export default addGroupMember;
