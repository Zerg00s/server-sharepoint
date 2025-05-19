// src/toolGroups/siteTools.ts
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { SharePointConfig } from '../config';

// Import site-related tools
import {
    getSite,
    updateSite,
    getSiteUsers,
    getSiteGroups,
    getGroupMembers,
    addGroupMember,
    removeGroupMember,
    getGlobalNavigationLinks,
    getQuickNavigationLinks,
    getSubsites,
    deleteSubsite,
    addNavigationLink,
    updateNavigationLink,
    deleteNavigationLink,
    // Types
    GetSiteParams,
    UpdateSiteParams,
    GetSiteUsersParams,
    GetSiteGroupsParams,
    GetGroupMembersParams,
    AddGroupMemberParams,
    RemoveGroupMemberParams,
    GetGlobalNavigationLinksParams,
    GetQuickNavigationLinksParams,
    GetSubsitesParams,
    DeleteSubsiteParams,
    AddNavigationLinkParams,
    UpdateNavigationLinkParams,
    DeleteNavigationLinkParams
} from '../tools';

/**
 * Register site management tools with the MCP server
 * @param server The MCP server instance
 * @param config SharePoint configuration
 * @returns void
 */
export function registerSiteTools(server: McpServer, config: SharePointConfig): void {
    console.error("Registering site management tools...");

    // Add getSite tool
    server.tool(
        "getSite",
        "Get the title of a SharePoint website",
        {
            url: z.string().url().describe("URL of the SharePoint website")
        },
        async (params: GetSiteParams) => {
            return await getSite(params, config);
        }
    );

    // Add updateSite tool
    server.tool(
        "updateSite",
        "Update a SharePoint site properties (Title, Description, etc.)",
        {
            url: z.string().url().describe("URL of the SharePoint website"),
            siteData: z.object({
                Title: z.string().optional().describe("New title for the site"),
                Description: z.string().optional().describe("New description for the site"),
                LogoUrl: z.string().optional().describe("URL for the site logo")
            }).passthrough().describe("Site properties to update")
        },
        async (params: UpdateSiteParams) => {
            return await updateSite(params, config);
        }
    );

    // Add getSiteUsers tool
    server.tool(
        "getSiteUsers",
        "Get users from a SharePoint site, optionally filtered by role",
        {
            url: z.string().url().describe("URL of the SharePoint website"),
            role: z.enum(["All", "Owners", "Members", "Visitors"]).optional().describe("Role to filter users by")
        },
        async (params: GetSiteUsersParams) => {
            return await getSiteUsers(params, config);
        }
    );

    // Add getSiteGroups tool
    server.tool(
        "getSiteGroups",
        "Get all SharePoint groups for a site",
        {
            url: z.string().url().describe("URL of the SharePoint website")
        },
        async (params: GetSiteGroupsParams) => {
            return await getSiteGroups(params, config);
        }
    );

    // Add getGroupMembers tool
    server.tool(
        "getGroupMembers",
        "Get members of a specific SharePoint group",
        {
            url: z.string().url().describe("URL of the SharePoint website"),
            groupId: z.number().int().positive().describe("ID of the SharePoint group")
        },
        async (params: GetGroupMembersParams) => {
            return await getGroupMembers(params, config);
        }
    );

    // Add addGroupMember tool
    server.tool(
        "addGroupMember",
        "Add a user to a SharePoint group",
        {
            url: z.string().url().describe("URL of the SharePoint website"),
            groupId: z.number().int().positive().describe("ID of the SharePoint group"),
            loginName: z.string().describe("Login name of the user to add")
        },
        async (params: AddGroupMemberParams) => {
            return await addGroupMember(params, config);
        }
    );

    // Add removeGroupMember tool
    server.tool(
        "removeGroupMember",
        "Remove a user from a SharePoint group",
        {
            url: z.string().url().describe("URL of the SharePoint website"),
            groupId: z.number().int().positive().describe("ID of the SharePoint group"),
            loginName: z.string().describe("Login name of the user to remove")
        },
        async (params: RemoveGroupMemberParams) => {
            return await removeGroupMember(params, config);
        }
    );

    // Add getGlobalNavigationLinks tool
    server.tool(
        "getGlobalNavigationLinks",
        "Get global navigation links from a SharePoint site",
        {
            url: z.string().url().describe("URL of the SharePoint website")
        },
        async (params: GetGlobalNavigationLinksParams) => {
            return await getGlobalNavigationLinks(params, config);
        }
    );

    // Add getQuickNavigationLinks tool
    server.tool(
        "getQuickNavigationLinks",
        "Get quick navigation links (left navigation) from a SharePoint site",
        {
            url: z.string().url().describe("URL of the SharePoint website")
        },
        async (params: GetQuickNavigationLinksParams) => {
            return await getQuickNavigationLinks(params, config);
        }
    );

    // Add getSubsites tool
    server.tool(
        "getSubsites",
        "Get all subsites from a SharePoint site",
        {
            url: z.string().url().describe("URL of the SharePoint website")
        },
        async (params: GetSubsitesParams) => {
            return await getSubsites(params, config);
        }
    );

    // Add deleteSubsite tool
    server.tool(
        "deleteSubsite",
        "Delete a SharePoint subsite",
        {
            url: z.string().url().describe("URL of the SharePoint website/subsite to delete"),
            confirmation: z.string().describe("Confirmation string that must match exactly the site name from the URL")
        },
        async (params: DeleteSubsiteParams) => {
            return await deleteSubsite(params, config);
        }
    );
    
    // Add addNavigationLink tool
    server.tool(
        "addNavigationLink",
        "Add a navigation link to a SharePoint site (global or quick navigation)",
        {
            url: z.string().url().describe("URL of the SharePoint website"),
            linkData: z.object({
                Title: z.string().describe("Title for the navigation link"),
                Url: z.string().describe("URL for the navigation link"),
                IsExternal: z.boolean().optional().describe("Whether the link is external to SharePoint"),
                ParentKey: z.string().optional().describe("Key of the parent navigation node for creating child links")
            }).describe("Navigation link properties"),
            navigationType: z.enum(["Global", "Quick"]).describe("Type of navigation: Global (top) or Quick (left)")
        },
        async (params: AddNavigationLinkParams) => {
            return await addNavigationLink(params, config);
        }
    );
    
    // Add updateNavigationLink tool
    server.tool(
        "updateNavigationLink",
        "Update a navigation link in a SharePoint site (global or quick navigation)",
        {
            url: z.string().url().describe("URL of the SharePoint website"),
            linkKey: z.string().describe("Key of the navigation link to update"),
            navigationType: z.enum(["Global", "Quick"]).describe("Type of navigation: Global (top) or Quick (left)"),
            updateData: z.object({
                Title: z.string().optional().describe("New title for the navigation link"),
                Url: z.string().optional().describe("New URL for the navigation link"),
                IsExternal: z.boolean().optional().describe("Whether the link is external to SharePoint")
            }).describe("Navigation link properties to update")
        },
        async (params: UpdateNavigationLinkParams) => {
            return await updateNavigationLink(params, config);
        }
    );
    
    // Add deleteNavigationLink tool
    server.tool(
        "deleteNavigationLink",
        "Delete a navigation link from a SharePoint site (global or quick navigation)",
        {
            url: z.string().url().describe("URL of the SharePoint website"),
            linkKey: z.string().describe("Key of the navigation link to delete"),
            navigationType: z.enum(["Global", "Quick"]).describe("Type of navigation: Global (top) or Quick (left)")
        },
        async (params: DeleteNavigationLinkParams) => {
            return await deleteNavigationLink(params, config);
        }
    );
}
