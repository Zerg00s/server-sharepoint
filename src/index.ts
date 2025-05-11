#!/usr/bin/env node

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
// Import from config without .js extension
import getSharePointConfig, { validateConfig } from './config';
import { 
    getSite, 
    getLists, 
    getListItems, 
    getListFields,
    updateListField,
    updateListItem,
    createListItem,
    createList,
    createListView,
    updateListView,
    deleteListItem,
    getSiteUsers,
    getSiteGroups,
    addGroupMember,
    removeGroupMember,
    getListViews,
    deleteListView,
    deleteList,
    createListField,
    deleteListField,
    getGroupMembers,
    getGlobalNavigationLinks,
    getQuickNavigationLinks,
    getSubsites,
    deleteSubsite,
    updateSite,
    addNavigationLink,
    updateNavigationLink,
    deleteNavigationLink,
    // New view field management tools
    getViewFields,
    addViewField,
    removeViewField,
    removeAllViewFields,
    moveViewFieldTo,
    GetSiteParams,
    GetListsParams,
    GetListItemsParams,
    GetListFieldsParams,
    UpdateListFieldParams,
    UpdateListItemParams,
    CreateListItemParams,
    CreateListParams,
    CreateListViewParams,
    UpdateListViewParams,
    DeleteListItemParams,
    GetSiteUsersParams,
    GetSiteGroupsParams,
    AddGroupMemberParams,
    RemoveGroupMemberParams,
    GetListViewsParams,
    DeleteListViewParams,
    DeleteListParams,
    CreateListFieldParams,
    DeleteListFieldParams,
    GetGroupMembersParams,
    GetGlobalNavigationLinksParams,
    GetQuickNavigationLinksParams,
    GetSubsitesParams,
    DeleteSubsiteParams,
    UpdateSiteParams,
    AddNavigationLinkParams,
    UpdateNavigationLinkParams,
    DeleteNavigationLinkParams,
    // New view field management tool params
    GetViewFieldsParams,
    AddViewFieldParams,
    RemoveViewFieldParams,
    RemoveAllViewFieldsParams,
    MoveViewFieldToParams
// Import from tools without .js extension
} from './tools';

// Get the SharePoint configuration
const config = getSharePointConfig();

// Create an MCP server
const server = new McpServer({
    name: "SharePoint MCP",
    version: "1.0.0",
    capabilities: {
        tools: {},
    }
});

// Only register tools if we have valid credentials
if (validateConfig(config)) {
    console.error("✅ SharePoint credentials are valid. Registering tools...");

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

    // Add getLists tool
    server.tool(
        "getLists",
        "Get the list of SharePoint lists along with their Titles, URLs, ItemCounts, last modified date, description and base templateID",
        {
            url: z.string().url().describe("URL of the SharePoint website")
        },
        async (params: GetListsParams) => {
            return await getLists(params, config);
        }
    );

    // Add getListItems tool
    server.tool(
        "getListItems",
        "Get all items from a specific SharePoint list identified by site URL and list title",
        {
            url: z.string().url().describe("URL of the SharePoint website"),
            listTitle: z.string().describe("Title of the SharePoint list to retrieve items from")
        },
        async (params: GetListItemsParams) => {
            return await getListItems(params, config);
        }
    );

    // Add getListFields tool
    server.tool(
        "getListFields",
        "Get detailed information about fields/columns in a SharePoint list",
        {
            url: z.string().url().describe("URL of the SharePoint website"),
            listTitle: z.string().describe("Title of the SharePoint list to retrieve fields from")
        },
        async (params: GetListFieldsParams) => {
            return await getListFields(params, config);
        }
    );

    // Add createListField tool
    server.tool(
        "createListField",
        "Create a new field (column) in a SharePoint list",
        {
            url: z.string().url().describe("URL of the SharePoint website"),
            listTitle: z.string().describe("Title of the SharePoint list"),
            fieldData: z.object({
                Title: z.string().describe("Display name for the field (can contain spaces)"),
                CleanName: z.string().optional().describe("Clean name without spaces (used for internal name generation)"),
                FieldTypeKind: z.number().int().describe(
                  "Field type value: 0=Invalid, 1=Integer, 2=Text, 3=Note, 4=DateTime, 5=Choice, 6=Lookup, " +
                  "7=Boolean (according to docs, but may not work), 8=Boolean, 9=Number, 10=Currency, 11=URL, " +
                  "15=MultiChoice, 17=Calculated, 19=User"),
                Required: z.boolean().optional().describe("Whether the field is required"),
                EnforceUniqueValues: z.boolean().optional().describe("Whether the field must have unique values"),
                StaticName: z.string().optional().describe("Static name, if not provided will be generated from CleanName or Title"),
                Description: z.string().optional().describe("Description for the field"),
                Choices: z.array(z.string()).optional().describe("For choice fields (FieldTypeKind=5) or MultiChoice fields (FieldTypeKind=15)"),
                DefaultValue: z.union([z.string(), z.number(), z.boolean()]).optional().describe("Default value for the field")
            }).passthrough().describe("Properties for the new field")
        },
        async (params: CreateListFieldParams) => {
            return await createListField(params, config);
        }
    );

    // Add updateListField tool
    server.tool(
        "updateListField",
        "Update a field/column in a SharePoint list including display name, choices, etc.",
        {
            url: z.string().url().describe("URL of the SharePoint website"),
            listTitle: z.string().describe("Title of the SharePoint list"),
            fieldInternalName: z.string().describe("Internal name of the field to update"),
            updateData: z.object({
                Title: z.string().optional().describe("New display name for the field"),
                Description: z.string().optional().describe("New description for the field"),
                Required: z.boolean().optional().describe("Whether the field is required"),
                EnforceUniqueValues: z.boolean().optional().describe("Whether the field must have unique values"),
                Choices: z.array(z.string()).optional().describe("New choices for choice fields"),
                DefaultValue: z.string().optional().describe("New default value for the field")
            }).describe("Field properties to update")
        },
        async (params: UpdateListFieldParams) => {
            return await updateListField(params, config);
        }
    );

    // Add updateListItem tool
    server.tool(
        "updateListItem",
        "Update an item in a SharePoint list",
        {
            url: z.string().url().describe("URL of the SharePoint website"),
            listTitle: z.string().describe("Title of the SharePoint list"),
            itemId: z.number().int().positive().describe("ID of the item to update"),
            itemData: z.record(z.any()).describe("Key-value pairs of field names and their values to update")
        },
        async (params: UpdateListItemParams) => {
            return await updateListItem(params, config);
        }
    );

    // Add createListItem tool
    server.tool(
        "createListItem",
        "Create a new item in a SharePoint list with specified field values",
        {
            url: z.string().url().describe("URL of the SharePoint website"),
            listTitle: z.string().describe("Title of the SharePoint list"),
            itemData: z.record(z.any()).describe("Key-value pairs of field names and their values for the new item")
        },
        async (params: CreateListItemParams) => {
            return await createListItem(params, config);
        }
    );

    // Add createList tool
    server.tool(
        "createList",
        "Create a new SharePoint list or document library",
        {
            url: z.string().url().describe("URL of the SharePoint website"),
            listData: z.object({
                Title: z.string().describe("Title of the new list"),
                Description: z.string().optional().describe("Description for the new list"),
                TemplateType: z.number().int().optional().describe("Template type (100 for generic list, 101 for document library)"),
                Url: z.string().optional().describe("Relative URL for the list (used in browser URLs)"),
                ContentTypesEnabled: z.boolean().optional().describe("Whether to enable content types"),
                AllowContentTypes: z.boolean().optional().describe("Whether to allow content types"),
                EnableVersioning: z.boolean().optional().describe("Whether to enable versioning"),
                EnableMinorVersions: z.boolean().optional().describe("Whether to enable minor versions (for document libraries)"),
                EnableModeration: z.boolean().optional().describe("Whether to enable content approval")
            }).describe("Properties for the new list")
        },
        async (params: CreateListParams) => {
            return await createList(params, config);
        }
    );

    // Add deleteList tool
    server.tool(
        "deleteList",
        "Delete a SharePoint list or document library",
        {
            url: z.string().url().describe("URL of the SharePoint website"),
            listTitle: z.string().describe("Title of the SharePoint list to delete"),
            confirmation: z.string().describe("Confirmation string that must match the list title exactly")
        },
        async (params: DeleteListParams) => {
            return await deleteList(params, config);
        }
    );

    // Add createListView tool
    server.tool(
        "createListView",
        "Create a new view for a SharePoint list with specified fields and settings",
        {
            url: z.string().url().describe("URL of the SharePoint website"),
            listTitle: z.string().describe("Title of the SharePoint list"),
            viewData: z.object({
                Title: z.string().describe("Title of the new view"),
                ViewQuery: z.string().optional().describe("CAML query for filtering items"),
                RowLimit: z.number().int().optional().describe("Maximum number of items to display per page"),
                ViewFields: z.array(z.string()).optional().describe("Array of field internal names to include in the view"),
                PersonalView: z.boolean().optional().describe("Whether this is a personal view"),
                SetAsDefaultView: z.boolean().optional().describe("Whether to set this as the default view")
            }).describe("Properties for the new view")
        },
        async (params: CreateListViewParams) => {
            return await createListView(params, config);
        }
    );

    // Add updateListView tool
    server.tool(
        "updateListView",
        "Update an existing view for a SharePoint list",
        {
            url: z.string().url().describe("URL of the SharePoint website"),
            listTitle: z.string().describe("Title of the SharePoint list"),
            viewTitle: z.string().describe("Title of the view to update"),
            updateData: z.object({
                Title: z.string().optional().describe("New title for the view"),
                ViewQuery: z.string().optional().describe("New CAML query for filtering items"),
                RowLimit: z.number().int().optional().describe("New maximum number of items to display per page"),
                ViewFields: z.array(z.string()).optional().describe("New array of field internal names to include in the view"),
                PersonalView: z.boolean().optional().describe("Whether this is a personal view"),
                SetAsDefaultView: z.boolean().optional().describe("Whether to set this as the default view")
            }).describe("Properties to update in the view"),
            appendFields: z.boolean().optional().describe("Whether to append fields to existing ones instead of replacing them")
        },
        async (params: UpdateListViewParams) => {
            return await updateListView(params, config);
        }
    );

    // Add getListViews tool
    server.tool(
        "getListViews",
        "Get all views from a SharePoint list with optional field details",
        {
            url: z.string().url().describe("URL of the SharePoint website"),
            listTitle: z.string().describe("Title of the SharePoint list to retrieve views from"),
            includeFields: z.boolean().optional().describe("Whether to include the fields for each view"),
            includeHidden: z.boolean().optional().describe("Whether to include hidden views")
        },
        async (params: GetListViewsParams) => {
            return await getListViews(params, config);
        }
    );
    
    // Add deleteListView tool
    server.tool(
        "deleteListView",
        "Delete a view from a SharePoint list",
        {
            url: z.string().url().describe("URL of the SharePoint website"),
            listTitle: z.string().describe("Title of the SharePoint list"),
            viewTitle: z.string().describe("Title of the view to delete")
        },
        async (params: DeleteListViewParams) => {
            return await deleteListView(params, config);
        }
    );

    // Add deleteListItem tool
    server.tool(
        "deleteListItem",
        "Delete an item from a SharePoint list",
        {
            url: z.string().url().describe("URL of the SharePoint website"),
            listTitle: z.string().describe("Title of the SharePoint list"),
            itemId: z.number().int().positive().describe("ID of the item to delete")
        },
        async (params: DeleteListItemParams) => {
            return await deleteListItem(params, config);
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

    // Add deleteListField tool
    server.tool(
        "deleteListField",
        "Delete a field (column) from a SharePoint list",
        {
            url: z.string().url().describe("URL of the SharePoint website"),
            listTitle: z.string().describe("Title of the SharePoint list"),
            fieldInternalName: z.string().describe("Internal name of the field to delete"),
            confirmation: z.string().describe("Confirmation string that must match the field internal name exactly")
        },
        async (params: DeleteListFieldParams) => {
            return await deleteListField(params, config);
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

    // Add getViewFields tool
    server.tool(
        "getViewFields",
        "Get all fields from a specific SharePoint list view",
        {
            url: z.string().url().describe("URL of the SharePoint website"),
            listTitle: z.string().describe("Title of the SharePoint list"),
            viewTitle: z.string().describe("Title of the view to get fields from")
        },
        async (params: GetViewFieldsParams) => {
            return await getViewFields(params, config);
        }
    );

    // Add addViewField tool
    server.tool(
        "addViewField",
        "Add a field to a SharePoint list view",
        {
            url: z.string().url().describe("URL of the SharePoint website"),
            listTitle: z.string().describe("Title of the SharePoint list"),
            viewTitle: z.string().describe("Title of the view to modify"),
            fieldName: z.string().describe("Internal name of the field to add to the view")
        },
        async (params: AddViewFieldParams) => {
            return await addViewField(params, config);
        }
    );

    // Add removeViewField tool
    server.tool(
        "removeViewField",
        "Remove a field from a SharePoint list view",
        {
            url: z.string().url().describe("URL of the SharePoint website"),
            listTitle: z.string().describe("Title of the SharePoint list"),
            viewTitle: z.string().describe("Title of the view to modify"),
            fieldName: z.string().describe("Internal name of the field to remove from the view")
        },
        async (params: RemoveViewFieldParams) => {
            return await removeViewField(params, config);
        }
    );

    // Add removeAllViewFields tool
    server.tool(
        "removeAllViewFields",
        "Remove all fields from a SharePoint list view",
        {
            url: z.string().url().describe("URL of the SharePoint website"),
            listTitle: z.string().describe("Title of the SharePoint list"),
            viewTitle: z.string().describe("Title of the view to modify")
        },
        async (params: RemoveAllViewFieldsParams) => {
            return await removeAllViewFields(params, config);
        }
    );

    // Add moveViewFieldTo tool
    server.tool(
        "moveViewFieldTo",
        "Move a field to a specific position in a SharePoint list view",
        {
            url: z.string().url().describe("URL of the SharePoint website"),
            listTitle: z.string().describe("Title of the SharePoint list"),
            viewTitle: z.string().describe("Title of the view to modify"),
            fieldName: z.string().describe("Internal name of the field to move"),
            index: z.number().int().min(0).describe("New position index (0-based)")
        },
        async (params: MoveViewFieldToParams) => {
            return await moveViewFieldTo(params, config);
        }
    );
} else {
    console.error("❌ SharePoint credentials are invalid. No tools will be registered.");
}

// Start server function
async function main() {
    try {
        console.error("Starting SharePoint MCP Server...");
        const transport = new StdioServerTransport();
        await server.connect(transport);
        console.error("SharePoint MCP Server running on stdio");
    } catch (error) {
        console.error("Error starting server:", error);
        process.exit(1);
    }
}

// Run the server
main();