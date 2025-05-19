// src/toolGroups/listTools.ts
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { SharePointConfig } from '../config';

// Import list-related tools
import {
    getLists,
    getListItems,
    getListFields,
    createListItem,
    updateListItem,
    deleteListItem,
    createList,
    deleteList,
    updateList,
    // Types
    GetListsParams,
    GetListItemsParams,
    GetListFieldsParams,
    CreateListItemParams,
    UpdateListItemParams,
    DeleteListItemParams,
    CreateListParams,
    DeleteListParams,
    UpdateListParams
} from '../tools';

/**
 * Register list management tools with the MCP server
 * @param server The MCP server instance
 * @param config SharePoint configuration
 * @returns void
 */
export function registerListTools(server: McpServer, config: SharePointConfig): void {
    console.error("Registering list management tools...");

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
    
    // Add updateList tool
    server.tool(
        "updateList",
        "Update a SharePoint list properties (Title, Description, versioning settings, etc.)",
        {
            url: z.string().url().describe("URL of the SharePoint website"),
            listTitle: z.string().describe("Title of the SharePoint list to update"),
            updateData: z.object({
                Title: z.string().optional().describe("New title for the list"),
                Description: z.string().optional().describe("New description for the list"),
                EnableVersioning: z.boolean().optional().describe("Whether to enable versioning"),
                EnableMinorVersions: z.boolean().optional().describe("Whether to enable minor versions"),
                EnableModeration: z.boolean().optional().describe("Whether to enable content approval"),
                DraftVersionVisibility: z.number().optional().describe("Draft visibility: 0=Reader, 1=Author, 2=Approver"),
                ContentTypesEnabled: z.boolean().optional().describe("Whether to enable content types"),
                Hidden: z.boolean().optional().describe("Whether the list is hidden"),
                Ordered: z.boolean().optional().describe("Whether list items can be manually ordered")
            }).describe("List properties to update")
        },
        async (params: UpdateListParams) => {
            return await updateList(params, config);
        }
    );
}
