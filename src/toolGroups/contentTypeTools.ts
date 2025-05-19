// src/toolGroups/contentTypeTools.ts
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { SharePointConfig } from '../config';

// Import content type-related tools
import {
    getListContentTypes,
    getListContentType,
    createListContentType,
    updateListContentType,
    deleteListContentType,
    // Types
    GetListContentTypesParams,
    GetListContentTypeParams,
    CreateListContentTypeParams,
    UpdateListContentTypeParams,
    DeleteListContentTypeParams
} from '../tools';

/**
 * Register content type management tools with the MCP server
 * @param server The MCP server instance
 * @param config SharePoint configuration
 * @returns void
 */
export function registerContentTypeTools(server: McpServer, config: SharePointConfig): void {
    console.error("Registering content type management tools...");

    // Add getListContentTypes tool
    server.tool(
        "getListContentTypes",
        "Get all content types from a specific SharePoint list",
        {
            url: z.string().url().describe("URL of the SharePoint website"),
            listTitle: z.string().describe("Title of the SharePoint list to retrieve content types from")
        },
        async (params: GetListContentTypesParams) => {
            return await getListContentTypes(params, config);
        }
    );

    // Add getListContentType tool
    server.tool(
        "getListContentType",
        "Get a specific content type from a SharePoint list",
        {
            url: z.string().url().describe("URL of the SharePoint website"),
            listTitle: z.string().describe("Title of the SharePoint list to retrieve content type from"),
            contentTypeId: z.string().describe("ID of the content type to retrieve")
        },
        async (params: GetListContentTypeParams) => {
            return await getListContentType(params, config);
        }
    );

    // Add createListContentType tool
    server.tool(
        "createListContentType",
        "Create a new content type in a SharePoint list",
        {
            url: z.string().url().describe("URL of the SharePoint website"),
            listTitle: z.string().describe("Title of the SharePoint list"),
            contentTypeData: z.object({
                Name: z.string().describe("Name of the new content type"),
                Description: z.string().optional().describe("Description for the new content type"),
                ParentContentTypeId: z.string().optional().describe("ID of the parent content type"),
                Group: z.string().optional().describe("Group for the new content type")
            }).describe("Properties for the new content type")
        },
        async (params: CreateListContentTypeParams) => {
            return await createListContentType(params, config);
        }
    );

    // Add updateListContentType tool
    server.tool(
        "updateListContentType",
        "Update a content type in a SharePoint list",
        {
            url: z.string().url().describe("URL of the SharePoint website"),
            listTitle: z.string().describe("Title of the SharePoint list"),
            contentTypeId: z.string().describe("ID of the content type to update"),
            updateData: z.object({
                Name: z.string().optional().describe("New name for the content type"),
                Description: z.string().optional().describe("New description for the content type"),
                Group: z.string().optional().describe("New group for the content type")
            }).describe("Properties to update in the content type")
        },
        async (params: UpdateListContentTypeParams) => {
            return await updateListContentType(params, config);
        }
    );

    // Add deleteListContentType tool
    server.tool(
        "deleteListContentType",
        "Delete a content type from a SharePoint list",
        {
            url: z.string().url().describe("URL of the SharePoint website"),
            listTitle: z.string().describe("Title of the SharePoint list"),
            contentTypeId: z.string().describe("ID of the content type to delete"),
            confirmation: z.string().describe("Confirmation string that must match the content type ID exactly")
        },
        async (params: DeleteListContentTypeParams) => {
            return await deleteListContentType(params, config);
        }
    );
}
