// src/toolGroups/siteContentTypeTools.ts
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { SharePointConfig } from '../config';

// Import site content type-related tools
import {
    getSiteContentTypes,
    getSiteContentType,
    updateSiteContentType,
    deleteSiteContentType,
    // Types
    GetSiteContentTypesParams,
    GetSiteContentTypeParams,
    UpdateSiteContentTypeParams,
    DeleteSiteContentTypeParams
} from '../tools';

/**
 * Register site content type management tools with the MCP server
 * @param server The MCP server instance
 * @param config SharePoint configuration
 * @returns void
 */
export function registerSiteContentTypeTools(server: McpServer, config: SharePointConfig): void {
    console.error("Registering site content type management tools...");

    // Add getSiteContentTypes tool
    server.tool(
        "getSiteContentTypes",
        "Get all content types from a SharePoint site",
        {
            url: z.string().url().describe("URL of the SharePoint website")
        },
        async (params: GetSiteContentTypesParams) => {
            return await getSiteContentTypes(params, config);
        }
    );

    // Add getSiteContentType tool
    server.tool(
        "getSiteContentType",
        "Get a specific content type from a SharePoint site",
        {
            url: z.string().url().describe("URL of the SharePoint website"),
            contentTypeId: z.string().describe("ID of the content type to retrieve")
        },
        async (params: GetSiteContentTypeParams) => {
            return await getSiteContentType(params, config);
        }
    );

    // createSiteContentType tool removed as it was buggy

    // Add updateSiteContentType tool
    server.tool(
        "updateSiteContentType",
        "Update a content type in a SharePoint site",
        {
            url: z.string().url().describe("URL of the SharePoint website"),
            contentTypeId: z.string().describe("ID of the content type to update"),
            updateData: z.object({
                Name: z.string().optional().describe("New name for the content type"),
                Description: z.string().optional().describe("New description for the content type"),
                Group: z.string().optional().describe("New group for the content type")
            }).describe("Properties to update in the content type")
        },
        async (params: UpdateSiteContentTypeParams) => {
            return await updateSiteContentType(params, config);
        }
    );

    // Add deleteSiteContentType tool
    server.tool(
        "deleteSiteContentType",
        "Delete a content type from a SharePoint site",
        {
            url: z.string().url().describe("URL of the SharePoint website"),
            contentTypeId: z.string().describe("ID of the content type to delete"),
            confirmation: z.string().describe("Confirmation string that must match the content type ID exactly")
        },
        async (params: DeleteSiteContentTypeParams) => {
            return await deleteSiteContentType(params, config);
        }
    );
}
