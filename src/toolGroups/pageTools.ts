// src/toolGroups/pageTools.ts
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { SharePointConfig } from '../config';

// Import page-related tools
import {
    createModernPage,
    getModernPages,
    getModernPage,
    deleteModernPage,
    // Types
    CreateModernPageParams,
    GetModernPagesParams,
    GetModernPageParams,
    DeleteModernPageParams
} from '../tools';

/**
 * Register page management tools with the MCP server
 * @param server The MCP server instance
 * @param config SharePoint configuration
 * @returns void
 */
export function registerPageTools(server: McpServer, config: SharePointConfig): void {
    console.error("Registering page management tools...");

    // Add createModernPage tool
    server.tool(
        "createModernPage",
        "Create a modern page in SharePoint",
        {
            url: z.string().url().describe("URL of the SharePoint website"),
            title: z.string().describe("Title of the page"),
            fileName: z.string().optional().describe("Optional filename for the page (e.g., 'sample.aspx')"),
            pageLayoutType: z.string().optional().describe("Page layout type (Article, Home, SingleWebPartAppPage, etc.)"),
            description: z.string().optional().describe("Description of the page"),
            thumbnailUrl: z.string().optional().describe("URL for the page thumbnail/banner image"),
            promotedState: z.number().optional().describe("Promotion state: 0=Not promoted, 1=Promoted, 2=Promoted to news"),
            publishPage: z.boolean().optional().describe("Whether to publish the page after creation"),
            content: z.string().optional().describe("HTML content for the page")
        },
        async (params: CreateModernPageParams) => {
            return await createModernPage(params, config);
        }
    );
    
    // Add getModernPages tool
    server.tool(
        "getModernPages",
        "Get modern pages from a SharePoint site",
        {
            url: z.string().url().describe("URL of the SharePoint website"),
            pageTitle: z.string().optional().describe("Optional - filter by page title"),
            limit: z.number().optional().describe("Maximum number of pages to return")
        },
        async (params: GetModernPagesParams) => {
            return await getModernPages(params, config);
        }
    );

    // Add getModernPage tool
    server.tool(
        "getModernPage",
        "Get a specific modern page by ID from a SharePoint site",
        {
            url: z.string().url().describe("URL of the SharePoint website"),
            pageId: z.number().int().positive().describe("ID of the page to retrieve")
        },
        async (params: GetModernPageParams) => {
            return await getModernPage(params, config);
        }
    );

    // Add deleteModernPage tool
    server.tool(
        "deleteModernPage",
        "Delete a modern page from SharePoint",
        {
            url: z.string().url().describe("URL of the SharePoint website"),
            pageId: z.number().int().positive().describe("ID of the page to delete"),
            confirmation: z.string().describe("Confirmation string that must match the page ID")
        },
        async (params: DeleteModernPageParams) => {
            return await deleteModernPage(params, config);
        }
    );
}
