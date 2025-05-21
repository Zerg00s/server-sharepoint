// src/toolGroups/searchTools.ts
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { SharePointConfig } from '../config';

// Import search-related tools
import {
    searchSharePointSite,
    // Types
    SearchSharePointSiteParams
} from '../tools';

/**
 * Register search tools with the MCP server
 * @param server The MCP server instance
 * @param config SharePoint configuration
 * @returns void
 */
export function registerSearchTools(server: McpServer, config: SharePointConfig): void {
    console.error("Registering search tools...");

    // Add searchSharePointSite tool
    server.tool(
        "searchSharePointSite",
        "Search within a SharePoint site using KQL query",
        {
            url: z.string().url().describe("URL of the SharePoint website"),
            query: z.string().describe("KQL (Keyword Query Language) search query"),
            rowLimit: z.number().int().positive().optional().describe("Maximum number of results to return (default: 50)"),
            startRow: z.number().int().min(0).optional().describe("Starting row for pagination (default: 0)"),
            selectProperties: z.array(z.string()).optional().describe("Properties to select in results (optional)"),
            sourceid: z.string().optional().describe("Source ID to limit search scope (optional)")
        },
        async (params: SearchSharePointSiteParams) => {
            return await searchSharePointSite(params, config);
        }
    );
}
