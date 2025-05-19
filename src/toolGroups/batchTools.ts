// src/toolGroups/batchTools.ts
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { SharePointConfig } from '../config';

// Import batch operation tools
import {
    batchCreateListItems,
    batchUpdateListItems,
    batchDeleteListItems,
    // Types
    BatchCreateListItemsParams,
    BatchUpdateListItemsParams,
    BatchDeleteListItemsParams
} from '../tools';

/**
 * Register batch operation tools with the MCP server
 * @param server The MCP server instance
 * @param config SharePoint configuration
 * @returns void
 */
export function registerBatchTools(server: McpServer, config: SharePointConfig): void {
    console.error("Registering batch operation tools...");

    // Add batchCreateListItems tool
    server.tool(
        "batchCreateListItems",
        "Create multiple items in a SharePoint list using a single batch request",
        {
            url: z.string().url().describe("URL of the SharePoint website"),
            listTitle: z.string().describe("Title of the SharePoint list"),
            items: z.array(z.record(z.any())).describe("Array of objects containing field names and values for the new items")
        },
        async (params: BatchCreateListItemsParams) => {
            return await batchCreateListItems(params, config);
        }
    );

    // Add batchUpdateListItems tool
    server.tool(
        "batchUpdateListItems",
        "Update multiple items in a SharePoint list using a single batch request",
        {
            url: z.string().url().describe("URL of the SharePoint website"),
            listTitle: z.string().describe("Title of the SharePoint list"),
            items: z.array(
                z.object({
                    id: z.number().int().positive().describe("ID of the item to update"),
                    data: z.record(z.any()).describe("Key-value pairs of field names and values to update")
                })
            ).describe("Array of objects containing item IDs and update data")
        },
        async (params: BatchUpdateListItemsParams) => {
            return await batchUpdateListItems(params, config);
        }
    );

    // Add batchDeleteListItems tool
    server.tool(
        "batchDeleteListItems",
        "Delete multiple items from a SharePoint list using a single batch request",
        {
            url: z.string().url().describe("URL of the SharePoint website"),
            listTitle: z.string().describe("Title of the SharePoint list"),
            itemIds: z.array(z.number().int().positive()).describe("Array of item IDs to delete")
        },
        async (params: BatchDeleteListItemsParams) => {
            return await batchDeleteListItems(params, config);
        }
    );
}
