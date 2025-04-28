#!/usr/bin/env node

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
// Import from config without .js extension
import getSharePointConfig, { validateConfig } from './config';
import { 
    getTitle, 
    getLists, 
    getListItems, 
    addMockData,
    GetTitleParams,
    GetListsParams,
    GetListItemsParams,
    AddMockDataParams
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

    // Add getTitle tool
    server.tool(
        "getTitle",
        "Get the title of a SharePoint website",
        {
            url: z.string().url().describe("URL of the SharePoint website")
        },
        async (params: GetTitleParams) => {
            return await getTitle(params, config);
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

    // Add addMockData tool
    server.tool(
        "addMockData",
        "Add mock data items to a specific SharePoint list",
        {
            url: z.string().url().describe("URL of the SharePoint website"),
            listTitle: z.string().describe("Title of the SharePoint list to add mock data to"),
            itemCount: z.number().int().min(1).max(100).describe("Number of mock items to create (1-100)")
        },
        async (params: AddMockDataParams) => {
            return await addMockData(params, config);
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