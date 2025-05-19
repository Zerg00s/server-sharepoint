#!/usr/bin/env node

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
// Import from config without .js extension
import getSharePointConfig, { validateConfig } from './config';
import { registerTools } from './toolRegistry';

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
    registerTools(server, config);
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
