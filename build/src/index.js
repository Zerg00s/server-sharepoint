#!/usr/bin/env node
"use strict";
var __createBinding = (this && this.__createBinding) || (Object.create ? (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    var desc = Object.getOwnPropertyDescriptor(m, k);
    if (!desc || ("get" in desc ? !m.__esModule : desc.writable || desc.configurable)) {
      desc = { enumerable: true, get: function() { return m[k]; } };
    }
    Object.defineProperty(o, k2, desc);
}) : (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    o[k2] = m[k];
}));
var __setModuleDefault = (this && this.__setModuleDefault) || (Object.create ? (function(o, v) {
    Object.defineProperty(o, "default", { enumerable: true, value: v });
}) : function(o, v) {
    o["default"] = v;
});
var __importStar = (this && this.__importStar) || (function () {
    var ownKeys = function(o) {
        ownKeys = Object.getOwnPropertyNames || function (o) {
            var ar = [];
            for (var k in o) if (Object.prototype.hasOwnProperty.call(o, k)) ar[ar.length] = k;
            return ar;
        };
        return ownKeys(o);
    };
    return function (mod) {
        if (mod && mod.__esModule) return mod;
        var result = {};
        if (mod != null) for (var k = ownKeys(mod), i = 0; i < k.length; i++) if (k[i] !== "default") __createBinding(result, mod, k[i]);
        __setModuleDefault(result, mod);
        return result;
    };
})();
Object.defineProperty(exports, "__esModule", { value: true });
const mcp_js_1 = require("@modelcontextprotocol/sdk/server/mcp.js");
const stdio_js_1 = require("@modelcontextprotocol/sdk/server/stdio.js");
const zod_1 = require("zod");
// Import from config without .js extension
const config_1 = __importStar(require("./config"));
const tools_1 = require("./tools");
// Get the SharePoint configuration
const config = (0, config_1.default)();
// Create an MCP server
const server = new mcp_js_1.McpServer({
    name: "SharePoint MCP",
    version: "1.0.0",
    capabilities: {
        tools: {},
    }
});
// Only register tools if we have valid credentials
if ((0, config_1.validateConfig)(config)) {
    console.error("✅ SharePoint credentials are valid. Registering tools...");
    // Add getTitle tool
    server.tool("getTitle", "Get the title of a SharePoint website", {
        url: zod_1.z.string().url().describe("URL of the SharePoint website")
    }, async (params) => {
        return await (0, tools_1.getTitle)(params, config);
    });
    // Add getLists tool
    server.tool("getLists", "Get the list of SharePoint lists along with their Titles, URLs, ItemCounts, last modified date, description and base templateID", {
        url: zod_1.z.string().url().describe("URL of the SharePoint website")
    }, async (params) => {
        return await (0, tools_1.getLists)(params, config);
    });
    // Add getListItems tool
    server.tool("getListItems", "Get all items from a specific SharePoint list identified by site URL and list title", {
        url: zod_1.z.string().url().describe("URL of the SharePoint website"),
        listTitle: zod_1.z.string().describe("Title of the SharePoint list to retrieve items from")
    }, async (params) => {
        return await (0, tools_1.getListItems)(params, config);
    });
    // Add addMockData tool
    server.tool("addMockData", "Add mock data items to a specific SharePoint list", {
        url: zod_1.z.string().url().describe("URL of the SharePoint website"),
        listTitle: zod_1.z.string().describe("Title of the SharePoint list to add mock data to"),
        itemCount: zod_1.z.number().int().min(1).max(100).describe("Number of mock items to create (1-100)")
    }, async (params) => {
        return await (0, tools_1.addMockData)(params, config);
    });
}
else {
    console.error("❌ SharePoint credentials are invalid. No tools will be registered.");
}
// Start server function
async function main() {
    try {
        console.error("Starting SharePoint MCP Server...");
        const transport = new stdio_js_1.StdioServerTransport();
        await server.connect(transport);
        console.error("SharePoint MCP Server running on stdio");
    }
    catch (error) {
        console.error("Error starting server:", error);
        process.exit(1);
    }
}
// Run the server
main();
