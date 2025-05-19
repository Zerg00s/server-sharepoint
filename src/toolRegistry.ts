// src/toolRegistry.ts
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { SharePointConfig } from './config';
import { registerAllToolGroups } from './toolGroups';

/**
 * Register all SharePoint tools with the MCP server
 * @param server The MCP server instance
 * @param config SharePoint configuration
 * @returns void
 */
export function registerTools(server: McpServer, config: SharePointConfig): void {
    console.error("✅ SharePoint credentials are valid. Registering tools...");

    // Register all tool groups
    registerAllToolGroups(server, config);
}
