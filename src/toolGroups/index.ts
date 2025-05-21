// src/toolGroups/index.ts
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { SharePointConfig } from '../config';
import { registerListTools } from './listTools';
import { registerSiteTools } from './siteTools';
import { registerViewTools } from './viewTools';
import { registerPageTools } from './pageTools';
import { registerFieldTools } from './fieldTools';
import { registerBatchTools } from './batchTools';
import { registerContentTypeTools } from './contentTypeTools';
import { registerSiteContentTypeTools } from './siteContentTypeTools';
import { registerRegionalAndFeaturesTools } from './regionalAndFeaturesTools';
import { registerSearchTools } from './searchTools';

/**
 * Register all tool groups with the server
 * @param server The MCP server instance
 * @param config SharePoint configuration
 */
export function registerAllToolGroups(server: McpServer, config: SharePointConfig): void {
    // Register list management tools
    registerListTools(server, config);
    
    // Register site management tools
    registerSiteTools(server, config);
    
    // Register view management tools
    registerViewTools(server, config);
    
    // Register page management tools
    registerPageTools(server, config);
    
    // Register field management tools
    registerFieldTools(server, config);
    
    // Register batch operation tools
    registerBatchTools(server, config);
    
    // Register content type management tools
    registerContentTypeTools(server, config);
    
    // Register site content type management tools
    registerSiteContentTypeTools(server, config);
    
    // Register regional settings and features management tools
    registerRegionalAndFeaturesTools(server, config);
    
    // Register search tools
    registerSearchTools(server, config);
}
