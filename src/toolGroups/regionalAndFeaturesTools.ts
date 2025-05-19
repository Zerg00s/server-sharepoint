// src/toolGroups/regionalAndFeaturesTools.ts
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { SharePointConfig } from '../config';

// Import regional settings and features-related tools
import {
    getRegionalSettings,
    getSiteCollectionFeatures,
    getSiteFeatures,
    getSiteFeature,
    // Types
    GetRegionalSettingsParams,
    GetSiteCollectionFeaturesParams,
    GetSiteFeaturesParams,
    GetSiteFeatureParams
} from '../tools';

/**
 * Register regional settings and features management tools with the MCP server
 * @param server The MCP server instance
 * @param config SharePoint configuration
 * @returns void
 */
export function registerRegionalAndFeaturesTools(server: McpServer, config: SharePointConfig): void {
    console.error("Registering regional settings and features management tools...");

    // Add getRegionalSettings tool
    server.tool(
        "getRegionalSettings",
        "Get regional settings from a SharePoint site",
        {
            url: z.string().url().describe("URL of the SharePoint website")
        },
        async (params: GetRegionalSettingsParams) => {
            return await getRegionalSettings(params, config);
        }
    );

    // Add getSiteCollectionFeatures tool
    server.tool(
        "getSiteCollectionFeatures",
        "Get all features from a SharePoint site collection",
        {
            url: z.string().url().describe("URL of the SharePoint website")
        },
        async (params: GetSiteCollectionFeaturesParams) => {
            return await getSiteCollectionFeatures(params, config);
        }
    );

    // Add getSiteFeatures tool
    server.tool(
        "getSiteFeatures",
        "Get all features from a SharePoint site",
        {
            url: z.string().url().describe("URL of the SharePoint website")
        },
        async (params: GetSiteFeaturesParams) => {
            return await getSiteFeatures(params, config);
        }
    );

    // Add getSiteFeature tool
    server.tool(
        "getSiteFeature",
        "Get a specific feature from a SharePoint site by feature ID",
        {
            url: z.string().url().describe("URL of the SharePoint website"),
            featureId: z.string().describe("Feature ID (GUID) of the feature to retrieve")
        },
        async (params: GetSiteFeatureParams) => {
            return await getSiteFeature(params, config);
        }
    );
}
