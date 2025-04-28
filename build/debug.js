#!/usr/bin/env node
"use strict";
// Sample debug script for testing SharePoint tools
// npx ts-node debug.ts getTitle "https://gocleverpointcom.sharepoint.com/sites/Dashboard-Communication"
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
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
const dotenv = __importStar(require("dotenv"));
const config_1 = __importDefault(require("./src/config"));
const tools_1 = require("./src/tools");
// Load environment variables
dotenv.config();
/**
 * Main function to test individual tool functions
 */
async function main() {
    console.log("Debug script starting...");
    // Parse command line args
    const args = process.argv.slice(2);
    console.log("Command line arguments:", args);
    const command = args[0] || 'getTitle'; // Default to getTitle if no command provided
    console.log(`Command to execute: ${command}`);
    // Get the site URL from command line or environment
    const siteUrl = args[1] || process.env.SHAREPOINT_SITE_URL || '';
    if (!siteUrl) {
        console.error('ERROR: No SharePoint site URL provided.');
        console.error('Usage: node debug.js [command] [siteUrl] [options]');
        console.error('Commands: getTitle, getLists, getListItems, addMockData');
        process.exit(1);
    }
    console.log(`Using site URL: ${siteUrl}`);
    // Load SharePoint configuration
    console.log("Loading SharePoint configuration...");
    const config = (0, config_1.default)();
    console.log("Configuration loaded:", {
        clientId: config.clientId ? "✓ Set" : "✗ Missing",
        clientSecret: config.clientSecret ? "✓ Set" : "✗ Missing",
        tenantId: config.tenantId ? "✓ Set" : "✗ Missing",
        siteUrl: config.siteUrl || "(Using command line URL)"
    });
    try {
        console.log(`Executing "${command}" command...`);
        let result;
        switch (command) {
            case 'getTitle':
                console.log("Calling getTitle tool function...");
                result = await (0, tools_1.getTitle)({ url: siteUrl }, config);
                break;
            case 'getLists':
                console.log("Calling getLists tool function...");
                result = await (0, tools_1.getLists)({ url: siteUrl }, config);
                break;
            case 'getListItems':
                // Get list title from args or prompt
                const listTitle = args[2] || 'Documents'; // Default to Documents if not provided
                console.log(`Calling getListItems tool function with list "${listTitle}"...`);
                result = await (0, tools_1.getListItems)({ url: siteUrl, listTitle }, config);
                break;
            case 'addMockData':
                // Get list title and count from args
                const mockListTitle = args[2] || 'Documents';
                const itemCount = parseInt(args[3] || '5', 10);
                console.log(`Calling addMockData tool function with list "${mockListTitle}" and count ${itemCount}...`);
                result = await (0, tools_1.addMockData)({
                    url: siteUrl,
                    listTitle: mockListTitle,
                    itemCount
                }, config);
                break;
            default:
                console.error(`Unknown command: ${command}`);
                console.error('Valid commands: getTitle, getLists, getListItems, addMockData');
                process.exit(1);
        }
        // Display the result
        console.log('\n===== RESULT =====');
        if (result && result.content && result.content.length > 0) {
            for (const item of result.content) {
                if (item.type === 'text') {
                    console.log(item.text);
                }
                else {
                    console.log(`[Content of type: ${item.type}]`);
                }
            }
        }
        else {
            console.log("No result content returned");
        }
        console.log('=================\n');
    }
    catch (error) {
        console.error("Error executing command:", error);
        process.exit(1);
    }
}
// Execute the main function with proper error handling
console.log("Starting debug script execution...");
main()
    .then(() => console.log('Debug completed successfully'))
    .catch(error => {
    console.error('Unhandled error in main execution:', error);
    process.exit(1);
});
