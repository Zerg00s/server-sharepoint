#!/usr/bin/env node

// Sample debug script for testing SharePoint tools
// npx ts-node debug.ts getSite "https://gocleverpointcom.sharepoint.com/sites/Dashboard-Communication"
// npx ts-node debug.ts getLists "https://gocleverpointcom.sharepoint.com/sites/Dashboard-Communication"
// npx ts-node debug.ts getListItems "https://gocleverpointcom.sharepoint.com/sites/Dashboard-Communication" "Sites Report"


import * as dotenv from 'dotenv';
import getSharePointConfig from './src/config';
import { getSite, getLists, getListItems, addMockData } from './src/tools';

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
  
  const command = args[0] || 'getSite'; // Default to getSite if no command provided
  console.log(`Command to execute: ${command}`);
  
  // Get the site URL from command line or environment
  const siteUrl = args[1] || process.env.SHAREPOINT_SITE_URL || '';
  if (!siteUrl) {
    console.error('ERROR: No SharePoint site URL provided.');
    console.error('Usage: node debug.js [command] [siteUrl] [options]');
    console.error('Commands: getSite, getLists, getListItems, addMockData');
    process.exit(1);
  }
  console.log(`Using site URL: ${siteUrl}`);
  
  // Load SharePoint configuration
  console.log("Loading SharePoint configuration...");
  const config = getSharePointConfig();
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
      case 'getSite':
        console.log("Calling getSite tool function...");
        result = await getSite({ url: siteUrl }, config);
        break;
        
      case 'getLists':
        console.log("Calling getLists tool function...");
        result = await getLists({ url: siteUrl }, config);
        break;
        
      case 'getListItems':
        // Get list title from args or prompt
        const listTitle = args[2] || 'Documents'; // Default to Documents if not provided
        console.log(`Calling getListItems tool function with list "${listTitle}"...`);
        result = await getListItems({ url: siteUrl, listTitle }, config);
        break;
        
      case 'addMockData':
        // Get list title and count from args
        const mockListTitle = args[2] || 'Documents';
        const itemCount = parseInt(args[3] || '5', 10);
        console.log(`Calling addMockData tool function with list "${mockListTitle}" and count ${itemCount}...`);
        result = await addMockData({ 
          url: siteUrl, 
          listTitle: mockListTitle, 
          itemCount 
        }, config);
        break;
        
      default:
        console.error(`Unknown command: ${command}`);
        console.error('Valid commands: getSite, getLists, getListItems, addMockData');
        process.exit(1);
    }
    
    // Display the result
    console.log('\n===== RESULT =====');
    if (result && result.content && result.content.length > 0) {
      for (const item of result.content) {
        if (item.type === 'text') {
          console.log(item.text);
        } else {
          console.log(`[Content of type: ${item.type}]`);
        }
      }
    } else {
      console.log("No result content returned");
    }
    console.log('=================\n');
    
  } catch (error) {
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