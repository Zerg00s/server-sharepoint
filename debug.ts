#!/usr/bin/env node

// Special debug script for testing batch operations in SharePoint
// Usage examples:
// npx ts-node debug.ts batchCreate
// npx ts-node debug.ts batchUpdate
// npx ts-node debug.ts batchDelete

import * as dotenv from 'dotenv';
import getSharePointConfig from './src/config';
import {
  getSite,
  getLists,
  getListItems,
  createListItem,
  batchCreateListItems,
  batchUpdateListItems,
  batchDeleteListItems,
} from './src/tools';
import { getSharePointHeaders, getRequestDigest } from './src/auth_factory';
import request from 'request-promise';

// Load environment variables
dotenv.config();

/**
 * Main function to test batch operations
 */
async function main() {
  console.log('=== BATCH OPERATIONS DEBUG SCRIPT ===');

  // Parse command line args
  const args = process.argv.slice(2);
  console.log('Command line arguments:', args);
  const command = args[0] || 'batchTest'; // Default to batchTest if no command provided
  console.log(`Command to execute: ${command}`);

  // Get the site URL from command line or environment
  const siteUrl = args[1] || process.env.SHAREPOINT_SITE_URL || '';
  if (!siteUrl) {
    console.error('ERROR: No SharePoint site URL provided.');
    console.error('Usage: npx ts-node debug.ts [command] [siteUrl] [listTitle]');
    console.error('Commands: batchCreate, batchUpdate, batchDelete, batchTest');
    process.exit(1);
  }
  console.log(`Using site URL: ${siteUrl}`);

  // Get list title from args or default
  const listTitle = args[2] || 'English Lessons';
  console.log(`Using list title: ${listTitle}`);

  // Print environment variables for debugging
  console.log('Environment variables:');
  console.log('- AZURE_APPLICATION_ID:', process.env.AZURE_APPLICATION_ID || '(not set)');
  console.log('- AZURE_APPLICATION_CERTIFICATE_THUMBPRINT:', process.env.AZURE_APPLICATION_CERTIFICATE_THUMBPRINT || '(not set)');
  console.log('- AZURE_APPLICATION_CERTIFICATE_PASSWORD:', process.env.AZURE_APPLICATION_CERTIFICATE_PASSWORD ? '(set)' : '(not set)');
  console.log('- M365_TENANT_ID:', process.env.M365_TENANT_ID || '(not set)');
  console.log('- SHAREPOINT_CLIENT_ID:', process.env.SHAREPOINT_CLIENT_ID || '(not set)');
  console.log('- SHAREPOINT_CLIENT_SECRET:', process.env.SHAREPOINT_CLIENT_SECRET ? '(set)' : '(not set)');
  console.log('- SHAREPOINT_SITE_URL:', process.env.SHAREPOINT_SITE_URL || '(not set)');

  // Create a proper config object explicitly to fix TypeScript discrimination issues
  const azureConfig = {
    clientId: process.env.AZURE_APPLICATION_ID || '',
    certificateThumbprint: process.env.AZURE_APPLICATION_CERTIFICATE_THUMBPRINT || '',
    certificatePassword: process.env.AZURE_APPLICATION_CERTIFICATE_PASSWORD || '',
    tenantId: process.env.M365_TENANT_ID || '',
    authType: 'certificate' as const // Use const assertion to ensure proper type
  };

  console.log('Using explicit certificate config for testing.');
  console.log('Certificate config valid:', 
    Boolean(azureConfig.clientId && azureConfig.certificateThumbprint && 
    azureConfig.certificatePassword && azureConfig.tenantId));
  console.log('Configuration loaded:', {
    authMethod: azureConfig.authType,
    clientId: azureConfig.clientId ? '✓ Set' : '✗ Missing',
    tenantId: azureConfig.tenantId ? '✓ Set' : '✗ Missing',
    certificateThumbprint: azureConfig.certificateThumbprint ? '✓ Set' : '✗ Missing',
    certificatePassword: azureConfig.certificatePassword ? '✓ Set' : '✗ Missing',
    siteUrl: siteUrl || '(Using command line URL)',
  });

  try {
    console.log(`\nExecuting "${command}" command...\n`);
    
    // First get site details to verify connection
    console.log('Verifying connection to site...');
    const siteResult = await getSite({ url: siteUrl }, azureConfig);
    console.log(`Connected to site: ${JSON.parse(siteResult.content[0].text as string).title}`);
    
    // Log auth details and test request digest
    console.log('\nTesting authorization and request digest...');
    try {
      const headers = await getSharePointHeaders(siteUrl, azureConfig);
      console.log('Auth headers obtained successfully.');
      
      const digest = await getRequestDigest(siteUrl, headers);
      console.log('Request digest obtained successfully:', digest ? '✓ Valid' : '✗ Invalid');
      
      // Test direct REST API call with auth
      console.log('\nTesting direct REST API call...');
      const directResult = await request({
        url: `${siteUrl}/_api/web/lists/getByTitle('${encodeURIComponent(listTitle)}')`,
        headers: { ...headers, 'Accept': 'application/json;odata=verbose' },
        json: true,
        method: 'GET',
        timeout: 15000,
        resolveWithFullResponse: true,
        simple: false
      });
      
      console.log('Direct API call response status:', directResult.statusCode);
      if (directResult.statusCode === 200) {
        console.log('List exists, direct API call successful.');
      } else if (directResult.statusCode === 404) {
        console.log('List does not exist. Will attempt to create it for testing.');
        
        // Create a test list
        await createTestList(siteUrl, listTitle, headers, digest);
      } else {
        console.log('Unexpected response from API.');
        console.error('Response:', directResult.body);
      }
    } catch (error) {
      console.error('Error testing authentication:', error);
    }
    
    // Execute the requested command
    let result;
    switch (command) {
      case 'batchCreate':
        console.log('\nTesting batch creation...');
        result = await batchCreateListItems({
          url: siteUrl,
          listTitle,
          items: [
            { Title: `Test Item 1 - ${new Date().toISOString()}` },
            { Title: `Test Item 2 - ${new Date().toISOString()}` }
          ]
        }, azureConfig);
        break;
        
      case 'batchUpdate':
        console.log('\nGetting items to update...');
        const itemsResult = await getListItems({ url: siteUrl, listTitle }, azureConfig);
        const items = JSON.parse(itemsResult.content[0].text as string).items;
        
        if (items && items.length >= 2) {
          console.log(`Found ${items.length} items, will update the first 2...`);
          result = await batchUpdateListItems({
            url: siteUrl,
            listTitle,
            items: [
              { id: items[0].ID, data: { Title: `Updated Item 1 - ${new Date().toISOString()}` } },
              { id: items[1].ID, data: { Title: `Updated Item 2 - ${new Date().toISOString()}` } }
            ]
          }, azureConfig);
        } else {
          console.log('Not enough items to update. Please run batchCreate first.');
        }
        break;
        
      case 'batchDelete':
        console.log('\nGetting items to delete...');
        const deleteItemsResult = await getListItems({ url: siteUrl, listTitle }, azureConfig);
        const deleteItems = JSON.parse(deleteItemsResult.content[0].text as string).items;
        
        if (deleteItems && deleteItems.length > 0) {
          console.log(`Found ${deleteItems.length} items, will delete them all...`);
          result = await batchDeleteListItems({
            url: siteUrl,
            listTitle,
            itemIds: deleteItems.map((item: any) => item.ID)
          }, azureConfig);
        } else {
          console.log('No items to delete. Please run batchCreate first.');
        }
        break;
        
      case 'batchTest':
        // This command will create items, update them, and then delete them
        console.log('\n=== COMPREHENSIVE BATCH TEST ===');
        
        console.log('\nStep 1: Creating items...');
        // Use batch create
        const createResult = await batchCreateListItems({
          url: siteUrl,
          listTitle,
          items: [
            { Title: `Batch Test Item 1 - ${new Date().toISOString()}` },
            { Title: `Batch Test Item 2 - ${new Date().toISOString()}` }
          ]
        }, azureConfig);
        
        console.log('Create result:', createResult.content[0].text);
        const createdItems = JSON.parse(createResult.content[0].text as string).createdItems;
        
        if (createdItems && createdItems.length > 0) {
          console.log('\nStep 2: Updating items...');
          // Use batch update
          const updateResult = await batchUpdateListItems({
            url: siteUrl,
            listTitle,
            items: createdItems.map((item: any, index: number) => ({
              id: item.id,
              data: { Title: `Updated Batch Item ${index + 1} - ${new Date().toISOString()}` }
            }))
          }, azureConfig);
          
          console.log('Update result:', updateResult.content[0].text);
          
          console.log('\nStep 3: Deleting items...');
          // Use batch delete
          const deleteResult = await batchDeleteListItems({
            url: siteUrl,
            listTitle,
            itemIds: createdItems.map((item: any) => item.id)
          }, azureConfig);
          
          console.log('Delete result:', deleteResult.content[0].text);
          result = deleteResult;
        }
        break;
        
      default:
        console.error(`Unknown command: ${command}`);
        console.error('Valid commands: batchCreate, batchUpdate, batchDelete, batchTest');
        process.exit(1);
    }

    // Display the final result
    console.log('\n===== FINAL RESULT =====');
    if (result && result.content && result.content.length > 0) {
      for (const item of result.content) {
        if (item.type === 'text') {
          console.log(item.text);
        } else {
          console.log(`[Content of type: ${item.type}]`);
        }
      }
    } else {
      console.log('No result content returned');
    }
    console.log('=================\n');
  } catch (error) {
    console.error('Error executing command:', error);
    process.exit(1);
  }
}

/**
 * Helper function to create a test list for batch operations
 */
async function createTestList(
  siteUrl: string,
  listTitle: string,
  headers: Record<string, string>,
  digest: string
): Promise<void> {
  try {
    console.log(`Creating test list "${listTitle}"...`);
    
    const createHeaders = {
      ...headers,
      'Accept': 'application/json;odata=verbose',
      'Content-Type': 'application/json;odata=verbose',
      'X-RequestDigest': digest
    };
    
    const listData = {
      __metadata: { type: 'SP.List' },
      Title: listTitle,
      Description: 'Test list for batch operations',
      AllowContentTypes: true,
      BaseTemplate: 100,
      ContentTypesEnabled: true
    };
    
    const createResult = await request({
      url: `${siteUrl}/_api/web/lists`,
      headers: createHeaders,
      method: 'POST',
      body: listData,
      json: true,
      timeout: 15000
    });
    
    console.log('Test list created successfully.');
  } catch (error) {
    console.error('Error creating test list:', error);
    throw error;
  }
}

// Execute the main function with proper error handling
console.log('Starting batch operations debug script...');
main()
  .then(() => console.log('Debug completed successfully'))
  .catch(error => {
    console.error('Unhandled error in main execution:', error);
    process.exit(1);
  });
