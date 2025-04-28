#!/usr/bin/env node

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import * as spauth from 'node-sp-auth';
import request from 'request-promise';
import * as dotenv from 'dotenv';
import { IList, ISharePointListResponse, ISharePointListItem, IFormattedListItem } from './interfaces.js';

dotenv.config();

// 🛠 Simple CLI parser
const args = process.argv.slice(2).reduce((acc: Record<string, string>, arg) => {
    const [key, value] = arg.split('=');
    if (key && value) {
        acc[key.replace(/^--/, '')] = value;
    }
    return acc;
}, {});

// 🧠 First priority: CLI args → fallback: ENV
const clientId = args.clientId || process.env.SHAREPOINT_CLIENT_ID || '';
const secret = args.clientSecret || process.env.SHAREPOINT_CLIENT_SECRET || '';
const tenantID = args.tenantId || process.env.SHAREPOINT_TENANT_ID || '';
const siteUrl = args.siteUrl || process.env.SHAREPOINT_SITE_URL || '';

if (!clientId || !secret || !tenantID) {
    console.error("ERROR: Missing SharePoint credentials!");
    console.error("Provide via environment variables or CLI arguments like:");
    console.error("--clientId=xxx --clientSecret=yyy --tenantId=zzz");
    process.exit(1);
} else {
    console.error("✅ SharePoint credentials loaded.");
}

// Create an MCP server
const server = new McpServer({
    name: "SharePoint MCP",
    version: "1.0.0",
    capabilities: {
        tools: {},
    }
});

// Validate credentials before adding the tool
if (!clientId || !secret || !tenantID) {
    console.error("ERROR: SharePoint credentials not provided in environment variables");
    console.error("Please set SHAREPOINT_CLIENT_ID, SHAREPOINT_CLIENT_SECRET, and SHAREPOINT_TENANT_ID");
} else {
    console.error("SharePoint credentials loaded from environment variables");

    // Add a tool to get the website title
    server.tool(
        "getTitle",
        "Get the title of a SharePoint website",
        {
            url: z.string().url().describe("URL of the SharePoint website")
        },
        async ({ url }) => {
            console.error("getTitle tool called with URL:", url);

            try {
                // Authenticate with SharePoint
                console.error("Authenticating with SharePoint...");
                const authData = await spauth.getAuth(url, {
                    clientId: clientId,
                    clientSecret: secret,
                    realm: tenantID
                });

                // Define headers from auth data
                const headers = { ...authData.headers };
                headers['Accept'] = 'application/json;odata=verbose';
                console.error("Headers prepared:", headers);

                // Make request to SharePoint API
                console.error("Making request to SharePoint API...");
                const response = await request({
                    url: `${url}/_api/web`,
                    headers: headers,
                    json: true,
                    method: 'GET',
                    timeout: 8000
                });

                console.error("SharePoint API response received");
                console.error("SharePoint site title:", response.d.Title);

                return {
                    content: [{
                        type: "text",
                        text: `SharePoint site title: ${response.d.Title}`
                    }]
                };
            } catch (error: unknown) {
                // Type-safe error handling
                let errorMessage: string;

                if (error instanceof Error) {
                    errorMessage = error.message;
                } else if (typeof error === 'string') {
                    errorMessage = error;
                } else {
                    errorMessage = "Unknown error occurred";
                }

                console.error("Error in getTitle tool:", errorMessage);

                return {
                    content: [{
                        type: "text",
                        text: `Error fetching title: ${errorMessage}`
                    }],
                    isError: true
                };
            }
        }
    );

    // Add the getLists tool to get SharePoint lists with their details
    server.tool(
        "getLists",
        "Get the list of SharePoint lists along with their Titles, URLs, ItemCounts, last modified date, description and base templateID",
        {
            url: z.string().url().describe("URL of the SharePoint website")
        },
        async ({ url }) => {
            console.error("getLists tool called with URL:", url);

            try {
                // Authenticate with SharePoint
                console.error("Authenticating with SharePoint...");
                const authData = await spauth.getAuth(url, {
                    clientId: clientId,
                    clientSecret: secret,
                    realm: tenantID
                });

                // Define headers from auth data
                const headers = { ...authData.headers };
                headers['Accept'] = 'application/json;odata=verbose';
                console.error("Headers prepared:", headers);

                // Make request to SharePoint API to get lists
                // Add Hidden and IsSystemList properties to the query
                console.error("Making request to SharePoint API for lists...");
                const response = await request({
                    url: `${url}/_api/web/lists?$select=Title,Id,ItemCount,LastItemModifiedDate,Description,BaseTemplate,Hidden,IsSystemList,RootFolder/ServerRelativeUrl&$expand=RootFolder`,
                    headers: headers,
                    json: true,
                    method: 'GET',
                    timeout: 15000
                });

                console.error(`SharePoint API response received with ${response.d.results.length} total lists`);
                
                // Filter out hidden and system lists
                const visibleLists = response.d.results.filter((list: ISharePointListResponse) => 
                    !list.Hidden && !list.IsSystemList
                );
                
                console.error(`Filtered to ${visibleLists.length} visible lists (excluding hidden and system lists)`);
                
                // Format the list data for display
                const lists: IList[] = visibleLists.map((list: ISharePointListResponse) => {
                    return {
                        Title: list.Title,
                        URL: `${url}${list.RootFolder.ServerRelativeUrl}`,
                        ItemCount: list.ItemCount,
                        LastModified: list.LastItemModifiedDate,
                        Description: list.Description || 'No description',
                        BaseTemplateID: list.BaseTemplate
                    };
                });

                return {
                    content: [{
                        type: "text",
                        text: JSON.stringify(lists, null, 2)
                    }]
                };
            } catch (error: unknown) {
                // Type-safe error handling
                let errorMessage: string;

                if (error instanceof Error) {
                    errorMessage = error.message;
                } else if (typeof error === 'string') {
                    errorMessage = error;
                } else {
                    errorMessage = "Unknown error occurred";
                }

                console.error("Error in getLists tool:", errorMessage);

                return {
                    content: [{
                        type: "text",
                        text: `Error fetching lists: ${errorMessage}`
                    }],
                    isError: true
                };
            }
        }
    );
    
    // Add the getListItems tool to get all items from a specific SharePoint list
    server.tool(
        "getListItems",
        "Get all items from a specific SharePoint list identified by site URL and list title",
        {
            url: z.string().url().describe("URL of the SharePoint website"),
            listTitle: z.string().describe("Title of the SharePoint list to retrieve items from")
        },
        async ({ url, listTitle }) => {
            console.error(`getListItems tool called with URL: ${url}, List Title: ${listTitle}`);

            try {
                // Authenticate with SharePoint
                console.error("Authenticating with SharePoint...");
                const authData = await spauth.getAuth(url, {
                    clientId: clientId,
                    clientSecret: secret,
                    realm: tenantID
                });

                // Define headers from auth data
                const headers = { ...authData.headers };
                headers['Accept'] = 'application/json;odata=verbose';
                console.error("Headers prepared:", headers);

                // Encode the list title to handle special characters
                const encodedListTitle = encodeURIComponent(listTitle);
                
                // First, get the list to validate it exists
                console.error(`Getting list details for "${listTitle}"...`);
                const listResponse = await request({
                    url: `${url}/_api/web/lists/getByTitle('${encodedListTitle}')`,
                    headers: headers,
                    json: true,
                    method: 'GET',
                    timeout: 10000
                });
                
                console.error(`List found: ${listResponse.d.Title}, ID: ${listResponse.d.Id}`);
                
                // Now get all items from the list
                console.error(`Retrieving items from list "${listTitle}"...`);
                const itemsResponse = await request({
                    url: `${url}/_api/web/lists/getByTitle('${encodedListTitle}')/items?$top=5000`,
                    headers: headers,
                    json: true,
                    method: 'GET',
                    timeout: 20000
                });
                
                const items: ISharePointListItem[] = itemsResponse.d.results;
                console.error(`Retrieved ${items.length} items from list "${listTitle}"`);
                
                // Extract field names by looking at the first item (if exists)
                let fieldNames: string[] = [];
                if (items.length > 0) {
                    fieldNames = Object.keys(items[0])
                        .filter(key => !key.startsWith('__') && 
                                       !['AttachmentFiles', 'Attachments', 'FirstUniqueAncestorSecurableObject',
                                         'RoleAssignments', 'ContentType', 'FieldValuesAsHtml', 'FieldValuesAsText', 
                                         'FieldValuesForEdit', 'File', 'Folder', 'ParentList'].includes(key));
                }
                
                console.error(`Fields available: ${fieldNames.join(', ')}`);
                
                // Format items for nicer display - only include relevant fields
                const formattedItems: IFormattedListItem[] = items.map((item: ISharePointListItem) => {
                    const formattedItem: IFormattedListItem = {};
                    fieldNames.forEach(field => {
                        formattedItem[field] = item[field];
                    });
                    return formattedItem;
                });

                return {
                    content: [{
                        type: "text",
                        text: JSON.stringify({
                            listTitle: listTitle,
                            totalItems: items.length,
                            fields: fieldNames,
                            items: formattedItems
                        }, null, 2)
                    }]
                };
            } catch (error: unknown) {
                // Type-safe error handling
                let errorMessage: string;

                if (error instanceof Error) {
                    errorMessage = error.message;
                } else if (typeof error === 'string') {
                    errorMessage = error;
                } else {
                    errorMessage = "Unknown error occurred";
                }

                console.error("Error in getListItems tool:", errorMessage);

                return {
                    content: [{
                        type: "text",
                        text: `Error fetching list items: ${errorMessage}`
                    }],
                    isError: true
                };
            }
        }
    );

    // Add the addMockData tool to generate and add mock data to a SharePoint list
    server.tool(
        "addMockData",
        "Add mock data items to a specific SharePoint list",
        {
            url: z.string().url().describe("URL of the SharePoint website"),
            listTitle: z.string().describe("Title of the SharePoint list to add mock data to"),
            itemCount: z.number().int().min(1).max(100).describe("Number of mock items to create (1-100)")
        },
        async ({ url, listTitle, itemCount }) => {
            console.error(`addMockData tool called with URL: ${url}, List Title: ${listTitle}, Item Count: ${itemCount}`);

            try {
                // Authenticate with SharePoint
                console.error("Authenticating with SharePoint...");
                const authData = await spauth.getAuth(url, {
                    clientId: clientId,
                    clientSecret: secret,
                    realm: tenantID
                });

                // Define headers from auth data
                const headers = { ...authData.headers };
                headers['Accept'] = 'application/json;odata=verbose';
                headers['Content-Type'] = 'application/json;odata=verbose';
                headers['X-RequestDigest'] = await getRequestDigest(url, headers);
                console.error("Headers prepared with request digest");

                // Encode the list title to handle special characters
                const encodedListTitle = encodeURIComponent(listTitle);
                
                // First, get the list schema to understand its fields
                console.error(`Getting list schema for "${listTitle}"...`);
                const listResponse = await request({
                    url: `${url}/_api/web/lists/getByTitle('${encodedListTitle}')`,
                    headers: { ...headers, 'Content-Type': undefined },
                    json: true,
                    method: 'GET',
                    timeout: 30000
                });
                
                // Get field details to understand which fields are writeable
                console.error(`Getting fields for list "${listTitle}"...`);
                const fieldsResponse = await request({
                    url: `${url}/_api/web/lists/getByTitle('${encodedListTitle}')/fields?$filter=ReadOnlyField eq false and Hidden eq false`,
                    headers: { ...headers, 'Content-Type': undefined },
                    json: true,
                    method: 'GET',
                    timeout: 45000
                });
                
                // Process fields to get writeable ones
                const writeableFields = fieldsResponse.d.results.filter((field: any) => {
                    // Skip system fields and fields we shouldn't modify
                    return !field.ReadOnlyField && 
                           !field.Hidden && 
                           field.InternalName !== 'ID' &&
                           field.InternalName !== 'Modified' &&
                           field.InternalName !== 'Created' &&
                           field.InternalName !== 'Author' &&
                           field.InternalName !== 'Editor' &&
                           field.InternalName !== 'GUID' &&
                           !field.InternalName.startsWith('_') 
                });
                
                console.error(`Found ${writeableFields.length} writeable fields`);
                
                // Get lookup data for lookup fields
                const lookupFields = writeableFields.filter((field: any) => 
                    field.TypeAsString?.toLowerCase().includes('lookup'));
                
                // Collect lookup data for each lookup field
                const lookupData: Record<string, any[]> = {};
                
                for (const lookupField of lookupFields) {
                    try {
                        if (lookupField.LookupList) {
                            console.error(`Getting lookup data for field ${lookupField.InternalName} from list ${lookupField.LookupList}...`);
                            
                            // Get the list schema first to find its web URL
                            const lookupListSchema = await request({
                                url: `${url}/_api/web/lists(guid'${lookupField.LookupList}')`,
                                headers: { ...headers, 'Content-Type': undefined },
                                json: true,
                                method: 'GET',
                                timeout: 30000
                            });
                            
                            // Get items from the lookup list
                            const lookupItems = await request({
                                url: `${url}/_api/web/lists(guid'${lookupField.LookupList}')/items?$select=ID,${lookupField.LookupField}&$top=100`,
                                headers: { ...headers, 'Content-Type': undefined },
                                json: true,
                                method: 'GET',
                                timeout: 45000
                            });
                            
                            if (lookupItems.d && lookupItems.d.results && lookupItems.d.results.length > 0) {
                                lookupData[lookupField.InternalName] = lookupItems.d.results.map((item: any) => ({
                                    ID: item.ID,
                                    Value: item[lookupField.LookupField]
                                }));
                                console.error(`Found ${lookupData[lookupField.InternalName].length} lookup values for ${lookupField.InternalName}`);
                            } else {
                                console.error(`No lookup data found for field ${lookupField.InternalName}`);
                                lookupData[lookupField.InternalName] = [];
                            }
                        }
                    } catch (error) {
                        console.error(`Error fetching lookup data for ${lookupField.InternalName}:`, error);
                        lookupData[lookupField.InternalName] = [];
                    }
                }
                
                // Add mock items
                const successfulItems: number[] = [];
                const failedItems: Array<{index: number, error: string}> = [];
                
                for (let i = 0; i < itemCount; i++) {
                    try {
                        // Create mock item data based on field types
                        const mockItemData: Record<string, any> = {};
                        
                        for (const field of writeableFields) {
                            const fieldName = field.InternalName;
                            const fieldType = field.TypeAsString?.toLowerCase() || '';
                            let mockValue = generateMockValueForField(field, i);
                            
                            // Handle lookup fields with real lookup data
                            if (mockValue && typeof mockValue === 'object' && mockValue.__lookupField) {
                                const fieldLookupData = lookupData[fieldName] || [];
                                
                                if (fieldLookupData.length > 0) {
                                    // Use modulo to cycle through available lookup values
                                    const lookupIndex = i % fieldLookupData.length;
                                    const lookupItem = fieldLookupData[lookupIndex];
                                    
                                    if (mockValue.multiple) {
                                        // Multi-value lookup requires array of lookup values
                                        mockItemData[fieldName] = {
                                            __metadata: { type: 'Collection(Edm.Int32)' },
                                            results: [lookupItem.ID]
                                        };
                                    } else {
                                        // Single-value lookup
                                        mockItemData[`${fieldName}Id`] = lookupItem.ID;
                                    }
                                    
                                    console.error(`Set lookup value for ${fieldName}: ${lookupItem.ID} (${lookupItem.Value})`);
                                } else {
                                    console.error(`No lookup data available for ${fieldName}, skipping field`);
                                }
                            } else if (mockValue !== null && mockValue !== undefined) {
                                mockItemData[fieldName] = mockValue;
                            }
                        }
                        
                        // Always include a Title if it exists in the writeable fields
                        if (writeableFields.some((f: any) => f.InternalName === 'Title') && !mockItemData['Title']) {
                            mockItemData['Title'] = `Mock Item ${i + 1}`;
                        }
                        
                        console.error(`Creating mock item ${i + 1}/${itemCount}...`);
                        console.error(`Data: ${JSON.stringify(mockItemData)}`);
                        
                        // Create the item in SharePoint
                        const createResponse = await request({
                            url: `${url}/_api/web/lists/getByTitle('${encodedListTitle}')/items`,
                            headers: headers,
                            json: true,
                            method: 'POST',
                            body: { __metadata: { type: listResponse.d.ListItemEntityTypeFullName }, ...mockItemData },
                            timeout: 15000
                        });
                        
                        successfulItems.push(i + 1);
                        console.error(`Successfully created item ${i + 1}`);
                    } catch (error: any) {
                        console.error(`Error creating item ${i + 1}:`, error.message);
                        failedItems.push({ index: i + 1, error: error.message });
                    }
                }
                
                return {
                    content: [{
                        type: "text",
                        text: JSON.stringify({
                            listTitle: listTitle,
                            writeableFields: writeableFields.map((f: any) => ({
                                name: f.InternalName,
                                title: f.Title,
                                type: f.TypeAsString || f.TypeDisplayName
                            })),
                            lookupFields: Object.keys(lookupData).map(key => ({
                                name: key, 
                                valuesFound: lookupData[key].length
                            })),
                            requested: itemCount,
                            successful: successfulItems.length,
                            failed: failedItems.length,
                            successfulItems,
                            failedItems
                        }, null, 2)
                    }]
                };
            } catch (error: unknown) {
                // Type-safe error handling
                let errorMessage: string;

                if (error instanceof Error) {
                    errorMessage = error.message;
                } else if (typeof error === 'string') {
                    errorMessage = error;
                } else {
                    errorMessage = "Unknown error occurred";
                }

                console.error("Error in addMockData tool:", errorMessage);

                return {
                    content: [{
                        type: "text",
                        text: `Error adding mock data: ${errorMessage}`
                    }],
                    isError: true
                };
            }
        }
    );
}

/**
 * Get a request digest for SharePoint POST operations
 */
async function getRequestDigest(url: string, headers: Record<string, string>): Promise<string> {
    try {
        const digestResponse = await request({
            url: `${url}/_api/contextinfo`,
            method: 'POST',
            headers: { ...headers, 'Content-Type': undefined },
            json: true
        });
        
        return digestResponse.d.GetContextWebInformation.FormDigestValue;
    } catch (error) {
        console.error('Error getting request digest:', error);
        throw new Error('Failed to get request digest required for creating items');
    }
}

/**
 * Generate mock data for a SharePoint field based on its type
 */
function generateMockValueForField(field: any, index: number): any {
    const fieldType = field.TypeDisplayName?.toLowerCase() || field.TypeAsString?.toLowerCase() || '';
    const fieldName = field.InternalName;
    
    // Common prefix for generated values to make them identifiable as mock data
    const mockPrefix = `Mock-${index + 1}`;
    
    switch (fieldType) {
        case 'single line of text':
        case 'text':
            if (fieldName.toLowerCase().includes('name')) {
                return `${mockPrefix}: Name`;
            } else if (fieldName.toLowerCase().includes('title')) {
                return `${mockPrefix}: Title`;
            } else if (fieldName.toLowerCase().includes('description')) {
                return `${mockPrefix}: Description text for this mock item`;
            } else {
                return `${mockPrefix}: ${field.Title || fieldName}`;
            }
            
        case 'multiple lines of text':
        case 'note':
            return `${mockPrefix}: This is a longer text for the field "${field.Title || fieldName}".\nThis is some additional text to make it multi-line.\nGenerated as mock data.`;
            
        case 'number':
            return index * 10 + Math.floor(Math.random() * 100);
            
        case 'currency':
            return (index * 10.25 + Math.random() * 100).toFixed(2);
            
        case 'date and time':
        case 'datetime':
            const mockDate = new Date();
            mockDate.setDate(mockDate.getDate() + index); // Each item gets a different date
            return mockDate.toISOString();
            
        case 'choice':
        case 'multichoice':
            // If choices are available, use one of them
            if (field.Choices && field.Choices.results && field.Choices.results.length > 0) {
                const choiceIndex = index % field.Choices.results.length;
                return field.Choices.results[choiceIndex];
            }
            return `Choice ${(index % 5) + 1}`;
            
        case 'yes/no':
        case 'boolean':
            return index % 2 === 0;
            
        case 'person or group':
        case 'user':
            // We can't truly populate this without knowing valid users
            return null;
            
        case 'hyperlink':
        case 'url':
            return `https://example.com/mock-link-${index}`;
            
        case 'lookup':
        case 'lookupfield':
        case 'lookup (allow multiple values)':
        case 'lookupfieldmulti':
            // For lookup fields, we need to use the correct format expected by SharePoint
            // Since we don't know valid IDs, we'll return a placeholder that the caller should replace
            console.error(`Field ${fieldName} is a lookup field. Need special handling.`);
            return { __lookupField: true, fieldName, multiple: fieldType.includes('multiple') };
            
        default:
            // For unsupported types, return null to skip
            console.error(`Unsupported field type: ${fieldType} for field ${fieldName}`);
            return null;
    }
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