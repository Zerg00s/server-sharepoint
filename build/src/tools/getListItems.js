"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.getListItems = getListItems;
// src/tools/getListItems.ts
const request_promise_1 = __importDefault(require("request-promise"));
const auth_1 = require("../auth");
/**
 * Get all items from a specific SharePoint list
 * @param params Parameters including site URL and list title
 * @param config SharePoint configuration
 * @returns Tool result with list items data
 */
async function getListItems(params, config) {
    const { url, listTitle } = params;
    console.error(`getListItems tool called with URL: ${url}, List Title: ${listTitle}`);
    try {
        // Authenticate with SharePoint
        const headers = await (0, auth_1.getSharePointHeaders)(url, config);
        console.error("Headers prepared:", headers);
        // Encode the list title to handle special characters
        const encodedListTitle = encodeURIComponent(listTitle);
        // First, get the list to validate it exists
        console.error(`Getting list details for "${listTitle}"...`);
        const listResponse = await (0, request_promise_1.default)({
            url: `${url}/_api/web/lists/getByTitle('${encodedListTitle}')`,
            headers: headers,
            json: true,
            method: 'GET',
            timeout: 10000
        });
        console.error(`List found: ${listResponse.d.Title}, ID: ${listResponse.d.Id}`);
        // Now get all items from the list
        console.error(`Retrieving items from list "${listTitle}"...`);
        const itemsResponse = await (0, request_promise_1.default)({
            url: `${url}/_api/web/lists/getByTitle('${encodedListTitle}')/items?$top=5000`,
            headers: headers,
            json: true,
            method: 'GET',
            timeout: 20000
        });
        const items = itemsResponse.d.results;
        console.error(`Retrieved ${items.length} items from list "${listTitle}"`);
        // Extract field names by looking at the first item (if exists)
        let fieldNames = [];
        if (items.length > 0) {
            fieldNames = Object.keys(items[0])
                .filter(key => !key.startsWith('__') &&
                !['AttachmentFiles', 'Attachments', 'FirstUniqueAncestorSecurableObject',
                    'RoleAssignments', 'ContentType', 'FieldValuesAsHtml', 'FieldValuesAsText',
                    'FieldValuesForEdit', 'File', 'Folder', 'ParentList'].includes(key));
        }
        console.error(`Fields available: ${fieldNames.join(', ')}`);
        // Format items for nicer display - only include relevant fields
        const formattedItems = items.map((item) => {
            const formattedItem = {};
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
    }
    catch (error) {
        // Type-safe error handling
        let errorMessage;
        if (error instanceof Error) {
            errorMessage = error.message;
        }
        else if (typeof error === 'string') {
            errorMessage = error;
        }
        else {
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
exports.default = getListItems;
