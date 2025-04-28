"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.getLists = getLists;
// src/tools/getLists.ts
const request_promise_1 = __importDefault(require("request-promise"));
const auth_1 = require("../auth");
/**
 * Get all visible SharePoint lists from a site
 * @param params Parameters including the SharePoint site URL
 * @param config SharePoint configuration
 * @returns Tool result with lists data
 */
async function getLists(params, config) {
    const { url } = params;
    console.error("getLists tool called with URL:", url);
    try {
        // Authenticate with SharePoint
        const headers = await (0, auth_1.getSharePointHeaders)(url, config);
        console.error("Headers prepared:", headers);
        // Make request to SharePoint API to get lists
        console.error("Making request to SharePoint API for lists...");
        const response = await (0, request_promise_1.default)({
            url: `${url}/_api/web/lists?$select=Title,Id,ItemCount,LastItemModifiedDate,Description,BaseTemplate,Hidden,IsSystemList,RootFolder/ServerRelativeUrl&$expand=RootFolder`,
            headers: headers,
            json: true,
            method: 'GET',
            timeout: 15000
        });
        console.error(`SharePoint API response received with ${response.d.results.length} total lists`);
        // Filter out hidden and system lists
        const visibleLists = response.d.results.filter((list) => !list.Hidden && !list.IsSystemList);
        console.error(`Filtered to ${visibleLists.length} visible lists (excluding hidden and system lists)`);
        // Format the list data for display
        const lists = visibleLists.map((list) => {
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
exports.default = getLists;
