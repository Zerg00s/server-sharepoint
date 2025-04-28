"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.getTitle = getTitle;
// src/tools/getTitle.ts
const request_promise_1 = __importDefault(require("request-promise"));
const auth_1 = require("../auth");
/**
 * Get the title of a SharePoint website
 * @param params Parameters including the SharePoint site URL
 * @param config SharePoint configuration
 * @returns Tool result with site title
 */
async function getTitle(params, config) {
    const { url } = params;
    console.error("getTitle tool called with URL:", url);
    try {
        // Authenticate with SharePoint
        const headers = await (0, auth_1.getSharePointHeaders)(url, config);
        console.error("Headers prepared:", headers);
        // Make request to SharePoint API
        console.error("Making request to SharePoint API...");
        const response = await (0, request_promise_1.default)({
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
exports.default = getTitle;
