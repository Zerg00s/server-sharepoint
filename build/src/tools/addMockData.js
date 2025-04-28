"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.addMockData = addMockData;
// src/tools/addMockData.ts
const request_promise_1 = __importDefault(require("request-promise"));
const auth_1 = require("../auth");
const mockDataGenerator_1 = require("../utils/mockDataGenerator");
/**
 * Add mock data items to a specific SharePoint list
 * @param params Parameters including site URL, list title, and item count
 * @param config SharePoint configuration
 * @returns Tool result with creation summary
 */
async function addMockData(params, config) {
    const { url, listTitle, itemCount } = params;
    console.error(`addMockData tool called with URL: ${url}, List Title: ${listTitle}, Item Count: ${itemCount}`);
    try {
        // Authenticate with SharePoint
        const headers = await (0, auth_1.getSharePointHeaders)(url, config);
        console.error("Headers prepared with authentication");
        // Get request digest for POST operations
        headers['X-RequestDigest'] = await (0, auth_1.getRequestDigest)(url, headers);
        headers['Content-Type'] = 'application/json;odata=verbose';
        console.error("Headers prepared with request digest");
        // Encode the list title to handle special characters
        const encodedListTitle = encodeURIComponent(listTitle);
        // First, get the list schema to understand its fields
        console.error(`Getting list schema for "${listTitle}"...`);
        const listResponse = await (0, request_promise_1.default)({
            url: `${url}/_api/web/lists/getByTitle('${encodedListTitle}')`,
            headers: { ...headers, 'Content-Type': undefined },
            json: true,
            method: 'GET',
            timeout: 30000
        });
        // Get field details to understand which fields are writeable
        console.error(`Getting fields for list "${listTitle}"...`);
        const fieldsResponse = await (0, request_promise_1.default)({
            url: `${url}/_api/web/lists/getByTitle('${encodedListTitle}')/fields?$filter=ReadOnlyField eq false and Hidden eq false`,
            headers: { ...headers, 'Content-Type': undefined },
            json: true,
            method: 'GET',
            timeout: 45000
        });
        // Process fields to get writeable ones
        const writeableFields = fieldsResponse.d.results.filter((field) => {
            // Skip system fields and fields we shouldn't modify
            return !field.ReadOnlyField &&
                !field.Hidden &&
                field.InternalName !== 'ID' &&
                field.InternalName !== 'Modified' &&
                field.InternalName !== 'Created' &&
                field.InternalName !== 'Author' &&
                field.InternalName !== 'Editor' &&
                field.InternalName !== 'GUID' &&
                !field.InternalName.startsWith('_');
        });
        console.error(`Found ${writeableFields.length} writeable fields`);
        // Get lookup data for lookup fields
        const lookupFields = writeableFields.filter((field) => field.TypeAsString?.toLowerCase().includes('lookup'));
        // Collect lookup data for each lookup field
        const lookupData = {};
        for (const lookupField of lookupFields) {
            try {
                if (lookupField.LookupList) {
                    console.error(`Getting lookup data for field ${lookupField.InternalName} from list ${lookupField.LookupList}...`);
                    // Get the list schema first to find its web URL
                    const lookupListSchema = await (0, request_promise_1.default)({
                        url: `${url}/_api/web/lists(guid'${lookupField.LookupList}')`,
                        headers: { ...headers, 'Content-Type': undefined },
                        json: true,
                        method: 'GET',
                        timeout: 30000
                    });
                    // Get items from the lookup list
                    const lookupItems = await (0, request_promise_1.default)({
                        url: `${url}/_api/web/lists(guid'${lookupField.LookupList}')/items?$select=ID,${lookupField.LookupField}&$top=100`,
                        headers: { ...headers, 'Content-Type': undefined },
                        json: true,
                        method: 'GET',
                        timeout: 45000
                    });
                    if (lookupItems.d && lookupItems.d.results && lookupItems.d.results.length > 0) {
                        lookupData[lookupField.InternalName] = lookupItems.d.results.map((item) => ({
                            ID: item.ID,
                            Value: item[lookupField.LookupField]
                        }));
                        console.error(`Found ${lookupData[lookupField.InternalName].length} lookup values for ${lookupField.InternalName}`);
                    }
                    else {
                        console.error(`No lookup data found for field ${lookupField.InternalName}`);
                        lookupData[lookupField.InternalName] = [];
                    }
                }
            }
            catch (error) {
                console.error(`Error fetching lookup data for ${lookupField.InternalName}:`, error);
                lookupData[lookupField.InternalName] = [];
            }
        }
        // Add mock items
        const successfulItems = [];
        const failedItems = [];
        for (let i = 0; i < itemCount; i++) {
            try {
                // Create mock item data based on field types
                const mockItemData = {};
                for (const field of writeableFields) {
                    const fieldName = field.InternalName;
                    const fieldType = field.TypeAsString?.toLowerCase() || '';
                    let mockValue = (0, mockDataGenerator_1.generateMockValueForField)(field, i);
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
                            }
                            else {
                                // Single-value lookup
                                mockItemData[`${fieldName}Id`] = lookupItem.ID;
                            }
                            console.error(`Set lookup value for ${fieldName}: ${lookupItem.ID} (${lookupItem.Value})`);
                        }
                        else {
                            console.error(`No lookup data available for ${fieldName}, skipping field`);
                        }
                    }
                    else if (mockValue !== null && mockValue !== undefined) {
                        mockItemData[fieldName] = mockValue;
                    }
                }
                // Always include a Title if it exists in the writeable fields
                if (writeableFields.some((f) => f.InternalName === 'Title') && !mockItemData['Title']) {
                    mockItemData['Title'] = `Mock Item ${i + 1}`;
                }
                console.error(`Creating mock item ${i + 1}/${itemCount}...`);
                console.error(`Data: ${JSON.stringify(mockItemData)}`);
                // Create the item in SharePoint
                const createResponse = await (0, request_promise_1.default)({
                    url: `${url}/_api/web/lists/getByTitle('${encodedListTitle}')/items`,
                    headers: headers,
                    json: true,
                    method: 'POST',
                    body: { __metadata: { type: listResponse.d.ListItemEntityTypeFullName }, ...mockItemData },
                    timeout: 15000
                });
                successfulItems.push(i + 1);
                console.error(`Successfully created item ${i + 1}`);
            }
            catch (error) {
                console.error(`Error creating item ${i + 1}:`, error.message);
                failedItems.push({ index: i + 1, error: error.message });
            }
        }
        // Prepare result object
        const result = {
            listTitle: listTitle,
            writeableFields: writeableFields.map((f) => ({
                name: f.InternalName,
                title: f.Title,
                type: f.TypeAsString || f.TypeDisplayName || 'Unknown'
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
        };
        return {
            content: [{
                    type: "text",
                    text: JSON.stringify(result, null, 2)
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
exports.default = addMockData;
