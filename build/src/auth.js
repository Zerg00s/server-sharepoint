"use strict";
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
exports.getSharePointHeaders = getSharePointHeaders;
exports.getRequestDigest = getRequestDigest;
// src/auth.ts
const spauth = __importStar(require("node-sp-auth"));
const request_promise_1 = __importDefault(require("request-promise"));
/**
 * Authenticate with SharePoint and get headers
 * @param url The SharePoint site URL
 * @param config The SharePoint configuration
 * @returns Authentication headers for API requests
 */
async function getSharePointHeaders(url, config) {
    try {
        console.error("Authenticating with SharePoint...");
        const authData = await spauth.getAuth(url, {
            clientId: config.clientId,
            clientSecret: config.clientSecret,
            realm: config.tenantId
        });
        // Define headers from auth data
        const headers = { ...authData.headers };
        headers['Accept'] = 'application/json;odata=verbose';
        console.error("SharePoint authentication successful");
        return headers;
    }
    catch (error) {
        console.error("SharePoint authentication failed:", error);
        throw new Error(`SharePoint authentication failed: ${error instanceof Error ? error.message : String(error)}`);
    }
}
/**
 * Get a request digest for SharePoint POST operations
 * @param url The SharePoint site URL
 * @param headers The authentication headers
 * @returns Request digest value
 */
async function getRequestDigest(url, headers) {
    try {
        console.error("Getting request digest for POST operations...");
        const digestResponse = await (0, request_promise_1.default)({
            url: `${url}/_api/contextinfo`,
            method: 'POST',
            headers: { ...headers, 'Content-Type': undefined },
            json: true
        });
        console.error("Request digest obtained successfully");
        return digestResponse.d.GetContextWebInformation.FormDigestValue;
    }
    catch (error) {
        console.error('Error getting request digest:', error);
        throw new Error('Failed to get request digest required for creating items');
    }
}
exports.default = {
    getSharePointHeaders,
    getRequestDigest
};
