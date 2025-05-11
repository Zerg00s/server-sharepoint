"use strict";
var __assign = (this && this.__assign) || function () {
    __assign = Object.assign || function(t) {
        for (var s, i = 1, n = arguments.length; i < n; i++) {
            s = arguments[i];
            for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p))
                t[p] = s[p];
        }
        return t;
    };
    return __assign.apply(this, arguments);
};
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g = Object.create((typeof Iterator === "function" ? Iterator : Object).prototype);
    return g.next = verb(0), g["throw"] = verb(1), g["return"] = verb(2), typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (g && (g = 0, op[0] && (_ = 0)), _) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.getSharePointHeaders = getSharePointHeaders;
exports.getRequestDigest = getRequestDigest;
// src/auth.ts
var spauth = require("node-sp-auth");
var request_promise_1 = require("request-promise");
/**
 * Authenticate with SharePoint and get headers
 * @param url The SharePoint site URL
 * @param config The SharePoint configuration
 * @returns Authentication headers for API requests
 */
function getSharePointHeaders(url, config) {
    return __awaiter(this, void 0, void 0, function () {
        var authData, headers, error_1, errorDetail;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    _a.trys.push([0, 2, , 3]);
                    console.error("Authenticating with SharePoint...");
                    // Validate required configuration
                    if (!config.clientId) {
                        throw new Error("SharePoint Client ID is missing from configuration");
                    }
                    if (!config.clientSecret) {
                        throw new Error("SharePoint Client Secret is missing from configuration");
                    }
                    if (!config.tenantId) {
                        throw new Error("SharePoint Tenant ID is missing from configuration");
                    }
                    console.error("Attempting to authenticate to URL: ".concat(url));
                    console.error("Using client ID: ".concat(config.clientId.substring(0, 5), "..."));
                    console.error("Using tenant ID: ".concat(config.tenantId.substring(0, 5), "..."));
                    return [4 /*yield*/, spauth.getAuth(url, {
                            clientId: config.clientId,
                            clientSecret: config.clientSecret,
                            realm: config.tenantId
                        })];
                case 1:
                    authData = _a.sent();
                    headers = __assign({}, authData.headers);
                    headers['Accept'] = 'application/json;odata=verbose';
                    console.error("SharePoint authentication successful");
                    console.error("Headers obtained:", Object.keys(headers).join(", "));
                    return [2 /*return*/, headers];
                case 2:
                    error_1 = _a.sent();
                    console.error("SharePoint authentication failed");
                    errorDetail = "";
                    if (error_1 instanceof Error) {
                        errorDetail = error_1.message;
                        console.error("Error stack:", error_1.stack);
                        // Check for specific known error patterns
                        if (error_1.message.includes("invalid_client")) {
                            errorDetail = "Invalid client credentials (client ID or client secret)";
                        }
                        else if (error_1.message.includes("invalid_grant")) {
                            errorDetail = "Invalid grant (tenant ID may be incorrect)";
                        }
                        else if (error_1.message.toLowerCase().includes("forbidden") ||
                            error_1.message.includes("403")) {
                            errorDetail = "Access forbidden - check app permissions";
                        }
                        else if (error_1.message.toLowerCase().includes("not found") ||
                            error_1.message.includes("404")) {
                            errorDetail = "Resource not found - check site URL";
                        }
                        else if (error_1.message.toLowerCase().includes("timeout")) {
                            errorDetail = "Request timed out - check network connection";
                        }
                    }
                    else {
                        errorDetail = String(error_1);
                    }
                    // If this appears to be a URL issue, suggest a fix
                    if (!url.endsWith("/")) {
                        console.error("Note: URL does not end with a trailing slash, which can cause issues with some SharePoint sites");
                    }
                    throw new Error("SharePoint authentication failed: ".concat(errorDetail, ". Verify your configuration and app permissions."));
                case 3: return [2 /*return*/];
            }
        });
    });
}
/**
 * Get a request digest for SharePoint POST operations
 * @param url The SharePoint site URL
 * @param headers The authentication headers
 * @returns Request digest value
 */
function getRequestDigest(url, headers) {
    return __awaiter(this, void 0, void 0, function () {
        var digestUrl, digestResponse, digestValue, error_2;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    _a.trys.push([0, 2, , 3]);
                    console.error("Getting request digest for POST operations...");
                    if (!headers || Object.keys(headers).length === 0) {
                        throw new Error("Headers are empty or undefined");
                    }
                    digestUrl = "".concat(url, "/_api/contextinfo");
                    console.error("Requesting digest from: ".concat(digestUrl));
                    return [4 /*yield*/, (0, request_promise_1.default)({
                            url: digestUrl,
                            method: 'POST',
                            headers: __assign(__assign({}, headers), { 'Content-Type': undefined }),
                            json: true,
                            timeout: 30000,
                            // Adding full response option to get more diagnostic info if needed
                            resolveWithFullResponse: true,
                            simple: false // Don't throw on non-2xx responses
                        })];
                case 1:
                    digestResponse = _a.sent();
                    // Check the response status
                    if (digestResponse.statusCode >= 400) {
                        throw new Error("HTTP Error ".concat(digestResponse.statusCode, ": ").concat(JSON.stringify(digestResponse.body)));
                    }
                    if (!digestResponse.body || !digestResponse.body.d || !digestResponse.body.d.GetContextWebInformation) {
                        console.error("Unexpected digest response format:", JSON.stringify(digestResponse.body));
                        throw new Error("Invalid digest response format");
                    }
                    digestValue = digestResponse.body.d.GetContextWebInformation.FormDigestValue;
                    console.error("Request digest obtained successfully");
                    return [2 /*return*/, digestValue];
                case 2:
                    error_2 = _a.sent();
                    console.error('Error getting request digest:');
                    // Enhanced error handling
                    if (error_2 instanceof Error) {
                        console.error("Error message: ".concat(error_2.message));
                        console.error("Error stack: ".concat(error_2.stack));
                        // Try to identify common issues
                        if (error_2.message.includes("403")) {
                            console.error("This appears to be a permissions issue. Check your app permissions.");
                        }
                        else if (error_2.message.includes("404")) {
                            console.error("The contextinfo endpoint could not be found. Check your site URL.");
                        }
                        else if (error_2.message.includes("timeout")) {
                            console.error("The request timed out. Check your network connection.");
                        }
                    }
                    else {
                        console.error(error_2);
                    }
                    throw new Error('Failed to get request digest required for operations. Check app permissions and URL.');
                case 3: return [2 /*return*/];
            }
        });
    });
}
exports.default = {
    getSharePointHeaders: getSharePointHeaders,
    getRequestDigest: getRequestDigest
};
