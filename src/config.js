"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.parseCliArgs = parseCliArgs;
exports.loadConfig = loadConfig;
exports.validateConfig = validateConfig;
exports.getSharePointConfig = getSharePointConfig;
// src/config.ts
var dotenv = require("dotenv");
// Load environment variables from .env file
dotenv.config();
// Parse command line arguments
function parseCliArgs() {
    return process.argv.slice(2).reduce(function (acc, arg) {
        var _a = arg.split('='), key = _a[0], value = _a[1];
        if (key && value) {
            acc[key.replace(/^--/, '')] = value;
        }
        return acc;
    }, {});
}
// Load and validate configuration
function loadConfig() {
    var args = parseCliArgs();
    // Priority: CLI args → fallback: ENV
    var config = {
        clientId: args.clientId || process.env.SHAREPOINT_CLIENT_ID || '',
        clientSecret: args.clientSecret || process.env.SHAREPOINT_CLIENT_SECRET || '',
        tenantId: args.tenantId || process.env.SHAREPOINT_TENANT_ID || '',
        siteUrl: args.siteUrl || process.env.SHAREPOINT_SITE_URL || ''
    };
    return config;
}
// Validate that required config values are present
function validateConfig(config) {
    var clientId = config.clientId, clientSecret = config.clientSecret, tenantId = config.tenantId;
    var isValid = Boolean(clientId && clientSecret && tenantId);
    if (!isValid) {
        console.error("ERROR: Missing SharePoint credentials!");
        console.error("Provide via environment variables or CLI arguments like:");
        console.error("--clientId=xxx --clientSecret=yyy --tenantId=zzz");
    }
    else {
        console.error("✅ SharePoint credentials loaded.");
    }
    return isValid;
}
// Get the SharePoint configuration, combining loading and validation
function getSharePointConfig() {
    var config = loadConfig();
    validateConfig(config);
    return config;
}
// Export as default
exports.default = getSharePointConfig;
