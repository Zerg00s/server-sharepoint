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
Object.defineProperty(exports, "__esModule", { value: true });
exports.parseCliArgs = parseCliArgs;
exports.loadConfig = loadConfig;
exports.validateConfig = validateConfig;
exports.getSharePointConfig = getSharePointConfig;
// src/config.ts
const dotenv = __importStar(require("dotenv"));
// Load environment variables from .env file
dotenv.config();
// Parse command line arguments
function parseCliArgs() {
    return process.argv.slice(2).reduce((acc, arg) => {
        const [key, value] = arg.split('=');
        if (key && value) {
            acc[key.replace(/^--/, '')] = value;
        }
        return acc;
    }, {});
}
// Load and validate configuration
function loadConfig() {
    const args = parseCliArgs();
    // Priority: CLI args → fallback: ENV
    const config = {
        clientId: args.clientId || process.env.SHAREPOINT_CLIENT_ID || '',
        clientSecret: args.clientSecret || process.env.SHAREPOINT_CLIENT_SECRET || '',
        tenantId: args.tenantId || process.env.SHAREPOINT_TENANT_ID || '',
        siteUrl: args.siteUrl || process.env.SHAREPOINT_SITE_URL || ''
    };
    return config;
}
// Validate that required config values are present
function validateConfig(config) {
    const { clientId, clientSecret, tenantId } = config;
    const isValid = Boolean(clientId && clientSecret && tenantId);
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
    const config = loadConfig();
    validateConfig(config);
    return config;
}
// Export as default
exports.default = getSharePointConfig;
