// src/auth.ts
import * as spauth from 'node-sp-auth';
import request from 'request-promise';
import { SharePointConfig } from './config';

// Interface for authentication headers
export interface AuthHeaders {
    [key: string]: string;
}

/**
 * Authenticate with SharePoint and get headers
 * @param url The SharePoint site URL
 * @param config The SharePoint configuration
 * @returns Authentication headers for API requests
 */
export async function getSharePointHeaders(url: string, config: SharePointConfig): Promise<AuthHeaders> {
    try {
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
        
        console.error(`Attempting to authenticate to URL: ${url}`);
        console.error(`Using client ID: ${config.clientId.substring(0, 5)}...`);
        console.error(`Using tenant ID: ${config.tenantId.substring(0, 5)}...`);
        
        // Try the authentication
        const authData = await spauth.getAuth(url, {
            clientId: config.clientId,
            clientSecret: config.clientSecret,
            realm: config.tenantId
        });

        // Define headers from auth data
        const headers = { ...authData.headers };
        headers['Accept'] = 'application/json;odata=verbose';
        
        console.error("SharePoint authentication successful");
        console.error("Headers obtained:", Object.keys(headers).join(", "));
        
        return headers;
    } catch (error) {
        console.error("SharePoint authentication failed");
        
        // Enhanced error handling with more context
        let errorDetail = "";
        
        if (error instanceof Error) {
            errorDetail = error.message;
            console.error("Error stack:", error.stack);
            
            // Check for specific known error patterns
            if (error.message.includes("invalid_client")) {
                errorDetail = "Invalid client credentials (client ID or client secret)";
            } else if (error.message.includes("invalid_grant")) {
                errorDetail = "Invalid grant (tenant ID may be incorrect)";
            } else if (error.message.toLowerCase().includes("forbidden") || 
                      error.message.includes("403")) {
                errorDetail = "Access forbidden - check app permissions";
            } else if (error.message.toLowerCase().includes("not found") || 
                      error.message.includes("404")) {
                errorDetail = "Resource not found - check site URL";
            } else if (error.message.toLowerCase().includes("timeout")) {
                errorDetail = "Request timed out - check network connection";
            }
        } else {
            errorDetail = String(error);
        }
        
        // If this appears to be a URL issue, suggest a fix
        if (!url.endsWith("/")) {
            console.error("Note: URL does not end with a trailing slash, which can cause issues with some SharePoint sites");
        }
        
        throw new Error(`SharePoint authentication failed: ${errorDetail}. Verify your configuration and app permissions.`);
    }
}

/**
 * Get a request digest for SharePoint POST operations
 * @param url The SharePoint site URL
 * @param headers The authentication headers
 * @returns Request digest value
 */
export async function getRequestDigest(url: string, headers: AuthHeaders): Promise<string> {
    try {
        console.error("Getting request digest for POST operations...");
        
        if (!headers || Object.keys(headers).length === 0) {
            throw new Error("Headers are empty or undefined");
        }
        
        const digestUrl = `${url}/_api/contextinfo`;
        console.error(`Requesting digest from: ${digestUrl}`);
        
        const digestResponse = await request({
            url: digestUrl,
            method: 'POST',
            headers: { ...headers, 'Content-Type': undefined },
            json: true,
            timeout: 30000,
            // Adding full response option to get more diagnostic info if needed
            resolveWithFullResponse: true,
            simple: false // Don't throw on non-2xx responses
        });
        
        // Check the response status
        if (digestResponse.statusCode >= 400) {
            throw new Error(`HTTP Error ${digestResponse.statusCode}: ${JSON.stringify(digestResponse.body)}`);
        }
        
        if (!digestResponse.body || !digestResponse.body.d || !digestResponse.body.d.GetContextWebInformation) {
            console.error("Unexpected digest response format:", JSON.stringify(digestResponse.body));
            throw new Error("Invalid digest response format");
        }
        
        const digestValue = digestResponse.body.d.GetContextWebInformation.FormDigestValue;
        console.error("Request digest obtained successfully");
        
        return digestValue;
    } catch (error) {
        console.error('Error getting request digest:');
        
        // Enhanced error handling
        if (error instanceof Error) {
            console.error(`Error message: ${error.message}`);
            console.error(`Error stack: ${error.stack}`);
            
            // Try to identify common issues
            if (error.message.includes("403")) {
                console.error("This appears to be a permissions issue. Check your app permissions.");
            } else if (error.message.includes("404")) {
                console.error("The contextinfo endpoint could not be found. Check your site URL.");
            } else if (error.message.includes("timeout")) {
                console.error("The request timed out. Check your network connection.");
            }
        } else {
            console.error(error);
        }
        
        throw new Error('Failed to get request digest required for operations. Check app permissions and URL.');
    }
}

export default {
    getSharePointHeaders,
    getRequestDigest
};