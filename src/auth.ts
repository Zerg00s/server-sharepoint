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
    } catch (error) {
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
export async function getRequestDigest(url: string, headers: AuthHeaders): Promise<string> {
    try {
        console.error("Getting request digest for POST operations...");
        
        const digestResponse = await request({
            url: `${url}/_api/contextinfo`,
            method: 'POST',
            headers: { ...headers, 'Content-Type': undefined },
            json: true
        });
        
        console.error("Request digest obtained successfully");
        return digestResponse.d.GetContextWebInformation.FormDigestValue;
    } catch (error) {
        console.error('Error getting request digest:', error);
        throw new Error('Failed to get request digest required for creating items');
    }
}

export default {
    getSharePointHeaders,
    getRequestDigest
};