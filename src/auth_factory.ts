// src/auth_factory.ts - Fixed version
import { SharePointConfig, SharePointSecretConfig, SharePointCertConfig } from './config';
import { getSharePointHeaders as getSharePointHeadersSecret, getRequestDigest as getRequestDigestSecret, AuthHeaders } from './auth';
import { getSharePointHeaders as getSharePointHeadersCert } from './azure_cert_auth';
import axios from 'axios';

/**
 * Get the appropriate SharePoint authentication headers based on the configuration
 * @param url The SharePoint site URL
 * @param config The SharePoint configuration
 * @returns Authentication headers for API requests
 */
export async function getSharePointHeaders(url: string, config: SharePointConfig): Promise<AuthHeaders> {
    // Determine which authentication method to use
    if (config.authType === 'certificate') {
        // Use certificate-based authentication
        const certConfig: SharePointCertConfig = config;
        return getSharePointHeadersCert(url, certConfig);
    } else {
        // Use client secret authentication
        const secretConfig: SharePointSecretConfig = config;
        return getSharePointHeadersSecret(url, secretConfig);
    }
}

/**
 * Get a request digest for SharePoint POST operations
 * This works with both certificate and client secret authentication
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
        
        const digestResponse = await axios({
            url: digestUrl,
            method: 'POST',
            headers: { 
                ...headers, 
                'Accept': 'application/json;odata=verbose' 
            },
            timeout: 30000
        });
        
        // Check the response status
        if (digestResponse.status >= 400) {
            throw new Error(`HTTP Error ${digestResponse.status}: ${JSON.stringify(digestResponse.data)}`);
        }
        
        if (!digestResponse.data || !digestResponse.data.d || !digestResponse.data.d.GetContextWebInformation) {
            console.error("Unexpected digest response format:", JSON.stringify(digestResponse.data));
            throw new Error("Invalid digest response format");
        }
        
        const digestValue = digestResponse.data.d.GetContextWebInformation.FormDigestValue;
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