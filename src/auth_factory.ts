// src/auth_factory.ts
import { SharePointConfig, SharePointSecretConfig, SharePointCertConfig } from './config';
import { getSharePointHeaders as getSharePointHeadersSecret, getRequestDigest as getRequestDigestSecret, AuthHeaders } from './auth';
import { getSharePointHeaders as getSharePointHeadersCert, getRequestDigest as getRequestDigestCert } from './azure_cert_auth';

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
 * @param url The SharePoint site URL
 * @param headers The authentication headers
 * @returns Request digest value
 */
export async function getRequestDigest(url: string, headers: AuthHeaders): Promise<string> {
    // Both authentication methods use the same approach to get the request digest
    // We can use either one, but we'll use the secret one to be consistent
    return getRequestDigestSecret(url, headers);
}

export default {
    getSharePointHeaders,
    getRequestDigest
};