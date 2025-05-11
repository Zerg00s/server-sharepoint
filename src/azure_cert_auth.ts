// src/azure_cert_auth.ts
import axios from 'axios';
import * as fs from 'fs';
import * as crypto from 'crypto';
import * as path from 'path';
import * as forge from 'node-forge';
import { ConfidentialClientApplication } from '@azure/msal-node';
import request from 'request-promise';
import { SharePointCertConfig } from './config';
import { AuthHeaders } from './auth';

/**
 * Create a JWT client assertion token signed with the certificate's private key
 * @param clientId The application client ID
 * @param tenantId The tenant ID
 * @param certificateThumbprint The certificate thumbprint
 * @param certificatePassword The certificate password
 * @returns The signed JWT token
 */
async function createClientAssertion(
    clientId: string,
    tenantId: string,
    certificateThumbprint: string,
    certificatePassword: string
): Promise<string> {
    try {
        // Look for certificate files in potential locations
        // First try the Documents folder with the default certificate name
        const userDocumentsPath = path.resolve(process.env.USERPROFILE || '', 'Documents');
        const defaultCertName = 'SharePoint-Server-MCP-Cert';
        const certPaths = [
            path.resolve(process.cwd(), `${certificateThumbprint}.pfx`),
            path.resolve(userDocumentsPath, `${defaultCertName}.pfx`),
            path.resolve(process.cwd(), `${defaultCertName}.pfx`)
        ];
        
        let certPath = '';
        let certExists = false;
        
        // Try each possible certificate path
        for (const potentialPath of certPaths) {
            console.error(`Looking for certificate at: ${potentialPath}`);
            if (fs.existsSync(potentialPath)) {
                certPath = potentialPath;
                certExists = true;
                console.error(`Certificate found at: ${certPath}`);
                break;
            }
        }
        
        // If certificate file doesn't exist in any of the expected locations
        if (!certExists) {
            console.error('Certificate not found in file system. Trying to export from Windows Certificate Store...');
            
            try {
                // Attempt to export the certificate from the Windows Certificate Store using PowerShell
                const tempCertPath = path.resolve(process.cwd(), `${certificateThumbprint}_temp.pfx`);
                const psCommand = `
                    # First check if the certificate exists
                    $cert = Get-ChildItem -Path Cert:\\CurrentUser\\My | Where-Object { $_.Thumbprint -eq "${certificateThumbprint}" }
                    
                    # If not found by thumbprint, try to find by subject
                    if (-not $cert) {
                        $cert = Get-ChildItem -Path Cert:\\CurrentUser\\My | Where-Object { $_.Subject -like "*SharePoint-Server-MCP-Cert*" }
                        if ($cert) {
                            Write-Output "Found certificate by subject: $($cert.Thumbprint)"
                        }
                    }
                    
                    if ($cert) {
                        $password = ConvertTo-SecureString -String "${certificatePassword}" -Force -AsPlainText
                        Export-PfxCertificate -Cert $cert -FilePath "${tempCertPath}" -Password $password
                        Write-Output "Certificate exported successfully"
                    } else {
                        Write-Error "Certificate not found in store"
                        exit 1
                    }
                `;
                
                console.error('Executing PowerShell to export certificate from store...');
                require('child_process').execSync(`powershell -Command "${psCommand}"`, {stdio: 'inherit'});
                
                if (fs.existsSync(tempCertPath)) {
                    console.error(`Certificate successfully exported to: ${tempCertPath}`);
                    certPath = tempCertPath;
                    certExists = true;
                }
            } catch (certStoreError) {
                console.error('Failed to export certificate from Windows Certificate Store:');
                console.error(certStoreError instanceof Error ? certStoreError.message : String(certStoreError));
            }
            
            // If still no certificate found, throw error
            if (!certExists) {
                throw new Error('Certificate not found in file system or Windows Certificate Store');
            }
        }
        
        // Read the certificate file
        const certBuffer = fs.readFileSync(certPath);
        const certData = forge.util.createBuffer(certBuffer.toString('binary'));
        
        // Parse the certificate with the password
        const p12Asn1 = forge.asn1.fromDer(certData);
        const p12 = forge.pkcs12.pkcs12FromAsn1(p12Asn1, certificatePassword);
        
        // Extract private key from the certificate (using any type to avoid TypeScript errors)
        let privateKey: any = null;
        p12.safeContents.forEach(safeContent => {
            safeContent.safeBags.forEach(safeBag => {
                if (safeBag.type === forge.pki.oids.pkcs8ShroudedKeyBag && safeBag.key) {
                    privateKey = safeBag.key;
                } else if (safeBag.type === forge.pki.oids.keyBag && safeBag.key) {
                    privateKey = safeBag.key;
                }
            });
        });
        
        if (!privateKey) {
            throw new Error('Private key not found in certificate');
        }
        
        // The audience should be the token endpoint
        const tokenEndpoint = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
        
        // Helper function to convert thumbprint to base64url format for x5t claim
        function thumbprintToBase64Url(thumbprint: string): string {
            // Convert the hexadecimal thumbprint to a binary buffer
            const buffer = Buffer.from(thumbprint.replace(/:/g, ''), 'hex');
            // Convert to base64url format
            return buffer.toString('base64url');
        }
        
        // Create the JWT header with properly encoded thumbprint
        const header = {
            alg: 'RS256',
            typ: 'JWT',
            x5t: thumbprintToBase64Url(certificateThumbprint), // Thumbprint in base64url format
        };
        
        // Create the JWT payload
        const now = Math.floor(Date.now() / 1000);
        const payload = {
            aud: tokenEndpoint,
            iss: clientId,
            sub: clientId,
            jti: crypto.randomUUID(),
            nbf: now,
            exp: now + 600 // Token valid for 10 minutes
        };
        
        // Encode the header and payload
        const encodedHeader = Buffer.from(JSON.stringify(header)).toString('base64url');
        const encodedPayload = Buffer.from(JSON.stringify(payload)).toString('base64url');
        
        // Create the signing input
        const signingInput = `${encodedHeader}.${encodedPayload}`;
        
        // Sign the JWT using the private key (using any type to avoid TypeScript errors)
        const signature = forge.md.sha256.create();
        signature.update(signingInput);
        const signatureBytes = privateKey.sign(signature);
        const signatureBase64 = forge.util.encode64(signatureBytes);
        
        // Create the JWT token
        const jwtToken = `${signingInput}.${Buffer.from(signatureBase64, 'base64').toString('base64url')}`;
        
        return jwtToken;
    } catch (error) {
        console.error('Error creating client assertion:');
        console.error(error instanceof Error ? error.message : String(error));
        throw error;
    }
}

/**
 * Authenticate with SharePoint using Azure AD certificate authentication
 * @param url The SharePoint site URL
 * @param config The SharePoint configuration
 * @returns Authentication headers for API requests
 */
export async function getSharePointHeaders(url: string, config: SharePointCertConfig): Promise<AuthHeaders> {
    try {
        console.error("Authenticating with SharePoint using Azure AD certificate...");
        
        // Validate required configuration
        if (!config.clientId) {
            throw new Error("Azure Application ID (Client ID) is missing from configuration");
        }
        
        if (!config.certificateThumbprint) {
            throw new Error("Azure Application Certificate Thumbprint is missing from configuration");
        }
        
        if (!config.certificatePassword) {
            throw new Error("Azure Application Certificate Password is missing from configuration");
        }
        
        if (!config.tenantId) {
            throw new Error("Azure Tenant ID is missing from configuration");
        }
        
        console.error(`Attempting to authenticate to URL: ${url}`);
        console.error(`Using client ID: ${config.clientId.substring(0, 5)}...`);
        console.error(`Using certificate thumbprint: ${config.certificateThumbprint.substring(0, 5)}...`);
        console.error(`Using tenant ID: ${config.tenantId.substring(0, 5)}...`);
        
        // Get SharePoint resource URL
        const resourceUrl = new URL(url);
        const resource = `${resourceUrl.protocol}//${resourceUrl.hostname}`;
        console.error(`Using resource: ${resource}`);
        
        let accessToken = '';
        
        // Try using MSAL first as it handles certificate authentication more robustly
        try {
            console.error('Attempting authentication using MSAL...');
            
            // For MSAL, we need to read the PFX certificate and provide it with the password
            // Let's try to find the certificate file first
            const userDocumentsPath = path.resolve(process.env.USERPROFILE || '', 'Documents');
            const defaultCertName = 'SharePoint-Server-MCP-Cert';
            const certPath = path.resolve(userDocumentsPath, `${defaultCertName}.pfx`);
            
            let msalConfig: any;
            
            if (fs.existsSync(certPath)) {
                console.error(`Certificate found at: ${certPath}`);
                
                // Read the PFX file
                const certBuffer = fs.readFileSync(certPath);
                
                msalConfig = {
                    auth: {
                        clientId: config.clientId,
                        authority: `https://login.microsoftonline.com/${config.tenantId}`,
                        clientCertificate: {
                            thumbprint: config.certificateThumbprint,
                            privateKey: certBuffer.toString('base64')
                        }
                    }
                };
            } else {
                // Try to use certificate from the Windows Certificate Store
                console.error('Certificate file not found, trying to use from Windows Certificate Store');
                msalConfig = {
                    auth: {
                        clientId: config.clientId,
                        authority: `https://login.microsoftonline.com/${config.tenantId}`,
                        clientCertificate: {
                            thumbprint: config.certificateThumbprint,
                            privateKey: null  // Will use the certificate from the Windows store
                        }
                    }
                };
            }
            
            const msalClient = new ConfidentialClientApplication(msalConfig);
            
            // Get token for SharePoint
            const tokenResponse = await msalClient.acquireTokenByClientCredential({
                scopes: [`${resource}/.default`]
            });
            
            if (tokenResponse && tokenResponse.accessToken) {
                console.error('Authentication successful using MSAL with certificate');
                accessToken = tokenResponse.accessToken;
            } else {
                throw new Error('No access token returned from MSAL');
            }
        } catch (msalError) {
            console.error('MSAL authentication failed:');
            console.error(msalError instanceof Error ? msalError.message : String(msalError));
            console.error('Falling back to manual certificate authentication...');
            
            // Create client assertion with our manual method
            const clientAssertion = await createClientAssertion(
                config.clientId, 
                config.tenantId, 
                config.certificateThumbprint, 
                config.certificatePassword
            );
            
            // Format the token endpoint URL
            const tokenUrl = `https://login.microsoftonline.com/${config.tenantId}/oauth2/v2.0/token`;
            
            // Prepare the form data for token request
            const formData = new URLSearchParams();
            formData.append('grant_type', 'client_credentials');
            formData.append('client_id', config.clientId);
            formData.append('client_assertion_type', 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer');
            formData.append('client_assertion', clientAssertion);
            formData.append('scope', `${resource}/.default`);
            
            console.error('Requesting access token...');
            
            try {
                // Make the token request
                const response = await axios.post(tokenUrl, formData.toString(), {
                    headers: {
                        'Content-Type': 'application/x-www-form-urlencoded'
                    }
                });
                
                // Check response for access token
                if (response.data && response.data.access_token) {
                    accessToken = response.data.access_token;
                    console.error('Access token obtained successfully');
                } else {
                    console.error('Token response did not contain an access token:', response.data);
                    throw new Error('No access token in response');
                }
            } catch (tokenError: any) {
                if (tokenError.response) {
                    // The request was made and the server responded with a status code outside the 2xx range
                    console.error(`Token request failed with status: ${tokenError.response.status}`);
                    console.error('Error details:');
                    console.error(tokenError.response.data);
                    
                    if (tokenError.response.data && tokenError.response.data.error_description) {
                        throw new Error(`Token request failed: ${tokenError.response.data.error_description}`);
                    }
                }
                throw tokenError;
            }
        }
        
        // Define headers with the access token
        const headers: AuthHeaders = {
            'Authorization': `Bearer ${accessToken}`,
            'Accept': 'application/json;odata=verbose'
        };
        
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
                errorDetail = "Invalid client credentials (client ID or certificate)";
            } else if (error.message.includes("invalid_grant")) {
                errorDetail = "Invalid grant (tenant ID may be incorrect)";
            } else if (error.message.toLowerCase().includes("certificate") || 
                      error.message.includes("thumbprint")) {
                errorDetail = "Certificate error - check thumbprint and password";
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