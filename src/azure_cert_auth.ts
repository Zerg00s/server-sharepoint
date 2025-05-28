// src/azure_cert_auth.ts - Fixed version
import axios from 'axios';
import * as fs from 'fs';
import * as crypto from 'crypto';
import * as path from 'path';
import * as forge from 'node-forge';
import { ConfidentialClientApplication } from '@azure/msal-node';
import { SharePointCertConfig } from './config';
import { AuthHeaders } from './auth';

/**
 * Get certificate from Windows Certificate Store using PowerShell
 */
async function getCertificateFromStore(thumbprint: string, password: string): Promise<Buffer | null> {
    try {
        const tempCertPath = path.resolve(process.cwd(), `${thumbprint}_temp.pfx`);
        
        // Clean up any existing temp file
        if (fs.existsSync(tempCertPath)) {
            fs.unlinkSync(tempCertPath);
        }
        
        // Create a single-line PowerShell command to avoid multiline parsing issues
        const psCommand = `$cert = Get-ChildItem -Path Cert:\\CurrentUser\\My | Where-Object { $_.Thumbprint -eq '${thumbprint}' }; if ($cert) { $password = ConvertTo-SecureString -String '${password}' -Force -AsPlainText; Export-PfxCertificate -Cert $cert -FilePath '${tempCertPath}' -Password $password; Write-Output 'Certificate exported successfully' } else { Write-Error 'Certificate not found in store'; exit 1 }`;
        
        console.error('Exporting certificate from Windows Certificate Store...');
        console.error(`Executing: powershell -Command "${psCommand}"`);
        
        const result = require('child_process').execSync(`powershell -Command "${psCommand}"`, {
            stdio: 'pipe',
            encoding: 'utf8'
        });
        
        console.error('PowerShell output:', result);
        
        if (fs.existsSync(tempCertPath)) {
            const certBuffer = fs.readFileSync(tempCertPath);
            // Clean up temp file
            fs.unlinkSync(tempCertPath);
            console.error('Certificate successfully exported from Windows Certificate Store');
            return certBuffer;
        }
        
        return null;
    } catch (error) {
        console.error('Failed to export certificate from store:', error);
        return null;
    }
}

/**
 * Find certificate file in common locations
 */
function findCertificateFile(thumbprint: string): string | null {
    const userDocumentsPath = path.resolve(process.env.USERPROFILE || '', 'Documents');
    const certPaths = [
        // Try with the exact thumbprint first
        path.resolve(process.cwd(), `${thumbprint}.pfx`),
        path.resolve(userDocumentsPath, `${thumbprint}.pfx`),
        // Try with the correct certificate name for XX9
        path.resolve(userDocumentsPath, 'SharePoint-Server-MCP-Cert-XX9.pfx'),
        // Try with common certificate names (older ones as fallback)
        path.resolve(userDocumentsPath, 'SharePoint-Server-MCP-Cert.pfx'),
        path.resolve(process.cwd(), 'SharePoint-Server-MCP-Cert-XX9.pfx'),
        path.resolve(process.cwd(), 'SharePoint-Server-MCP-Cert.pfx')
    ];
    
    for (const certPath of certPaths) {
        console.error(`Looking for certificate at: ${certPath}`);
        if (fs.existsSync(certPath)) {
            console.error(`Certificate found at: ${certPath}`);
            return certPath;
        }
    }
    
    return null;
}

/**
 * Create a JWT client assertion token signed with the certificate's private key
 */
async function createClientAssertion(
    clientId: string,
    tenantId: string,
    certificateThumbprint: string,
    certificatePassword: string
): Promise<string> {
    try {
        let certBuffer: Buffer | null = null;
        
        // First try to get certificate from Windows Certificate Store
        console.error('Attempting to get certificate from Windows Certificate Store...');
        certBuffer = await getCertificateFromStore(certificateThumbprint, certificatePassword);
        
        // If that fails, try to find certificate file on disk
        if (!certBuffer) {
            console.error('Certificate not found in store, looking for file on disk...');
            const certPath = findCertificateFile(certificateThumbprint);
            
            if (!certPath) {
                throw new Error(`Certificate not found in Windows Certificate Store or file system. Expected thumbprint: ${certificateThumbprint}`);
            }
            
            certBuffer = fs.readFileSync(certPath);
        }
        
        // Parse the certificate
        console.error('Parsing certificate...');
        const certData = forge.util.createBuffer(certBuffer.toString('binary'));
        const p12Asn1 = forge.asn1.fromDer(certData);
        
        // Use the provided password to decrypt
        console.error(`Attempting to decrypt certificate with provided password...`);
        const p12 = forge.pkcs12.pkcs12FromAsn1(p12Asn1, certificatePassword);
        
        // Extract private key
        let privateKey: any = null;
        let certificate: any = null;
        
        p12.safeContents.forEach(safeContent => {
            safeContent.safeBags.forEach(safeBag => {
                if (safeBag.type === forge.pki.oids.pkcs8ShroudedKeyBag && safeBag.key) {
                    privateKey = safeBag.key;
                } else if (safeBag.type === forge.pki.oids.keyBag && safeBag.key) {
                    privateKey = safeBag.key;
                } else if (safeBag.type === forge.pki.oids.certBag && safeBag.cert) {
                    certificate = safeBag.cert;
                }
            });
        });
        
        if (!privateKey) {
            throw new Error('Private key not found in certificate');
        }
        
        if (!certificate) {
            throw new Error('Certificate not found in PKCS#12 file');
        }
        
        console.error('Successfully extracted private key and certificate');
        
        // Create JWT header and payload
        const tokenEndpoint = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
        
        // Convert certificate to proper x5t format
        const certDer = forge.asn1.toDer(forge.pki.certificateToAsn1(certificate)).getBytes();
        const certBuffer2 = Buffer.from(certDer, 'binary');
        const certHash = crypto.createHash('sha1').update(certBuffer2).digest();
        const x5t = certHash.toString('base64url');
        
        const header = {
            alg: 'RS256',
            typ: 'JWT',
            x5t: x5t
        };
        
        const now = Math.floor(Date.now() / 1000);
        const payload = {
            aud: tokenEndpoint,
            iss: clientId,
            sub: clientId,
            jti: crypto.randomUUID(),
            nbf: now,
            exp: now + 600
        };
        
        // Create and sign JWT
        const encodedHeader = Buffer.from(JSON.stringify(header)).toString('base64url');
        const encodedPayload = Buffer.from(JSON.stringify(payload)).toString('base64url');
        const signingInput = `${encodedHeader}.${encodedPayload}`;
        
        // Sign using forge
        const md = forge.md.sha256.create();
        md.update(signingInput);
        const signature = privateKey.sign(md);
        const signatureBase64url = Buffer.from(signature, 'binary').toString('base64url');
        
        const jwtToken = `${signingInput}.${signatureBase64url}`;
        
        console.error('Successfully created client assertion JWT');
        return jwtToken;
        
    } catch (error) {
        console.error('Error creating client assertion:', error);
        if (error instanceof Error && error.message.includes('wrong password')) {
            throw new Error(`Certificate password is incorrect. Please verify the password for certificate ${certificateThumbprint}`);
        }
        throw error;
    }
}

/**
 * Authenticate with SharePoint using Azure AD certificate authentication
 */
export async function getSharePointHeaders(url: string, config: SharePointCertConfig): Promise<AuthHeaders> {
    try {
        console.error("Authenticating with SharePoint using Azure AD certificate...");
        console.error(`Client ID: ${config.clientId.substring(0, 8)}...`);
        console.error(`Certificate Thumbprint: ${config.certificateThumbprint}`);
        console.error(`Tenant ID: ${config.tenantId.substring(0, 8)}...`);
        
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
        
        // Get SharePoint resource URL
        const resourceUrl = new URL(url);
        const resource = `${resourceUrl.protocol}//${resourceUrl.hostname}`;
        console.error(`Target resource: ${resource}`);
        
        // Create client assertion
        const clientAssertion = await createClientAssertion(
            config.clientId,
            config.tenantId,
            config.certificateThumbprint,
            config.certificatePassword
        );
        
        // Request access token
        const tokenUrl = `https://login.microsoftonline.com/${config.tenantId}/oauth2/v2.0/token`;
        
        const formData = new URLSearchParams();
        formData.append('grant_type', 'client_credentials');
        formData.append('client_id', config.clientId);
        formData.append('client_assertion_type', 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer');
        formData.append('client_assertion', clientAssertion);
        formData.append('scope', `${resource}/.default`);
        
        console.error('Requesting access token from Azure AD...');
        
        const response = await axios.post(tokenUrl, formData.toString(), {
            headers: {
                'Content-Type': 'application/x-www-form-urlencoded'
            },
            timeout: 30000
        });
        
        if (!response.data || !response.data.access_token) {
            throw new Error('No access token received from Azure AD');
        }
        
        console.error('Successfully obtained access token');
        
        // Return headers with the access token
        const headers: AuthHeaders = {
            'Authorization': `Bearer ${response.data.access_token}`,
            'Accept': 'application/json;odata=verbose'
        };
        
        return headers;
        
    } catch (error) {
        console.error("SharePoint certificate authentication failed");
        
        let errorMessage = "";
        if (error instanceof Error) {
            errorMessage = error.message;
            console.error("Error details:", error.stack);
        } else {
            errorMessage = String(error);
        }
        
        // Provide specific guidance based on error type
        if (errorMessage.includes('wrong password')) {
            errorMessage += ". Please check your AZURE_APPLICATION_CERTIFICATE_PASSWORD environment variable.";
        } else if (errorMessage.includes('Certificate not found')) {
            errorMessage += ". Please ensure the certificate is installed in your Windows Certificate Store or the PFX file exists.";
        } else if (errorMessage.includes('invalid_client')) {
            errorMessage += ". Please check your AZURE_APPLICATION_ID and ensure the app registration is correct.";
        }
        
        throw new Error(`SharePoint authentication failed: ${errorMessage}`);
    }
}

export default {
    getSharePointHeaders
};