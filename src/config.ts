// src/config.ts
import * as dotenv from 'dotenv';

// Load environment variables from .env file
dotenv.config();

// Parse command line arguments
export function parseCliArgs(): Record<string, string> {
    return process.argv.slice(2).reduce((acc: Record<string, string>, arg) => {
        const [key, value] = arg.split('=');
        if (key && value) {
            acc[key.replace(/^--/, '')] = value;
        }
        return acc;
    }, {});
}

// Configuration interface for client secret auth
export interface SharePointSecretConfig {
    clientId: string;
    clientSecret: string;
    tenantId: string;
    siteUrl?: string;
    authType: 'secret';
}

// Configuration interface for certificate auth
export interface SharePointCertConfig {
    clientId: string;
    certificateThumbprint: string;
    certificatePassword: string;
    tenantId: string;
    siteUrl?: string;
    authType: 'certificate';
}

// Union type for both authentication methods
export type SharePointConfig = SharePointSecretConfig | SharePointCertConfig;

// Load configuration
export function loadConfig(): SharePointConfig {
    const args = parseCliArgs();

    // Check if Azure certificate auth variables are present
    const azureAppId = args.azureAppId || process.env.AZURE_APPLICATION_ID || '';
    const azureCertThumbprint = args.azureCertThumbprint || process.env.AZURE_APPLICATION_CERTIFICATE_THUMBPRINT || '';
    const azureCertPassword = args.azureCertPassword || process.env.AZURE_APPLICATION_CERTIFICATE_PASSWORD || '';
    
    // Check if SharePoint client secret auth variables are present
    const spClientId = args.clientId || process.env.SHAREPOINT_CLIENT_ID || '';
    const spClientSecret = args.clientSecret || process.env.SHAREPOINT_CLIENT_SECRET || '';
    
    // Common settings
    const tenantId = args.tenantId || process.env.M365_TENANT_ID || '';
    const siteUrl = args.siteUrl || process.env.SHAREPOINT_SITE_URL || '';
    
    // Determine which authentication method to use based on available credentials
    const useAzureCert = Boolean(azureAppId && azureCertThumbprint && azureCertPassword);
    const useClientSecret = Boolean(spClientId && spClientSecret);
    
    console.error(`Azure Certificate Auth Credentials Available: ${useAzureCert ? 'Yes' : 'No'}`);
    console.error(`Client Secret Auth Credentials Available: ${useClientSecret ? 'Yes' : 'No'}`);
    
    // Prefer Azure certificate auth if both are available
    if (useAzureCert) {
        console.error("Using Azure AD Certificate Authentication");
        return {
            clientId: azureAppId,
            certificateThumbprint: azureCertThumbprint,
            certificatePassword: azureCertPassword,
            tenantId: tenantId,
            siteUrl: siteUrl,
            authType: 'certificate'
        };
    } else {
        console.error("Using Client Secret Authentication");
        return {
            clientId: spClientId,
            clientSecret: spClientSecret,
            tenantId: tenantId,
            siteUrl: siteUrl,
            authType: 'secret'
        };
    }
}

// Validate that required config values are present
export function validateConfig(config: SharePointConfig): boolean {
    const { clientId, tenantId } = config;
    let isValid = Boolean(clientId && tenantId);
    
    if (config.authType === 'secret') {
        isValid = isValid && Boolean(config.clientSecret);
    } else if (config.authType === 'certificate') {
        isValid = isValid && Boolean(config.certificateThumbprint && config.certificatePassword);
    }
    
    if (!isValid) {
        console.error("ERROR: Missing SharePoint credentials!");
        if (config.authType === 'secret') {
            console.error("Provide client secret auth via environment variables:");
            console.error("SHAREPOINT_CLIENT_ID, SHAREPOINT_CLIENT_SECRET, M365_TENANT_ID");
            console.error("Or CLI arguments: --clientId=xxx --clientSecret=yyy --tenantId=zzz");
        } else {
            console.error("Provide certificate auth via environment variables:");
            console.error("AZURE_APPLICATION_ID, AZURE_APPLICATION_CERTIFICATE_THUMBPRINT, AZURE_APPLICATION_CERTIFICATE_PASSWORD, M365_TENANT_ID");
            console.error("Or CLI arguments: --azureAppId=xxx --azureCertThumbprint=yyy --azureCertPassword=zzz --tenantId=aaa");
        }
    } else {
        console.error(`✅ SharePoint credentials loaded (${config.authType} authentication).`);
    }
    
    return isValid;
}

// Get the SharePoint configuration, combining loading and validation
export function getSharePointConfig(): SharePointConfig {
    const config = loadConfig();
    validateConfig(config);
    return config;
}

// Export as default
export default getSharePointConfig;