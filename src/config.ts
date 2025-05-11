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

// Configuration interface
export interface SharePointConfig {
    clientId: string;
    clientSecret: string;
    tenantId: string;
    siteUrl?: string;
}

// Load and validate configuration
export function loadConfig(): SharePointConfig {
    const args = parseCliArgs();

    // Priority: CLI args → fallback: ENV
    const config = {
        // Use with auth.ts
        clientId: args.clientId || process.env.SHAREPOINT_CLIENT_ID || '',
        clientSecret: args.clientSecret || process.env.SHAREPOINT_CLIENT_SECRET || '',
        AzureAppId: args.clientId || process.env.AZURE_APPLICATION_ID || '',
        
        //TODO: Use with azure_cert_auth.ts (azure_cert_auth.ts needs to be implemented based on the sample from get-site-title.ts)
        // TODO: azure_cert_auth.ts should similar to how auth.ts is implemented, in terms so that the MCP tools functions can be used with both
        // auth.ts and azure_cert_auth.ts, depending on the authentication method used
        // if clientId is set - it's auuth.ts
        // if AzureAppId is set - it's azure_cert_auth.ts
        AzureAppCertificateThumbprint: args.clientSecret || process.env.AZURE_APPLICATION_CERTIFICATE_THUMBPRINT || '',
        AzureAppCertificatePassword: args.clientSecret || process.env.AZURE_APPLICATION_CERTIFICATE_PASSWORD || '',
        
        tenantId: args.tenantId || process.env.SHAREPOINT_TENANT_ID || '',
        siteUrl: args.siteUrl || process.env.SHAREPOINT_SITE_URL || ''
    };

    return config;
}

// Validate that required config values are present
export function validateConfig(config: SharePointConfig): boolean {
    const { clientId, clientSecret, tenantId } = config;
    
    const isValid = Boolean(clientId && clientSecret && tenantId);
    
    if (!isValid) {
        console.error("ERROR: Missing SharePoint credentials!");
        console.error("Provide via environment variables or CLI arguments like:");
        console.error("--clientId=xxx --clientSecret=yyy --tenantId=zzz");
    } else {
        console.error("✅ SharePoint credentials loaded.");
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