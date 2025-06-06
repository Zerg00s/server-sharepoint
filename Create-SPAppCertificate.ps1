# Create-SPAppCertificate.ps1
# This script creates a self-signed certificate and registers an Azure AD application
# with SharePoint permissions, then adds the certificate to the application.

#Requires -Modules Microsoft.Graph.Applications, Microsoft.Graph.Authentication

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$CertPassword = "TemporaryP@ssw0rd",
    [string]$AppName = "SharePoint-Server-MCP",
    
    [Parameter(Mandatory = $false)]
    [string]$CertName = "SharePoint-Server-MCP-Cert",
    
    [Parameter(Mandatory = $false)]
    [string]$CertPath = "$env:USERPROFILE\Documents",
    
    [Parameter(Mandatory = $false)]
    [int]$CertValidityYears = 2,
    
    [Parameter(Mandatory = $false)]
    [string]$ConfigOutputPath = ".\SharePointApp-Config.xml"
)

function Write-Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [string]$ForegroundColor = "White"
    )
    
    Write-Host $Message -ForegroundColor $ForegroundColor
}

function Create-SelfSignedCertificate {
    param(
        [Parameter(Mandatory = $true)]
        [string]$CertName,
        
        [Parameter(Mandatory = $true)]
        [string]$CertPath,
        
        [Parameter(Mandatory = $true)]
        [int]$ValidityYears,
        
        [Parameter(Mandatory = $true)]
        [string]$Password
    )
    
    # Check if certificate with same name already exists
    $existingCert = Get-ChildItem -Path Cert:\CurrentUser\My | Where-Object { $_.Subject -eq "CN=$CertName" }
    
    if ($existingCert) {
        Write-Log "Certificate with name '$CertName' already exists with thumbprint: $($existingCert.Thumbprint)" -ForegroundColor Yellow
        return $existingCert
    }
    
    Write-Log "Creating self-signed certificate: $CertName" -ForegroundColor Cyan
    
    # Create self-signed certificate
    $notAfter = (Get-Date).AddYears($ValidityYears)
    $certParams = @{
        Subject = "CN=$CertName"
        NotAfter = $notAfter
        CertStoreLocation = "Cert:\CurrentUser\My"
        KeyExportPolicy = "Exportable"
        KeySpec = "Signature"
        Provider = "Microsoft Enhanced RSA and AES Cryptographic Provider"
        HashAlgorithm = "SHA256"
    }
    
    $certificate = New-SelfSignedCertificate @certParams
    
    # Export certificate to PFX (with private key)
    $securePassword = ConvertTo-SecureString -String $Password -Force -AsPlainText
    $pfxPath = Join-Path -Path $CertPath -ChildPath "$CertName.pfx"
    Export-PfxCertificate -Cert $certificate -FilePath $pfxPath -Password $securePassword | Out-Null
    
    # Export certificate to CER (public key only)
    $cerPath = Join-Path -Path $CertPath -ChildPath "$CertName.cer"
    Export-Certificate -Cert $certificate -FilePath $cerPath -Type CERT | Out-Null
    
    Write-Log "Certificate created and exported to:" -ForegroundColor Green
    Write-Log "  - PFX (with private key): $pfxPath" -ForegroundColor Green
    Write-Log "  - CER (public key only): $cerPath" -ForegroundColor Green
    Write-Log "  - Certificate password: $Password" -ForegroundColor Green
    
    return $certificate
}

function Register-AzureApplication {
    param(
        [Parameter(Mandatory = $true)]
        [string]$AppName,
        
        [Parameter(Mandatory = $true)]
        [System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate
    )
    
    # Check if app with same name already exists
    $existingApp = Get-MgApplication -Filter "DisplayName eq '$AppName'" -ErrorAction SilentlyContinue
    
    if ($existingApp) {
        Write-Log "Application with name '$AppName' already exists with ID: $($existingApp.AppId)" -ForegroundColor Yellow
        return $existingApp
    }
    
    Write-Log "Registering new Azure AD application: $AppName" -ForegroundColor Cyan
    
    # Create application without redirect URI since this is a daemon/service app
    $appParams = @{
        DisplayName = $AppName
        SignInAudience = "AzureADMyOrg"
    }
    
    # Create application
    $application = New-MgApplication @appParams
    
    # Now add the certificate to the application
    $keyCredential = @{
        Type = "AsymmetricX509Cert"
        Usage = "Verify"
        Key = $Certificate.GetRawCertData()
        DisplayName = "$AppName Certificate"
        EndDateTime = $Certificate.NotAfter
        StartDateTime = $Certificate.NotBefore
    }
    
    # Update the application with the certificate
    Update-MgApplication -ApplicationId $application.Id -KeyCredentials @($keyCredential)
    
    # Create service principal for the application
    $servicePrincipal = New-MgServicePrincipal -AppId $application.AppId
    
    Write-Log "Application registered with App ID: $($application.AppId)" -ForegroundColor Green
    Write-Log "Service Principal created with Object ID: $($servicePrincipal.Id)" -ForegroundColor Green
    
    return $application
}

function Add-SharePointPermissions {
    param(
        [Parameter(Mandatory = $true)]
        [Microsoft.Graph.PowerShell.Models.MicrosoftGraphApplication]$Application
    )
    
    Write-Log "Adding SharePoint permissions to application" -ForegroundColor Cyan
    
    # SharePoint API information
    # SharePoint Online API ID is always this value
    $sharepointApiId = "00000003-0000-0ff1-ce00-000000000000"
    
    # Generate admin consent URL early
    $tenantId = (Get-MgContext).TenantId
    $adminConsentUrl = "https://login.microsoftonline.com/$tenantId/adminconsent?client_id=$($Application.AppId)"
    
    # Check if SharePoint API service principal exists
    $sharepointSp = Get-MgServicePrincipal -Filter "appId eq '$sharepointApiId'"
    
    if (-not $sharepointSp) {
        Write-Log "SharePoint API service principal not found. Make sure you have access to it." -ForegroundColor Red
        return [string]$adminConsentUrl
    }
    
    # Find the Sites.FullControl.All permission
    $sitesFullControlPermission = $sharepointSp.AppRoles | Where-Object { $_.Value -eq "Sites.FullControl.All" }
    
    if (-not $sitesFullControlPermission) {
        Write-Log "SharePoint permission 'Sites.FullControl.All' not found." -ForegroundColor Red
        return [string]$adminConsentUrl
    }
    
    # Define the required resource access
    $resourceAccess = @{
        Id = $sitesFullControlPermission.Id
        Type = "Role"
    }
    
    $requiredResourceAccess = @{
        ResourceAppId = $sharepointApiId
        ResourceAccess = @($resourceAccess)
    }
    
    # Get existing required resource access
    $existingResourceAccess = @($Application.RequiredResourceAccess)
    
    # Check if SharePoint permission already exists
    $spPermissionExists = $existingResourceAccess | Where-Object { $_.ResourceAppId -eq $sharepointApiId }
    
    if ($spPermissionExists) {
        Write-Log "SharePoint permissions already exist on this application. Updating..." -ForegroundColor Yellow
        # Filter out existing SharePoint permissions
        $existingResourceAccess = $existingResourceAccess | Where-Object { $_.ResourceAppId -ne $sharepointApiId }
    }
    
    # Add the new SharePoint permission
    $existingResourceAccess += $requiredResourceAccess
    
    # Update the application with the new permissions
    Update-MgApplication -ApplicationId $Application.Id -RequiredResourceAccess $existingResourceAccess
    
    Write-Log "Added 'Sites.FullControl.All' permission to application" -ForegroundColor Green
    Write-Log "IMPORTANT: You still need to grant admin consent for this permission!" -ForegroundColor Yellow

    Write-Log "`n================== ADMIN CONSENT REQUIRED ==================" -ForegroundColor Yellow
    Write-Log "Opening browser for admin consent in 5 seconds..." -ForegroundColor Cyan
    Write-Log "After granting consent:" -ForegroundColor Cyan
    Write-Log "  1. You'll see an error about 'no reply address'" -ForegroundColor White
    Write-Log "  2. This is NORMAL - consent has been granted!" -ForegroundColor Green
    Write-Log "  3. Simply close the browser tab" -ForegroundColor White
    Write-Log "============================================================" -ForegroundColor Yellow
    
    Start-Sleep -Seconds 5
    Start-Process $adminConsentUrl
    
    Write-Log "`nAdmin Consent URL: $adminConsentUrl" -ForegroundColor Cyan
    
    # Wait for user to complete consent
    Write-Log "`nPress Enter after you have granted admin consent and closed the browser tab..." -ForegroundColor Yellow
    Read-Host | Out-Null
    
    Write-Log "Admin consent process completed!" -ForegroundColor Green
    
    # Explicitly return string
    return [string]$adminConsentUrl
}

function Output-ConfigDetails {
    param(
        [Parameter(Mandatory = $true)]
        [Microsoft.Graph.PowerShell.Models.MicrosoftGraphApplication]$Application,
        
        [Parameter(Mandatory = $true)]
        [System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate,
        
        [Parameter(Mandatory = $true)]
        [string]$OutputPath,
        
        [Parameter(Mandatory = $true)]
        [string]$AdminConsentUrl,
        
        [Parameter(Mandatory = $true)]
        [string]$CertificatePassword
    )
    
    $tenantId = (Get-MgContext).TenantId
    
    # Create XML content
    $xmlContent = @"
<?xml version="1.0" encoding="utf-8"?>
<root>
  <!-- SharePoint Server MCP Configuration -->
  <!-- Generated: $(Get-Date) -->
  <graphtenantid>$tenantId</graphtenantid>
  <graphclientid>$($Application.AppId)</graphclientid>
  <graphcertificate>$($Certificate.Thumbprint)</graphcertificate>
  <certificatepassword>$CertificatePassword</certificatepassword>
  <!-- Admin Consent URL (open in browser and sign in as admin to grant permissions) -->
  <!-- $AdminConsentUrl -->
  <mailboxpermissions>Yes</mailboxpermissions>
  <mfadetails>Yes</mfadetails>
</root>
"@
    
    Set-Content -Path $OutputPath -Value $xmlContent
    
    # Also create a text file with complete information
    $txtPath = $OutputPath.Replace(".xml", ".txt")
    $txtContent = @"
# SharePoint App Configuration Details
# Generated: $(Get-Date)

# Azure AD Application Details
AppName = $($Application.DisplayName)
ClientID = $($Application.AppId)
TenantID = $tenantId

# Certificate Details
CertificateName = $($Certificate.Subject.Replace("CN=", ""))
CertificateThumbprint = $($Certificate.Thumbprint)
CertificatePassword = $CertificatePassword
CertificateNotBefore = $($Certificate.NotBefore)
CertificateNotAfter = $($Certificate.NotAfter)

# Integration in server-sharepoint project
# Add these values to your config.xml or environment variables:

SHAREPOINT_CLIENT_ID = $($Application.AppId)
M365_TENANT_ID = $tenantId
SHAREPOINT_CERTIFICATE = $($Certificate.Thumbprint)
SHAREPOINT_CERTIFICATE_PASSWORD = $CertificatePassword

# Admin Consent URL (open in browser and sign in as admin to grant permissions)
AdminConsentURL = $AdminConsentUrl
"@
    
    Set-Content -Path $txtPath -Value $txtContent
    
    Write-Log "Configuration details saved to:" -ForegroundColor Green
    Write-Log "  - XML Config: $OutputPath" -ForegroundColor Green
    Write-Log "  - Text Details: $txtPath" -ForegroundColor Green
}

function Output-SampleMcpConfig {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TenantId,
        
        [Parameter(Mandatory = $true)]
        [string]$AppId,
        
        [Parameter(Mandatory = $true)]
        [string]$Thumbprint,
        
        [Parameter(Mandatory = $true)]
        [string]$Password,
        
        [Parameter(Mandatory = $false)]
        [string]$OutputPath = ".\claude_desktop_config.json"
    )

    # Create the JSON content manually for proper formatting
    $jsonContent = @"
{
  "mcpServers": {
    "server-sharepoint": {
      "command": "npx",
      "args": [
        "-y",
        "server-sharepoint"
      ],
      "env": {
        "M365_TENANT_ID": "$TenantId",
        "AZURE_APPLICATION_ID": "$AppId",
        "AZURE_APPLICATION_CERTIFICATE_THUMBPRINT": "$Thumbprint",
        "AZURE_APPLICATION_CERTIFICATE_PASSWORD": "$Password"
      }
    }
  }
}
"@

    # Write UTF8 without BOM
    $utf8NoBom = New-Object System.Text.UTF8Encoding $false
    [System.IO.File]::WriteAllText($OutputPath, $jsonContent, $utf8NoBom)

    Write-Log "Sample MCP config saved to: $OutputPath" -ForegroundColor Green
}

# Main execution
try {
    # Check if Microsoft Graph PowerShell is installed and user is logged in
    try {
        # Check if Microsoft Graph PowerShell module is installed
        if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
            Write-Log "Microsoft Graph PowerShell module is not installed." -ForegroundColor Yellow
            $installModule = Read-Host "Do you want to install it now? (Y/N)"
            if ($installModule -eq "Y" -or $installModule -eq "y") {
                Write-Log "Installing Microsoft Graph PowerShell module..." -ForegroundColor Cyan
                Install-Module Microsoft.Graph -Scope CurrentUser -Force
                Write-Log "Microsoft Graph PowerShell module installed successfully." -ForegroundColor Green
            } else {
                Write-Log "Microsoft Graph PowerShell module is required to run this script." -ForegroundColor Red
                exit
            }
        }
        
        # Try to get current context
        $graphContext = Get-MgContext -ErrorAction SilentlyContinue
        
        # Define required scopes
        $requiredScopes = @("Application.ReadWrite.All", "AppRoleAssignment.ReadWrite.All")
        
        # Check if user is logged in with the correct permissions
        $needsAuth = $false
        if (-not $graphContext) {
            Write-Log "Not logged in to Microsoft Graph." -ForegroundColor Yellow
            $needsAuth = $true
        } else {
            # Check permissions
            $missingScopes = $requiredScopes | Where-Object { $graphContext.Scopes -notcontains $_ }
            if ($missingScopes) {
                Write-Log "Missing required scopes: $($missingScopes -join ', ')" -ForegroundColor Yellow
                $needsAuth = $true
            }
        }
        
        # Authenticate if needed
        if ($needsAuth) {
            Write-Log "Authenticating with Microsoft Graph..." -ForegroundColor Cyan
            Connect-MgGraph -Scopes $requiredScopes
            
            # Verify connection was successful
            $graphContext = Get-MgContext -ErrorAction Stop
            if (-not $graphContext) {
                Write-Log "Authentication failed. Please try again." -ForegroundColor Red
                exit
            }
            Write-Log "Authentication successful!" -ForegroundColor Green
        }
    }
    catch {
        Write-Log "Error with Microsoft Graph PowerShell module or authentication." -ForegroundColor Red
        Write-Log $_.Exception.Message -ForegroundColor Red
        exit
    }
    
    Write-Log "Connected to Azure tenant: $($graphContext.TenantId)" -ForegroundColor Green
    
    # Create certificate
    $certificate = Create-SelfSignedCertificate -CertName $CertName -CertPath $CertPath -ValidityYears $CertValidityYears -Password $CertPassword
    
    # Register application
    $application = Register-AzureApplication -AppName $AppName -Certificate $certificate
    
    # Add SharePoint permissions
    $result = Add-SharePointPermissions -Application $application
    
    # Extract just the URL string from the result
    if ($result -is [array]) {
        $adminConsentUrl = $result[-1].ToString()
    } else {
        $adminConsentUrl = $result.ToString()
    }
    
    # Ensure we have a valid URL
    if (-not $adminConsentUrl -or $adminConsentUrl -notlike "https://*") {
        $adminConsentUrl = "https://login.microsoftonline.com/$($graphContext.TenantId)/adminconsent?client_id=$($application.AppId)"
    }
    
    # Output configuration details
    Output-ConfigDetails -Application $application -Certificate $certificate -OutputPath $ConfigOutputPath -AdminConsentUrl $adminConsentUrl -CertificatePassword $CertPassword

    # Output sample MCP config
    Output-SampleMcpConfig -TenantId $graphContext.TenantId `
                           -AppId $application.AppId `
                           -Thumbprint $certificate.Thumbprint `
                           -Password $CertPassword
    
    Write-Log "`nSetup complete!" -ForegroundColor Green
    Write-Log "1. If you haven't already, complete the admin consent process" -ForegroundColor Yellow
    Write-Log "2. Update your configuration with the values in the output files" -ForegroundColor Yellow
    Write-Log "3. The certificate is in your certificate store (CurrentUser\My) and exported to: $CertPath" -ForegroundColor Yellow
    
    # Also output the values for easy access in config.xml
    Write-Log "`nConfig Values:" -ForegroundColor Cyan
    Write-Log "<graphtenantid>$($graphContext.TenantId)</graphtenantid>" -ForegroundColor White
    Write-Log "<graphclientid>$($application.AppId)</graphclientid>" -ForegroundColor White
    Write-Log "<graphcertificate>$($certificate.Thumbprint)</graphcertificate>" -ForegroundColor White
    Write-Log "<certificatepassword>$CertPassword</certificatepassword>" -ForegroundColor White
}
catch {
    Write-Log "An error occurred:" -ForegroundColor Red
    Write-Log $_.Exception.Message -ForegroundColor Red
    Write-Log $_.ScriptStackTrace -ForegroundColor Red
}