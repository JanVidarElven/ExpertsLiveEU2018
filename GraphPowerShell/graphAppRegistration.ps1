# --------------------------------------------------------------------------------------------------------
# Name: graphAppRegistration.ps1
# Description: Create Azure AD App Registration for use against Microsoft Graph Api
# Created by: Jan Vidar Elven, Skill AS
# Last modified: 24.10.2018
# --------------------------------------------------------------------------------------------------------

# Connect to Azure AD
$tenantId = "elven.onmicrosoft.com"
Connect-AzureAD -TenantId $tenantId

# Prepate Microsoft Graph API permissions with Azure AD Service Principal for Microsoft Graph
$graphPrincipal = Get-AzureADServicePrincipal -All $true | Where-Object {$_.DisplayName -eq "Microsoft Graph"}
$graphAccess = New-Object -TypeName "Microsoft.Open.AzureAD.Model.RequiredResourceAccess"
$graphAccess.ResourceAppId = $graphPrincipal.AppId

# Helper for listing available Application Permissions
$graphPrincipal.AppRoles

# Helper for listing available Delegated Permissions
$graphPrincipal.Oauth2Permissions 

# Example: Get all Intune delegated permissions
$graphDelegatedPermissions = $graphPrincipal.Oauth2Permissions | Where-Object {$_.Value -like 'DeviceManagement*'}

# Example: Get all Guest Invite permissions
$graphDelegatedPermissions = $graphPrincipal.Oauth2Permissions | Where-Object {$_.Value -like 'User.Invite*'}

# Loop through and assign all delegated permissions
ForEach ($graphPermission in $graphDelegatedPermissions ) {
    Write-Host -ForegroundColor Green ("Adding Delegated Permission for " + $graphPermission.Value)
    $graphAddPermission = New-Object -TypeName "microsoft.open.azuread.model.resourceAccess" -ArgumentList $graphPermission.Id, "Scope"
    $graphAccess.ResourceAccess += $graphAddPermission
}

# Create App Registration for PowerShell Graph Api with specified permissions
$nativeAppReplyUri = "urn:ietf:wg:oauth:2.0:oob"
$graphApp = New-AzureADApplication -DisplayName "PowerShell Graph Api" -PublicClient $true -ReplyUrls $nativeAppReplyUri  -RequiredResourceAccess @($graphAccess)

# Add Service Principal to App. Note tag: https://docs.microsoft.com/en-us/powershell/module/azuread/new-azureadserviceprincipal?view=azureadps-2.0
$graphAppSpn = New-AzureADServicePrincipal -AppId $graphApp.AppId -Tags @("WindowsAzureActiveDirectoryIntegratedApp")

# TODO:
# 1. Remember to Grant Admin Consent in the Azure AD Portal for delegated permissions that requires it (most do)
# 2. Note down Client Id (App Id) for the registered app. ($graphApp.AppId)

Write-Host $graphApp.AppId


