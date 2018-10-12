# --------------------------------------------------------------------------------------------------------
# Name: CalendarGraphApi_Setup.ps1
# Description: Create Azure AD App Registration for using Calendar Graph Api (Microsoft Graph)
# Created By: Jan Vidar Elven, MVP, Skill AS
# Last changed: 09.10.2018
# --------------------------------------------------------------------------------------------------------

# Connect to Azure AD
$tenantId = "yourtenant.onmicrosoft.com"
Connect-AzureAD -TenantId $tenantId

# Prepare Microsoft Graph API permissions with Azure AD Service Principal for Microsoft Graph
$graphPrincipal = Get-AzureADServicePrincipal -All $true | Where-Object {$_.DisplayName -eq "Microsoft Graph"}
$graphAccess = New-Object -TypeName "Microsoft.Open.AzureAD.Model.RequiredResourceAccess"
$graphAccess.ResourceAppId = $graphPrincipal.AppId

# List available Application Permissions
$graphPrincipal.AppRoles

# List available Delegated Permissions
$graphPrincipal.Oauth2Permissions 

# Get all Calendar delegated permissions
$graphDelegatedPermissions = $graphPrincipal.Oauth2Permissions | Where-Object {$_.Value -like 'Calendar*'}

# Loop through and assign all permissions
ForEach ($graphPermission in $graphDelegatedPermissions ) {
    Write-Host -ForegroundColor Green ("Adding Delegated Permission for " + $graphPermission.Value)
    $graphAddPermission = New-Object -TypeName "microsoft.open.azuread.model.resourceAccess" -ArgumentList $graphPermission.Id, "Scope"
    $graphAccess.ResourceAccess += $graphAddPermission
}

# Create Native App Registration for Calendar Graph Api with specified permissions
$graphApp = New-AzureADApplication -DisplayName "Calendar Graph Api" -PublicClient $true -ReplyUrls "urn:ietf:wg:oauth:2.0:oob"  -RequiredResourceAccess @($graphAccess)

# Add Service Principal for App. Note use of tag: https://docs.microsoft.com/en-us/powershell/module/azuread/new-azureadserviceprincipal?view=azureadps-2.0
$graphAppSpn = New-AzureADServicePrincipal -AppId $graphApp.AppId -Tags @("WindowsAzureActiveDirectoryIntegratedApp")

# TODO:
# 1. Remember to Grant Admin Consent in Azure AD Portal for delegated permissions that requires it (most do)
# 2. Note ned Client Id (App Id) for the registered app. ($graphApp.AppId)

Write-Host $graphApp.AppId


