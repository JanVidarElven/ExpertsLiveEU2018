# --------------------------------------------------------------------------------------------------------
# Name: graphPowerShell.ps1
# Description: The following PowerShell commands use the Microsoft Graph to access User and Intune Objects
# Requirements: 
# Created by: Jan Vidar Elven, Skill AS
# Last modified: 25.10.2018
# --------------------------------------------------------------------------------------------------------

# User for Delegated Permission
$User = "jan.vidar@elven.no"
# Azure AD App Registration
$ClientId = "1950a258-227b-4e31-a9cf-717495945fc2"
# Well known Client Id for Intune PowerShell:
#$clientId = "d1ddf0e4-d672-4dae-b554-9d5bdfd93547"
# Well known Client Id for Azure PowerShell:
#$ClientId = "1950a258-227b-4e31-a9cf-717495945fc2"


# Get-AuthToken function, from Intune Graph API samples
function Get-AuthToken {
    <#
    .SYNOPSIS
    This function is used to authenticate with the Graph API REST interface
    .DESCRIPTION
    The function authenticate with the Graph API Interface with the tenant name
    .EXAMPLE
    Get-AuthToken
    Authenticates you with the Graph API interface
    .NOTES
    NAME: Get-AuthToken
    #>
    [cmdletbinding()]
    param
    (
        [Parameter(Mandatory=$true)]
        $User,
        [Parameter(Mandatory = $true)]
        $ClientId
    )
    $userUpn = New-Object "System.Net.Mail.MailAddress" -ArgumentList $User
    $tenant = $userUpn.Host
    
    Write-Host "Checking for AzureAD module..."
        $AadModule = Get-Module -Name "AzureAD" -ListAvailable
        if ($AadModule -eq $null) {
            Write-Host "AzureAD PowerShell module not found, looking for AzureADPreview"
            $AadModule = Get-Module -Name "AzureADPreview" -ListAvailable
        }
        if ($AadModule -eq $null) {
            write-host
            write-host "AzureAD Powershell module not installed..." -f Red
            write-host "Install by running 'Install-Module AzureAD' or 'Install-Module AzureADPreview' from an elevated PowerShell prompt" -f Yellow
            write-host "Script can't continue..." -f Red
            write-host
            exit
        }

    # Getting path to ActiveDirectory Assemblies
    # If the module count is greater than 1 find the latest version
        if($AadModule.count -gt 1){
            $Latest_Version = ($AadModule | Select-Object version | Sort-Object)[-1]
            $aadModule = $AadModule | ? { $_.version -eq $Latest_Version.version }
                # Checking if there are multiple versions of the same module found
                if($AadModule.count -gt 1){
                $aadModule = $AadModule | Select-Object -Unique
                }
            $adal = Join-Path $AadModule.ModuleBase "Microsoft.IdentityModel.Clients.ActiveDirectory.dll"
            $adalforms = Join-Path $AadModule.ModuleBase "Microsoft.IdentityModel.Clients.ActiveDirectory.Platform.dll"
        }
        else {
            $adal = Join-Path $AadModule.ModuleBase "Microsoft.IdentityModel.Clients.ActiveDirectory.dll"
            $adalforms = Join-Path $AadModule.ModuleBase "Microsoft.IdentityModel.Clients.ActiveDirectory.Platform.dll"
        }
    
    [System.Reflection.Assembly]::LoadFrom($adal) | Out-Null
    [System.Reflection.Assembly]::LoadFrom($adalforms) | Out-Null
    
    $redirectUri = "urn:ietf:wg:oauth:2.0:oob"
    $resourceAppIdURI = "https://graph.microsoft.com"
    $authority = "https://login.microsoftonline.com/$Tenant"
  
    try {
        $authContext = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext" -ArgumentList $authority
        # https://msdn.microsoft.com/en-us/library/azure/microsoft.identitymodel.clients.activedirectory.promptbehavior.aspx
        # Change the prompt behaviour to force credentials each time: Auto, Always, Never, RefreshSession
        $platformParameters = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.PlatformParameters" -ArgumentList "Auto"
        $userId = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.UserIdentifier" -ArgumentList ($User, "OptionalDisplayableId")
        $authResult = $authContext.AcquireTokenAsync($resourceAppIdURI,$clientId,$redirectUri,$platformParameters,$userId).Result
            # If the accesstoken is valid then create the authentication header
            if($authResult.AccessToken){
            # Creating header for Authorization token
            $authHeader = @{
                'Content-Type'='application/json'
                'Authorization'="Bearer " + $authResult.AccessToken
                'ExpiresOn'=$authResult.ExpiresOn
                }
            return $authHeader
            }
            else {
            Write-Host
            Write-Host "Authorization Access Token is null, please re-run authentication..." -ForegroundColor Red
            Write-Host
            break
            }
        }
        catch {
        write-host $_.Exception.Message -f Red
        write-host $_.Exception.ItemName -f Red
        write-host
        break
        }
    }
    
####################################################

#region Authentication
write-host
# Checking if authToken exists before running authentication
if($global:authToken){
    # Setting DateTime to Universal time to work in all timezones
    $DateTime = (Get-Date).ToUniversalTime()
    # If the authToken exists checking when it expires
    $TokenExpires = ($authToken.ExpiresOn.datetime - $DateTime).Minutes
        if($TokenExpires -le 0){
        write-host "Authentication Token expired" $TokenExpires "minutes ago" -ForegroundColor Yellow
        write-host
            # Defining User Principal Name if not present
            if($User -eq $null -or $User -eq ""){
            $User = Read-Host -Prompt "Please specify your user principal name for Azure Authentication"
            Write-Host
            }
        $global:authToken = Get-AuthToken -User $User
        }
}
# Authentication doesn't exist, calling Get-AuthToken function
else {
    if($User -eq $null -or $User -eq ""){
    $User = Read-Host -Prompt "Please specify your user principal name for Azure Authentication"
    Write-Host
    }
# Getting the authorization token
$global:authToken = Get-AuthToken -User $User -ClientId $ClientId
}

#endregion

# The commands under uses the $authToken Authorization in the Header

# 1. Some different Graph URI Endpoints for listing User objects
# All Users from a Department
$userlisttURI = "https://graph.microsoft.com/v1.0/users?`$filter=Department eq 'Seinfeld'"
# All Member Users
$userlisttURI = "https://graph.microsoft.com/v1.0/users?`$filter=userType eq 'Member'"
# All Users including Guests
$userlisttURI = "https://graph.microsoft.com/v1.0/users?`$top=5"

# 2. Get the User objects via an authenticated request to Graph API with the help of Bearer Token in authorization header
$graphResponseUsers = Invoke-RestMethod -Method Get -Uri $userlisttURI -Headers @{"Authorization"=$authToken.Authorization}  

# 3. Loop through PowerShell object returned from Graph query
foreach ($user in $graphResponseUsers.value)
{
    Write-Host $user.userPrincipalName -ForegroundColor Green
    $upn = $user.userPrincipalName
}

# 4. Lets check whether there are more objects to be returned via paging
# This is done by checking if there are a @odata.nextLink with skiptoken
# Looping through until all pages are found
$moregraphresponseusers = $graphresponseusers
$numberOfUsers = $graphResponseUsers.value.Count
if ($graphresponseusers.'@odata.nextLink'){

    $moregraphresponseusers.'@odata.nextLink' = $graphresponseusers.'@odata.nextLink'

    do
        {

            $moregraphresponseusers = Invoke-RestMethod -Method Get -Uri $moregraphresponseusers.'@odata.nextLink' -Headers @{"Authorization"=$authToken.Authorization}

            $numberOfUsers += $moregraphresponseusers.value.count
            Write-Host $moregraphresponseusers.value.count ".. more objects --> " $numberOfUsers " .. total .." -ForegroundColor Blue

            foreach ($user in $moregraphresponseusers.value)
            {
                Write-Host $user.userPrincipalName -ForegroundColor Green
                $upn = $user.userPrincipalName
            }
        } while ($moregraphresponseusers.'@odata.nextLink')

}

# 5. Lets access Intune data and Managed Apps Graph URI Endpoints
$managedAppsURI = "https://graph.microsoft.com/beta/deviceAppManagement/managedAppRegistrations"

# 6. Get the managed apps objects via an authenticated request to Graph API with the help of Bearer Token in authorization header
$graphResponseManagedApps = Invoke-RestMethod -Method Get -Uri $managedAppsURI -Headers @{"Authorization"=$authToken.Authorization}  

# 7. Loop through Managed App registrations
foreach ($managedApp in $graphResponseManagedApps.value)
{
    Write-Host "Device Type: " $managedApp.deviceType -ForegroundColor Green
    Write-Host "Device Name: " $managedApp.deviceName -ForegroundColor Green
    Write-Host "Version: " $managedApp.platformVersion -ForegroundColor Green
    Write-Host "Mobile App: " $managedApp.appIdentifier.bundleId -ForegroundColor Green
    $userId = $managedApp.userId
    $userRegisteredURI = "https://graph.microsoft.com/v1.0/users?`$filter=id eq '$userId'&`$select=displayName"
    $graphResponseUserRegistered = Invoke-RestMethod -Method Get -Uri $userRegisteredURI -Headers @{"Authorization"=$authToken.Authorization}  
    Write-Host "User Registered: " $graphResponseUserRegistered.value.displayName -ForegroundColor Yellow
}
