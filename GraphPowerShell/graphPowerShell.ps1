# --------------------------------------------------------------------------------------------------------
# Name: graphPowerShell.ps1
# Description: The following PowerShell commands use the Microsoft Graph to access User and Intune Objects
# Requirements: 
# Created by: Jan Vidar Elven, Skill AS
# Last modified: 24.10.2018
# --------------------------------------------------------------------------------------------------------

# 1. Specify Tenant & App information from App Registration
$tenantId = "<your-tenantid-here>"
$appId = "<your-applicationid-here>"
$tenantId = "elven.onmicrosoft.com"
$appRedirectUri = "https://login.microsoftonline.com/common/oauth2/nativeclient"

# 2. The following tenant specific URI is read from App Registrations Blade under Endpoints for Oauth
$tokenAuthURI = "https://login.microsoftonline.com/$tenantId/oauth2/token"

# In the following section, either use 3a. or 3b. depending on type of authentication with or without user

# 3. Log in with user - delegated permission
# Build authorize URI
$authorizeAuthURI =  "https://login.microsoftonline.com/$tenantId/oauth2/authorize?" + `
"client_id=$appId" + `
"&response_type=code" + `
"&redirect_uri=$appRedirectUri" + `
"&response_mode=query" + `
"&scope=offline_access%20directory.read.all" + `
"&state=ELEU18"

# Copy Authorize URI, user needs to go to this URI to get code
$authorizeAuthURI | Set-Clipboard

# After authenticating, copy and paste code in this variable, which is the result of user authorization
$appCode = "<paste-in-your-code-from-authorize-uri-here>"

# Build request body for delegated token access via authorization code
$requestBody = "grant_type=authorization_code" + 
"&client_id=$appId" +
"&client_secret=$appSecret" +
"&resource=https://graph.microsoft.com/" +
"&code=$appCode" +
"&scope=directory.read.all" +
"&redirect_uri=$appRedirectUri"

# 4. Thereafter via Oauth Token Endpoint URI submit the values in the request body
$tokenResponse = Invoke-RestMethod -Method Post -Uri $tokenAuthURI -body $requestBody -ContentType "application/x-www-form-urlencoded"

# 5. This response should give us a Bearer Token for later use in Graph API calls
$accessToken = $tokenResponse.access_token
$refreshToken = $tokenResponse.refresh_token

# 6. Some different Graph URI Endpoints for listing User objects
# All Users from a Department
$userlisttURI = "https://graph.microsoft.com/v1.0/users?`$filter=Department eq 'Seinfeld'"
# All Member Users
$userlisttURI = "https://graph.microsoft.com/v1.0/users?`$filter=userType eq 'Member'"
# All Users including Guests
$userlisttURI = "https://graph.microsoft.com/v1.0/users?`$top=5"

# 7. Get the User objects via an authenticated request to Graph API with the help of Bearer Token in authorization header
$graphResponseUsers = Invoke-RestMethod -Method Get -Uri $userlisttURI -Headers @{"Authorization"="Bearer $accessToken"}  

# 8. Loop through PowerShell object returned from Graph query
foreach ($user in $graphResponseUsers.value)
{
    Write-Host $user.userPrincipalName -ForegroundColor Green
    $upn = $user.userPrincipalName
}

# 9. Lets check whether there are more objects to be returned via paging
# This is done by checking if there are a @odata.nextLink with skiptoken
# Looping through until all pages are found
$moregraphresponseusers = $graphresponseusers
$numberOfUsers = $graphResponseUsers.value.Count
if ($graphresponseusers.'@odata.nextLink'){

    $moregraphresponseusers.'@odata.nextLink' = $graphresponseusers.'@odata.nextLink'

    do
        {

            $moregraphresponseusers = Invoke-RestMethod -Method Get -Uri $moregraphresponseusers.'@odata.nextLink' -Headers @{"Authorization"="Bearer $accessToken"}

            $numberOfUsers += $moregraphresponseusers.value.count
            Write-Host $moregraphresponseusers.value.count ".. more objects --> " $numberOfUsers " .. total .." -ForegroundColor Blue

            foreach ($user in $moregraphresponseusers.value)
            {
                Write-Host $user.userPrincipalName -ForegroundColor Green
                $upn = $user.userPrincipalName
            }
        } while ($moregraphresponseusers.'@odata.nextLink')

}

# 10. Lets access Intune data and Managed Apps Graph URI Endpoints
$managedAppsURI = "https://graph.microsoft.com/beta/deviceAppManagement/managedAppRegistrations"

# 11. Get the managed apps objects via an authenticated request to Graph API with the help of Bearer Token in authorization header
$graphResponseManagedApps = Invoke-RestMethod -Method Get -Uri $managedAppsURI -Headers @{"Authorization"="Bearer $accessToken"}  

# 12. Loop through Managed App registrations
foreach ($managedApp in $graphResponseManagedApps.value)
{
    Write-Host "Device Type: " $managedApp.deviceType -ForegroundColor Green
    Write-Host "Device Name: " $managedApp.deviceName -ForegroundColor Green
    Write-Host "Version: " $managedApp.platformVersion -ForegroundColor Green
    Write-Host "Mobile App: " $managedApp.appIdentifier.bundleId -ForegroundColor Green
    $userId = $managedApp.userId
    $userRegisteredURI = "https://graph.microsoft.com/v1.0/users?`$filter=id eq '$userId'&`$select=displayName"
    $graphResponseUserRegistered = Invoke-RestMethod -Method Get -Uri $userRegisteredURI -Headers @{"Authorization"="Bearer $accessToken"}  
    Write-Host "User Registered: " $graphResponseUserRegistered.value.displayName -ForegroundColor Yellow
}

#region refresh token

# If the Access Token is Expired (1 hour), use the Refresh Token to get a new Access Token
$refreshBody = "grant_type=refresh_token" +
"&redirect_uri=$appRedirectUri" +
"&client_id=$appId" + 
"&client_secret=$appSecret" + 
"&refresh_token=$refreshToken" +
"&resource=https://graph.microsoft.com/"

$tokenResponse = Invoke-RestMethod -Method Post -Uri $tokenAuthURI -body $refreshBody -ContentType "application/x-www-form-urlencoded"
$accessToken = $tokenResponse.access_token
$refreshToken = $tokenResponse.refresh_token

#endregion
