# --------------------------------------------------------------------------------------------------------
# Name: RetrieveSessionsAndAddAppointmentsViaGraph.ps1
# Description: Create Azure AD App Registration for using Calendar Graph Api (Microsoft Graph)
# Created By: Jan Vidar Elven, MVP, Skill AS
# Built On: Original Idea from Oskar Landman, OWL IT, http://www.owl-it.nl/uncategorized/powershell-retrieve-session-information-from-internet-and-create-ics-file/
# Last changed: 12.10.2018
# --------------------------------------------------------------------------------------------------------

[Cmdletbinding()]
param(
    [ValidateNotNullOrEmpty()]
    [Parameter(Mandatory = $true, HelpMessage = 'The User UPN of which Calendar should be updated')]
    [string] $User,
    [Parameter(Mandatory=$true, HelpMessage = 'The Client Id of the Native App registered in Azure AD for Microsoft Graph Access')]
    [string]$ClientId
)

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
Function Add-CalendarEvent(){
    <#
    .SYNOPSIS
    This function is used to create a Calendar Event using the Graph API REST interface
    .DESCRIPTION
    The function connects to the Graph API Interface and creates a calendar event
    .EXAMPLE
    Add-CalendarEvent -JSON $JSON
    Creates Calendar Event using Graph API
    .NOTES
    NAME:Add-CalendarEvent
    #>
    [cmdletbinding()]
    param
    (
        $JSON
    )
    $graphApiVersion = "v1.0"
    $App_resource = "me/events"
        try {
            if(!$JSON){
            write-host "No JSON was passed to the function, provide a JSON variable" -f Red
            break
            }
            else {
            Test-JSON -JSON $JSON
            $uri = "https://graph.microsoft.com/$graphApiVersion/$($App_resource)"
            Invoke-RestMethod -Uri $uri -Method Post -ContentType "application/json" -Body $JSON -Headers $authToken
            }
        }
        catch {
            $ex = $_.Exception
            $errorResponse = $ex.Response.GetResponseStream()
            $reader = New-Object System.IO.StreamReader($errorResponse)
            $reader.BaseStream.Position = 0
            $reader.DiscardBufferedData()
            $responseBody = $reader.ReadToEnd();
            Write-Host "Response content:`n$responseBody" -f Red
            Write-Error "Request to $Uri failed with HTTP Status $($ex.Response.StatusCode) $($ex.Response.StatusDescription)"
            write-host
            break
        }
    }

####################################################
Function Update-CalendarEvent(){
    <#
    .SYNOPSIS
    This function is used to update an existing Calendar Event using the Graph API REST interface
    .DESCRIPTION
    The function connects to the Graph API Interface and updates a calendar event
    .EXAMPLE
    Update-CalendarEvent -calenderEventId $id -JSON $JSON
    Updates Calendar Event using Graph API
    .NOTES
    NAME:Update-CalendarEvent
    #>
    [cmdletbinding()]
    param
    (
        $JSON,
        $calendarEventId
    )
    $graphApiVersion = "v1.0"
    $App_resource = "me/events"
        try {
            if(!$JSON){
            write-host "No JSON was passed to the function, provide a JSON variable" -f Red
            break
            }
            elseif (!$calendarEventId) {
            write-host "No Event Id was passed to the function, provide a event id variable" -f Red
            }
            else {
            Test-JSON -JSON $JSON
            $uri = "https://graph.microsoft.com/$graphApiVersion/$($App_resource)/$calendarEventId"
            Invoke-RestMethod -Uri $uri -Method Patch -ContentType "application/json" -Body $JSON -Headers $authToken
            }
        }
        catch {
            $ex = $_.Exception
            $errorResponse = $ex.Response.GetResponseStream()
            $reader = New-Object System.IO.StreamReader($errorResponse)
            $reader.BaseStream.Position = 0
            $reader.DiscardBufferedData()
            $responseBody = $reader.ReadToEnd();
            Write-Host "Response content:`n$responseBody" -f Red
            Write-Error "Request to $Uri failed with HTTP Status $($ex.Response.StatusCode) $($ex.Response.StatusDescription)"
            write-host
            break
        }
    }

####################################################
Function Find-CalendarEvents(){
    <#
    .SYNOPSIS
    This function is used to find existing Calendar Events using the Graph API REST interface
    .DESCRIPTION
    The function connects to the Graph API Interface and finds existing calendar events
    .EXAMPLE
    Find-CalendarEvents -StartDateTime "YYYY-MM-DDTHH:MM:SS.000Z" -EndDateTime "YYYY-MM-DDTHH:MM:SS.000Z"
    Creates Calendar Event using Graph API
    .NOTES
    NAME:Find-CalendarEvents
    #>
    [cmdletbinding()]
    param
    (
        [Parameter(Mandatory=$true)]
        $StartDateTime,
        [Parameter(Mandatory=$true)]
        $EndDateTime,
        [Parameter(Mandatory=$false)]
        $pageSize=10
    )
    $graphApiVersion = "v1.0"
    $App_resource = "me/calendarview"
        try {
            if(!$StartDateTime){
            write-host "No StartDateTime was passed to the function, provide a StartDateTime variable" -f Red
            break
            }
            else {
            $uri = "https://graph.microsoft.com/$graphApiVersion/$($App_resource)?`$top=$pageSize&startdatetime=$StartDateTime&enddatetime=$EndDateTime"
            Invoke-RestMethod -Uri $uri -Method Get -ContentType "application/json" -Headers $authToken
            }
        }
        catch {
            $ex = $_.Exception
            $errorResponse = $ex.Response.GetResponseStream()
            $reader = New-Object System.IO.StreamReader($errorResponse)
            $reader.BaseStream.Position = 0
            $reader.DiscardBufferedData()
            $responseBody = $reader.ReadToEnd();
            Write-Host "Response content:`n$responseBody" -f Red
            Write-Error "Request to $Uri failed with HTTP Status $($ex.Response.StatusCode) $($ex.Response.StatusDescription)"
            write-host
            break
        }
    }    

####################################################

Function Test-JSON(){
    <#
    .SYNOPSIS
    This function is used to test if the JSON passed to a REST Post request is valid
    .DESCRIPTION
    The function tests if the JSON passed to the REST Post is valid
    .EXAMPLE
    Test-JSON -JSON $JSON
    Test if the JSON is valid before calling the Graph REST interface
    .NOTES
    NAME: Test-JSON
    #>
    param (
    $JSON
    )
        try {
        $TestJSON = ConvertFrom-Json $JSON -ErrorAction Stop
        $validJson = $true
        }
        catch {
        $validJson = $false
        $_.Exception
        }
        if (!$validJson){
        Write-Host "Provided JSON isn't in valid JSON format" -f Red
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

#region Get Sessions
$global:ie=new-object -com "internetexplorer.application"
$ie.visible=$true
$ie.navigate("https://www.expertslive.eu/sessions--call-for-speakers.html")

#wait for IE to be ready before pulling the sessions
while ($ie.readystate -lt 4){start-sleep -milliseconds 200}
write-host "Retrieving sessions from site" 
$Sessions = $ie.Document.getElementsByTagName("li") |Where-Object {$_.className -eq "sz-session sz-session--full"}

$Amount = $Sessions.Count
Write-host "Number of sessions: $Amount"

# Get existing events in calendar for timeframe of conference
$sessionEvents = Find-CalendarEvents -StartDateTime "2018-10-24T08:00:00.000Z" -EndDateTime "2018-10-26T18:00:00.000Z" -pageSize 200

foreach ($Session in $Sessions)
{
    # Remove variants of qoutes and other special characters from description/titles to support JSON
    $title = ($Session.getElementsByClassName("sz-session__title")[0].Innertext -replace """|`„|`“|`”","")
    $Description = ($Session.getElementsByClassName("sz-session__description")[0].Innertext -replace """|`„|`“|`”","")
    
    # Remove hard enter and trailing spaces from Speaker
    $Speaker = ($Session.getElementsByClassName("sz-session__speakers")[0].Innertext -replace "`t|`n|`r","").Trim()
    $room = $Session.getElementsByClassName("sz-session__room")[0].Innertext
    $Time = $Session.getElementsByClassName("sz-session__time")[0].Innertext
    # Remove hard enter and trailing spaces from Session Tags
    $tags = $session.getElementsByClassName("sz-session__tags")[0].Innertext -replace "`t|`n|`r",""
    $tag = $tags.Split(" ")
    $level = $tag[0]

    $tags = $tags.Replace($level,"").Trim()
    $tags = $tags.Replace(" ","/")

    $location = "Room($room) - Level($level) - Speaker($speaker) - Tags($Tags)"

    $day = $time.split(" ")
    $weekday = $day[0]

    switch ($weekday)
        {
            Wed {$day='10/24/2018'}
            Thu {$day='10/25/2018'}
            Fri {$day='10/26/2018'}
        }
    $parts = $time.Split("-")
    $startString = ($parts[0]).Replace($weekday,"")
    [datetime]$start = $day + "," + $startString
    $Start = $start.ToUniversalTime()
    $endString = ($parts[1]).Replace($weekday,"")
    [datetime]$end = $day + "," + $endString
    $end = $end.ToUniversalTime()

    $guid = New-Guid
    # Convert Date Time to ISO 8601 Format
    $StartTime = Get-date $start -Format O
    $EndTime = Get-date $end -Format O

    Write-host "==========================================================="
    Write-host "($Speaker)$title"

    # Build a JSON object for adding/updating calendar event
    $JSON_Calendar = @"

    {
    "subject": "$title",
    "body": {
      "contentType": "HTML",
      "content": "$Description"
    },
    "start": {
      "dateTime": "$starttime",
      "timeZone": "W. Europe Standard Time"
    },
    "end": {
      "dateTime": "$endtime",
      "timeZone": "W. Europe Standard Time"
    },
    "location": {
        "displayName": "$location"
      },
      "reminderMinutesBeforeStart": 15,
      "isReminderOn": true,
      "iCalUId": "$guid"                
  }

"@

    # Check if session title already exists
    If ($filterSession = $sessionEvents.value | Where-Object {$_.subject -match $title}) {
        Write-Host "Session $title already exists in calendar, updating..."
        $calendarEventId = $filterSession.id
        # Update existing session
        Update-CalendarEvent -calendarEventId $calendarEventId -JSON $JSON_Calendar                
    }
    Else {
        Write-Host "Adding session $title to calendar..."
        # Adding session to calendar
        Add-CalendarEvent -JSON $JSON_Calendar
    }
}

$global:ie.Quit()

#endregion

####################################################
