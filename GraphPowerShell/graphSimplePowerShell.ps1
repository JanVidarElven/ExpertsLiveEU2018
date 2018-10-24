# The following PowerShell commands use the Microsoft Graph to access Resources
# Requires that you previously have aquired an access token
# Jan Vidar Elven, June 2018

# Paste your access token here
$accessToken = "..."

# Specify a Uri for querying for resources
$resourceUri = "https://graph.microsoft.com/v1.0/users?`$filter=Department eq 'Seinfeld'"

# Executing Graph query
Invoke-RestMethod -Method Get -Uri $resourceUri -Headers @{"Authorization"="Bearer $accessToken"}

# Saving response value
$response = Invoke-RestMethod -Method Get -Uri $resourceUri -Headers @{"Authorization"="Bearer $accessToken"}

# Getting items from response value
$response.value
