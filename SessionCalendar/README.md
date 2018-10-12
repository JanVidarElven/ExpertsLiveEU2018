# Retrieve Sessions and Add Appointments via Microsoft Graph

This solution will read sessions from the https://expertslive.eu website and use Microsoft Graph to add sessions as events in your calendar.

The solution is based on original idea by Oskar Landman, OWL IT, http://www.owl-it.nl/uncategorized/powershell-retrieve-session-information-from-internet-and-create-ics-file/, a PowerShell script that creates a ics file for import to calendar.

This solution builds on that, but after retrieving the sessions they are imported using Microsoft Graph to the logged on users calendar.

## Solution Contents

The solution consists of:
* CalendarGraphApi_Setup.ps1. This script contains the cmdlets for creating a native app registration in Azure AD, with the permissions accessing calendar events via Microsoft Graph. (https://developer.microsoft.com/en-us/graph/docs/api-reference/v1.0/resources/event). This registration will provide you with a client id you will need in the next script.
* RetrieveSessionsAndAddAppointmentsViaGraph.ps1. This script reads the sessions from https://expertslive.eu and adds or updates calendar events. The script requires to parameters, client id (for the app registration) and upn for your user in which the calendar should be updated.

Note! For creating an app registration you need to have a global administrator or application administrator role for your tenant, or you can have your tenant administrator run the first script for you.

## Requirements

This solution requires:

* You need to have an Office 365 Subscription with an Exchange Online mailbox.
* You need to have installed either the AzureAD or AzureADPreview PowerShell Module in your computer.

## Usage

1. Run the CalendarGraphApi_Setup.ps1 script as a global admin or application administrator to create the app registration.
1. Run the script and specify parameters like .\RetrieveSessionsAndAddAppointmentsViaGraph.ps1 -User "yourname@domain.com" -ClientId "your-clientid-or-well-known-clientid"

## Known Issues
The script will only add the sessions if they are not already there, so you can run the script multiple times in case there are any description, rooms, or other updates. Some session titles contains some punctuation or dash characters making the JSON updates to calendar and matching with existing events difficult. So these sessions are added as new each time even though they are there from before. I'll try to fix it, and commit to this solution, but you are welcome to contribute to this GitHub project ;)

## Blog Details
Look to my upcoming blog post (coming soon!) [gotoguy.blog](http://gotoguy.blog) for more explanation of code and usage scenarios.
