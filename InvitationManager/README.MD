# Invite Azure AD B2B Users via Microsoft Graph and PowerApps

This solution contains examples of POST requests for inviting B2B users with Microsoft Graph, as well as a swagger definition for a Custom Connector to be used in PowerApps and Flows, and an exported package containing the PowerApp and the Flow used to send an invite to a B2B user and write a chat message in Teams.

## Invitation JSON Body samples

To create an invite using Microsoft Graph, for example by using Graph Explorer, do a POST request to either https://graph.microsoft.com/1.0/invitations (or https://graph.microsoft.com/beta/invitations), and specify a JSON body request as the examples provided:

* AzureADB2BInvitation_Minimum.json
* AzureADB2BInvitation_Simple.json
* AzureADB2BInvitation_Full.json

## Swagger Definition

The provided file: "ExpertsLive-Custom-Connector.swagger", contains the needed definitions for the operations used for inviting Azure AD B2B User and for creating a chat thread in Teams. You can either import this as a new Connector in your organization, or update an existing connector. 

Look to this blog post for more explanation and details about importing and exporting OpenAPI swagger definition:[https://gotoguy.blog/2018/09/19/exporting-and-importing-powerapps-and-flows-package-that-use-a-custom-connector/](https://gotoguy.blog/2018/09/19/exporting-and-importing-powerapps-and-flows-package-that-use-a-custom-connector/) 

## PowerApps Package

The provided zip file: "InviteAzureADB2BUserPowerApp_***", contains the PowerApps App and the related Flow. It will import as creating a new PowerApp and Flow, and specify existing Custom Connector and Connections. 

Look to this blog post for more explanation and details about importing and exporting PowerApps packaged:[https://gotoguy.blog/2018/09/19/exporting-and-importing-powerapps-and-flows-package-that-use-a-custom-connector/](https://gotoguy.blog/2018/09/19/exporting-and-importing-powerapps-and-flows-package-that-use-a-custom-connector/) 

## Blog Details

Look to my blog [gotoguy.blog](http://gotoguy.blog) for other usage scenarios with Microsoft Graph, PowerApps, and more.
