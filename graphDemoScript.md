# Demo Script Microsoft Graph
 
## Demo 1: Graph Explorer

1. Go to URL [graph.microsoft.com](https://graph.microsoft.com)
 
 Explain this link is one-stop place for documentation, examples and graph explorer.
 
2. Go to Graph Explorer
3. Show some sample queries from sample account
4. Log in to own tenant to use real data, explain briefly permissions, we will get back to that later in the session.
5. Some sample queries:

  Filter Department..

  Filter on specific attributes..

6. Example queries:

https://graph.microsoft.com/beta/users/?$filter=userType eq 'Guest'

https://graph.microsoft.com/beta/users/?$filter=userType eq 'Guest'&$select=externalUserState,externalUserStateChangeDateTime,UserType

https://graph.microsoft.com/beta/me/joinedTeams

https://graph.microsoft.com/beta/teams/{team-id}/channels

https://graph.microsoft.com/beta/teams/{team-id}/channels/{channel-id}/chatThreads

https://graph.microsoft.com/beta/users/?$filter=userType eq 'Guest'&$select=externalUserState,externalUserStateChangeDateTime,UserType

7. Batch
https://graph.microsoft.com/beta/$batch
{
  "requests": [
    {
      "id": "1",
      "method": "GET",
      "url": "privilegedRoleAssignments/my"
    },
    {
      "id": "2",
      "dependsOn": [
        "1"
      ],
      "method": "GET",
      "url": "/directoryroles"
    }
  ]
}

## Demo 2: Explore Azure AD and Intune Ojects

1. Register App inn Azure AD App Registration
2. Explain Keys and Permissions
3. Walkthrough PowerShell script commands

# Demo 3: Serverless

Show PowerApps, Flows using **Microsoft Graph**, and how this can integrate with for example Microsoft Teams, Power BI and more.
