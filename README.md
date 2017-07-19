# teams-module
PowerShell module for Microsoft Teams


AzureAD PowerShell module is required.
PowerShell 4.0 or higher is required.

connect-teamsservice -user admin@contoso.com -tenant contoso.onmicrosoft.com

using teams_v2.psm1 uses the undocumented Teams API.

current functions;
add-team.

the team.psm1 uses the groups graph api but is unable to provision a real team, while team_v2.psm1 is able to provision a correct team.
