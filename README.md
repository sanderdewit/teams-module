# teams-module
PowerShell module for Microsoft Teams


AzureAD PowerShell module is required.
PowerShell 4.0 or higher is required.

connect-teamsservice -user admin@contoso.com -tenant contoso.onmicrosoft.com

current functions;
add-team, remove-team, get-team, get-teammembers, add-teammember, remove-teammember, add-teamowner, remove-teamowner

add-team currently only adds an o365 unified group. Unfortunately the API does not support this yet.
