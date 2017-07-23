# teams-module
PowerShell module for Microsoft Teams using the rest API that are not documented from api.teams.skype.com.

Requirements;
AzureAD PowerShell module
PowerShell 4.0 or higher
A Microsoft Teams license

Usage:
import-module teams_v2.psm1

connect-teamsservice -user admin@contoso.com -tenant contoso.onmicrosoft.com
get-Team
add-Team 
remove-Team
add-TeamMember
remove-TeamMember
convert-TeamMemberToOwner
convert-TeamOwnerToMember
