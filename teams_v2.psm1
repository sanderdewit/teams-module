function connect-TeamsService {
  <#
      .SYNOPSIS
      This function is used to authenticate with the Teams Graph API
 
      .DESCRIPTION
      The function authenticate with the Graph API Interface with the tenant name

      .EXAMPLE
      connect-TeamsService -user user@contoso.com -tenant contoso.onmicrosoft.com
 
      .NOTES
 
      NAME: connect-TeamsService
 
  #>
 
  [cmdletbinding()]
  param
  (
    [Parameter(Mandatory=$true)]$User,
    [Parameter(Mandatory=$true)]$tenant,
    [Parameter(Mandatory=$false)][switch]$silent
  )
  Write-Verbose "Checking for AzureAD module..."
  $AadModule = Get-Module -Name "AzureAD" -ListAvailable
if ($AadModule -eq $null) {
    write-warning "AzureAD Powershell module not installed..."
    Write-Warning "Install by running 'Install-Module AzureAD' from an elevated PowerShell prompt"
    Write-Warning "Script can't continue..."
    throw ('no AzureAD module found, the AzureAD module is required. please run install-module AzureAD')
  }
 
  # Getting path to ActiveDirectory Assemblies
  # If the module count is greater than 1 find the latest version
  if($AadModule.count -gt 1){
    $Latest_Version = ($AadModule | Select-Object version | Sort-Object)[-1]
    $aadModule = $AadModule | Where-Object { $_.version -eq $Latest_Version.version }
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
  $clientId = 'd3590ed6-52b3-4102-aeff-aad2292ab01c'
  $redirectUri = "urn:ietf:wg:oauth:2.0:oob"
  $resourceAppIdURI = "https://api.spaces.skype.com" #the API url.
  $authority = "https://login.windows.net/$Tenant"
  try {
    $authContext = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext" -ArgumentList $authority
    # https://msdn.microsoft.com/en-us/library/azure/microsoft.identitymodel.clients.activedirectory.promptbehavior.aspx
    # Change the prompt behaviour to force credentials each time: Auto, Always, Never, RefreshSession
    $platformParameters = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.PlatformParameters" -ArgumentList "Always"
    $userId = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.UserIdentifier" -ArgumentList ($User, "OptionalDisplayableId")
    #New-Object 'Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationParameters'
    if ($silent){$authResult = $authContext.AcquireTokenSilentAsync($resourceAppIdURI,$clientId,$userid,$platformParameters).Result}
    else {
    $authResult = $authContext.AcquireTokenAsync($resourceAppIdURI,$clientId,$redirectUri,$platformParameters,$userId).Result}
    # If the accesstoken is valid then create the authentication header
    if($authResult.AccessToken){
      # Creating header for Authorization token
      $TeamsauthHeader = @{
        'Content-Type'='application/json'
        'Authorization'="Bearer " + $authResult.AccessToken
        'ExpiresOn'=$authResult.ExpiresOn
      }
      #retrieving Skype token required for Teams.
      $AuthSkypeResult = Invoke-RestMethod -Uri 'https://api.teams.skype.com/beta/auth/skypetoken' -Method Post -Headers $TeamsauthHeader
      if ($($AuthSkypeResult.tokens.skypetoken) -eq $null){
      Write-Error "unable to retrieve Skype Token. $($authResult|convertfrom-json)"
      throw ('No valid Skype Token retrieved')
      }else {
        $TeamsauthHeader += @{
        'X-Skypetoken' = $($AuthSkypeResult.tokens.skypetoken)
      }
        $SkypeHeader = @{
        Authentication = "skypetoken=$($TeamsauthHeader.'X-Skypetoken')"
        'content-type' = 'application/json'
        }
      }
      $global:TeamsAuthToken = $TeamsauthHeader
      $global:SkypeToken = $SkypeHeader
    }
    else {
      Write-Warning "Authorization Access Token is null, please re-run authentication..."
      break
    }
  }
 
  catch {
    write-error $_.Exception.Message
    write-error $_.Exception.ItemName
    break
  }
}
 
 function new-Team {
  <#
  .SYNOPSIS
  Invoking a rest request to the Microsoft Teams graph api to add a team
  .DESCRIPTION
  Describe the function in more detail
  .EXAMPLE
  add-team -displayname 'Team Display' -description 'Team Description' -smtpaddress 'team@contoso.com' -alias 'team' -type 'public'
  #>
  [CmdletBinding()]
  #[CmdletBinding(SupportsShouldProcess=$True,ConfirmImpact='Low')]
  param
  (
        [Parameter(Mandatory=$true)]$description,
        [Parameter(Mandatory=$true)]$displayname,
        [Parameter(Mandatory=$true)]$smtpaddress,
        [Parameter(Mandatory=$true)]$alias,
        [Parameter(Mandatory=$true)]$Type
  )
  begin {
  write-verbose "checking for teams token"
    if (!($TeamsAuthToken))
    {throw 'please run connect-TeamsService first'}
  }
  process {
    write-verbose "start to invoke rest request"
    $AcccessType = switch ($Type){
    Private {1}
    Public {3}
    }
    $uri = 'https://api.teams.skype.com/emea/beta/teams/create'
    $postparams = @{
    'alias' = $alias
    'description' = $description
    'displayName' = $displayname
    'smtpAddress'=  $smtpaddress
    'AccessType' = $AcccessType
    }
    $result = Invoke-RestMethod -Uri $uri -Headers $TeamsAuthToken -Method post -Body $($postparams|convertto-json)
    Write-Verbose "added team $displayName"
    Write-Verbose "$($result.value)"
    "team created with id $($result.value.SiteInfo.groupid)"
  }
}

 function add-TeamMember {
  <#
  .SYNOPSIS
  Invoking a rest request to the Microsoft graph api to add a teammember
  .DESCRIPTION
  This invoke the restapi to add a teammember to a team.
  .EXAMPLE
  add-teammember -team 'teamtest' -member 'user@contoso.com'
  #>
  [CmdletBinding()]
  #[CmdletBinding(SupportsShouldProcess=$True,ConfirmImpact='Low')]
  param
  (
        [Parameter(Mandatory=$true)]$Team,
        [Parameter(Mandatory=$true)]$Members
  )
  begin {
  write-verbose "checking for teams token"
    if (!($TeamsAuthToken))
    {throw 'please run connect-TeamsService first'}
  }
  process {
    write-verbose 'start to invoke rest request'
    #check if ID is given as parameter.
    Write-Verbose 'validating parameter'
    if ($($team.length) -eq '36' -and $Team -match “[0123456789abcdef-]{$($Team.length)}”){
    Write-Verbose "ip detected, skipping lookup"
    $TeamId = $Team
    } #check if ID is given as parameter.
    else {  #finding team based on wildcard search
    $Teams = Invoke-RestMethod -Uri 'https://api.teams.skype.com/emea/beta/teams/usergroups?teamType=null' -Method get -Headers $TeamsAuthToken
    Write-Verbose "$teams"
    $TeamId = $Teams|Where-Object {$_.displayName -like "*$Team*"} #filtering does not yet work using $filter, so retrieving all teams and filters on the output
    Write-Verbose "ID: $($TeamId.groupId)"
    if ($($TeamId.groupId) -eq $null){write-error 'team not found'
    throw ("team $Team couldn't be found")}
    }
    Write-Verbose "finding team member"
    #finding team member
    $Members = foreach ($Member in $Members){
    if ($($Member.length) -eq '36' -and $Member -match “[0123456789abcdef-]{$($Member.length)}”){
    $MemberResult = $Member}
    else {
    $MemberResult = (Invoke-RestMethod -Uri ' https://api.teams.skype.com/emea/beta/users/search?includeDLs=true&includeBots=false&enableGuest=false&skypeTeamsInfo=true' -Method Post -Headers $TeamsAuthToken -Body $Member).value.objectId #assuming correct UPN to be send.
    if ($MemberResult -eq $null){write-error 'member not found'
    throw ("UserPrincipalName $Member couldn't be found")}
    }
    [pscustomobject]@{ #create the members object
    mri = "8:orgid:$MemberResult"
    role = '0' #role 0 is used for members, 1 is used for owners
    }}
    Write-Verbose "members: $Members"
    $teaminfo = (Invoke-RestMethod -Uri 'https://emea-client-ss.msg.skype.com/v1/users/ME/conversations?view=msnp24Equivalent&pageSize=300&startTime=1&targetType=Passport|Skype|Lync|Thread|NotificationStream' -Method Get -Headers $skypetoken).conversations|Where-Object {$_.threadProperties.isdeleted -ne 'true' -and $_.threadProperties.threadType -eq 'space'}  
    $teamresults = foreach ($team in $teams){
    try {
    $teamres = ([uri]($teaminfo|Where-Object {$_.threadProperties.spaceThreadTopic -eq $team.displayName}).targetLink).Segments[3]
    [pscustomobject]@{
    groupId = $team.groupId
    displayName = $team.displayName
    targetLink = $teamres}
    }catch {}
    }
    $teamurl = ($teamresults|Where-Object {$_.groupid -eq $($TeamId.groupId)}).targetlink
    Write-Verbose "$teamurl"
    $uri = "https://api.teams.skype.com/emea/beta/teams/$($teamurl)/bulkUpdateRoledMembers?allowBotsInChannel=true"

    $postparams = @{
    'users' = @($Members)
    'groupId' = $($TeamId.groupId)
    }
    
    $result = Invoke-RestMethod -Uri $uri -Headers $TeamsAuthToken -Method Put -Body $($postparams|convertto-json)
    Write-Verbose "added team $displayName"
    Write-Verbose "$($result.value)"
    $result.value
  }
}

Export-ModuleMember connect-TeamsService, new-Team, add-TeamMember
