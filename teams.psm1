function connect-TeamsService {
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
    [Parameter(Mandatory=$true)]$User,
    [Parameter(Mandatory=$true)]$tenant
  )
  Write-Verbose "Checking for AzureAD module..."
  $AadModule = Get-Module -Name "AzureAD" -ListAvailable
 if ($AadModule -eq $null) {
    write-warning "AzureAD Powershell module not installed..."
    Write-Warning "Install by running 'Install-Module AzureAD' from an elevated PowerShell prompt"
    Write-Warning "Script can't continue..."
    exit
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
  #$clientId = 'cc15fd57-2c6c-4117-a88c-83b1d56b4bbe'
  $clientId = 'd1ddf0e4-d672-4dae-b554-9d5bdfd93547'
  $redirectUri = "urn:ietf:wg:oauth:2.0:oob"
  $resourceAppIdURI = "https://graph.microsoft.com"
  $authority = "https://login.windows.net/$Tenant"
  try {
    $authContext = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext" -ArgumentList $authority
    # https://msdn.microsoft.com/en-us/library/azure/microsoft.identitymodel.clients.activedirectory.promptbehavior.aspx
    # Change the prompt behaviour to force credentials each time: Auto, Always, Never, RefreshSession
    $platformParameters = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.PlatformParameters" -ArgumentList "Always"
    $userId = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.UserIdentifier" -ArgumentList ($User, "OptionalDisplayableId")
    #New-Object 'Microsoft.IdentityModel.Clients.ActiveDirectory.Authentication​Parameters'
    #$query = 'wauth=http://schemas.microsoft.com/ws/2008/06/identity/authenticationmethod/password'
    $authResult = $authContext.AcquireTokenAsync($resourceAppIdURI,$clientId,$redirectUri,$platformParameters,$userId).Result
    # If the accesstoken is valid then create the authentication header
    if($authResult.AccessToken){
      # Creating header for Authorization token
      $authHeader = @{
        'Content-Type'='application/json'
        'Authorization'="Bearer " + $authResult.AccessToken
        'ExpiresOn'=$authResult.ExpiresOn
      }
      $global:TeamsAuthToken = $authHeader
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

function get-team {
  <#
  .SYNOPSIS
  Invoking a rest request to the Microsoft graph api to get current teams
  .DESCRIPTION
  Describe the function in more detail
  .EXAMPLE
  Give an example of how to use it
  #>
  [CmdletBinding()]
  #[CmdletBinding(SupportsShouldProcess=$True,ConfirmImpact='Low')]
  param
  (
        [Parameter(Mandatory=$false)]$team
  )
  begin {
  write-verbose "checking for teams token"
    if (!($TeamsAuthToken))
    {throw 'please run connect-TeamsService first'}
  }
  process {
    write-verbose "start to invoke rest request"
    $nextlink = $true
    $uri = "https://graph.microsoft.com/beta/groups?$('$filter')=groupTypes/any(c:c+eq+'Unified')"
    $teams = Invoke-RestMethod -Uri $uri -Headers $TeamsAuthToken -Method get
    while ($nextlink -eq $true){
    if ($($teams.'@odata.nextLink') -ne $null){
      $results = $teams.value
      $teams = Invoke-RestMethod -Uri $($teams.'@odata.nextLink')-Headers $TeamsAuthToken -Method get
    }
    if ($($teams.'@odata.nextLink') -eq $null){
      $results = $teams.value
      $teams = $null
    $nextlink = $false}
    write-debug "getting info $teams"
    $objects += $results
  }
  $objects
  }
}
function add-Team {
  <#
  .SYNOPSIS
  Invoking a rest request to the Microsoft graph api to add a team
  .DESCRIPTION
  Describe the function in more detail
  .EXAMPLE
  Give an example of how to use it
  #>
  [CmdletBinding()]
  #[CmdletBinding(SupportsShouldProcess=$True,ConfirmImpact='Low')]
  param
  (
        [Parameter(Mandatory=$true)]$description,
        [Parameter(Mandatory=$true)]$displayname,
        [Parameter(Mandatory=$true)][boolean]$mailEnabled,
        [Parameter(Mandatory=$true)][boolean]$securityEnabled,
        [Parameter(Mandatory=$true)]$mailNickname
  )
  begin {
  write-verbose "checking for teams token"
    if (!($TeamsAuthToken))
    {throw 'please run connect-TeamsService first'}
  }
  process {
    write-verbose "start to invoke rest request"
    $uri = 'https://graph.microsoft.com/beta/groups'
    $postparams = @{
    description = $description
    displayName = $displayname
    mailEnabled = $mailEnabled
    mailNickname = $mailNickname
    securityEnabled = $securityEnabled
    groupTypes = @("Unified")
    }
    Invoke-RestMethod -Uri $uri -Headers $TeamsAuthToken -Method post -Body $($postparams|convertto-json)
    write-debug "added team $displayName"
  }
}

function remove-Team {
  <#
  .SYNOPSIS
  Invoking a rest request to the Microsoft graph api to add a team
  .DESCRIPTION
  Describe the function in more detail
  .EXAMPLE
  Give an example of how to use it
  #>
  [CmdletBinding()]
  #[CmdletBinding(SupportsShouldProcess=$True,ConfirmImpact='Low')]
  param
  (
        [Parameter(Mandatory=$true)]$teamid
  )
  begin {
  write-verbose "checking for teams token"
    if (!($TeamsAuthToken))
    {throw 'please run connect-TeamsService first'}
  }
  process {
    write-verbose "start to invoke rest request"
    $uri = "https://graph.microsoft.com/beta/groups/$teamid"
    Invoke-RestMethod -Uri $uri -Headers $TeamsAuthToken -Method delete
    write-debug "removing team $teamid"
  }
}
function get-teammembers {
  <#
  .SYNOPSIS
  Invoking a rest request to the Microsoft graph api to get current teammembers
  .DESCRIPTION
  Describe the function in more detail
  .EXAMPLE
  Give an example of how to use it
  #>
  [CmdletBinding()]
  #[CmdletBinding(SupportsShouldProcess=$True,ConfirmImpact='Low')]
  param
  (
        [Parameter(Mandatory=$true)]$teamid
  )
  begin {
  write-verbose "checking for teams token"
    if (!($TeamsAuthToken))
    {throw 'please run connect-TeamsService first'}
  }
  process {
    write-verbose "start to invoke rest request"
    $nextlink = $true
    $uri = "https://graph.microsoft.com/beta/groups/$teamid/members"
    $teams = Invoke-RestMethod -Uri $uri -Headers $TeamsAuthToken -Method get
    while ($nextlink -eq $true){
    if ($($teams.'@odata.nextLink') -ne $null){
      $results = $teams.value
      $teams = Invoke-RestMethod -Uri $($teams.'@odata.nextLink')-Headers $TeamsAuthToken -Method get
    }
    if ($($teams.'@odata.nextLink') -eq $null){
      $results = $teams.value
      $teams = $null
    $nextlink = $false}
    write-debug "getting groupmember for $teamid"
    $objects += $results
  }
  $objects
  }
}

function add-TeamMember {
  <#
  .SYNOPSIS
  Invoking a rest request to the Microsoft graph api to add teammembers
  .DESCRIPTION
  Describe the function in more detail
  .EXAMPLE
  Give an example of how to use it
  #>
  [CmdletBinding()]
  #[CmdletBinding(SupportsShouldProcess=$True,ConfirmImpact='Low')]
  param
  (
        [Parameter(Mandatory=$true)]$teamid,
        [Parameter(Mandatory=$true)]$member
  )
  begin {
  write-verbose "checking for teams token"
    if (!($TeamsAuthToken))
    {throw 'please run connect-TeamsService first'}
  }
  process {
    write-verbose "start to invoke rest request"
    $uri = "https://graph.microsoft.com/beta/groups/$teamid/members/" + '$ref'
    $postparams = @{
    '@odata.id' = "https://graph.microsoft.com/beta/users/$member"
    }
    Invoke-RestMethod -Uri $uri -Headers $TeamsAuthToken -Method post -Body $($postparams|convertto-json)
    write-debug "adding teammember $member for $teamid"
  }
}

function add-TeamOwner {
  <#
  .SYNOPSIS
  Invoking a rest request to the Microsoft graph api to add TeamOwner
  .DESCRIPTION
  Describe the function in more detail
  .EXAMPLE
  Give an example of how to use it
  #>
  [CmdletBinding()]
  #[CmdletBinding(SupportsShouldProcess=$True,ConfirmImpact='Low')]
  param
  (
        [Parameter(Mandatory=$true)]$teamid,
        [Parameter(Mandatory=$true)]$member
  )
  begin {
  write-verbose "checking for teams token"
    if (!($TeamsAuthToken))
    {throw 'please run connect-TeamsService first'}
  }
  process {
    write-verbose "start to invoke rest request"
    $uri = "https://graph.microsoft.com/beta/groups/$teamid/owners/" + '$ref'
    $postparams = @{
    '@odata.id' = "https://graph.microsoft.com/beta/users/$member"
    }
    Invoke-RestMethod -Uri $uri -Headers $TeamsAuthToken -Method post -Body $($postparams|convertto-json)
    write-debug "adding owner $member for $teamid"
  }
}
function remove-TeamMember {
  <#
  .SYNOPSIS
  Invoking a rest request to the Microsoft graph api to remove teammembers
  .DESCRIPTION
  Describe the function in more detail
  .EXAMPLE
  Give an example of how to use it
  #>
  [CmdletBinding()]
  #[CmdletBinding(SupportsShouldProcess=$True,ConfirmImpact='Low')]
  param
  (
        [Parameter(Mandatory=$true)]$teamid,
        [Parameter(Mandatory=$true)]$member
  )
  begin {
  write-verbose "checking for teams token"
    if (!($TeamsAuthToken))
    {throw 'please run connect-TeamsService first'}
  }
  process {
    write-verbose "start to invoke rest request"
    if ($member -like '*@*'){
      Write-Debug "found UPN instead of ID, converting..."
      $member = (Invoke-RestMethod -Uri "https://graph.microsoft.com/beta/users/$member/id" -Headers $TeamsAuthToken -Method get).value
    }
        $uri = "https://graph.microsoft.com/beta/groups/$teamid/members/$member" + '/$ref'
        Invoke-RestMethod -Uri $uri -Headers $TeamsAuthToken -Method Delete -Body $postparams
        write-debug "removing teammember $member for $teamid"
  }
}

function remove-TeamOwner {
  <#
  .SYNOPSIS
  Invoking a rest request to the Microsoft graph api to remove teamOwners
  .DESCRIPTION
  Describe the function in more detail
  .EXAMPLE
  Give an example of how to use it
  #>
  [CmdletBinding()]
  #[CmdletBinding(SupportsShouldProcess=$True,ConfirmImpact='Low')]
  param
  (
        [Parameter(Mandatory=$true)]$teamid,
        [Parameter(Mandatory=$true)]$member
  )
  begin {
  write-verbose "checking for teams token"
    if (!($TeamsAuthToken))
    {throw 'please run connect-TeamsService first'}
  }
  process {
    write-verbose "start to invoke rest request"
    if ($member -like '*@*'){
      Write-Debug "found UPN instead of ID, converting..."
      $member = (Invoke-RestMethod -Uri "https://graph.microsoft.com/beta/users/$member/id" -Headers $TeamsAuthToken -Method get).value
    }
        $uri = "https://graph.microsoft.com/beta/groups/$teamid/owners/$member" + '/$ref'
        Invoke-RestMethod -Uri $uri -Headers $TeamsAuthToken -Method Delete -Body $postparams
        write-debug "removing teamowner $member for $teamid"
  }
}
Export-ModuleMember connect-TeamsService, get-team, get-teammembers, add-TeamMember, remove-teammember, add-TeamOwner, remove-TeamOwner, add-Team, remove-team