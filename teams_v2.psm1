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
  #$clientId = '5e3ce6c0-2b1f-4285-8d4b-75ee78787346'
  $clientId = 'd3590ed6-52b3-4102-aeff-aad2292ab01c'
  #$redirectUri = "https://teams.microsoft.com/go"
  $redirectUri = "urn:ietf:wg:oauth:2.0:oob"
  #$redirectUri = "urn:federation:MicrosoftOnline"
  $resourceAppIdURI = "https://api.spaces.skype.com"
  $authority = "https://login.windows.net/$Tenant"
  try {
    $authContext = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext" -ArgumentList $authority
    # https://msdn.microsoft.com/en-us/library/azure/microsoft.identitymodel.clients.activedirectory.promptbehavior.aspx
    # Change the prompt behaviour to force credentials each time: Auto, Always, Never, RefreshSession
    $platformParameters = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.PlatformParameters" -ArgumentList "Always"
    $userId = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.UserIdentifier" -ArgumentList ($User, "OptionalDisplayableId")
    #New-Object 'Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationParameters'
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
        $TeamsauthHeader = @{
        'Content-Type'='application/json'
        'Authorization'="Bearer " + $authResult.AccessToken
        'ExpiresOn'=$authResult.ExpiresOn
        'X-Skypetoken' = "$((Invoke-RestMethod -Uri 'https://api.teams.skype.com/beta/auth/skypetoken' -Method Post -Headers $authHeader).tokens.skypetoken)"
      }
      $global:TeamsAuthToken = $authHeader
      $global:TeamsAuthToken2 = $TeamsauthHeader
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
        [Parameter(Mandatory=$true)]$smtpaddress,
        [Parameter(Mandatory=$true)]$alias
  )
  begin {
  write-verbose "checking for teams token"
    if (!($TeamsAuthToken2))
    {throw 'please run connect-TeamsService first'}
  }
  process {
    write-verbose "start to invoke rest request"
    $uri = 'https://api.teams.skype.com/emea/beta/teams/create'
    $postparams = @{
    'alias' = $alias
    'description' = $description
    'displayName' = $displayname
    'smtpAddress'=  $smtpaddress
    'AccessType' = 1
    }
    Invoke-RestMethod -Uri $uri -Headers $TeamsAuthToken2 -Method post -Body $($postparams|convertto-json)
    write-debug "added team $displayName"
  }
}

Export-ModuleMember connect-TeamsService, new-team