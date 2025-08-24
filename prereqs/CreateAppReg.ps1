# Check if Microsoft.Entra module is installed
$moduleName = "Microsoft.Entra"
$module = Get-Module -ListAvailable -Name $moduleName -ErrorAction SilentlyContinue

if (-not $module) {
    Write-Host "Required module '$moduleName' is not installed. Installing now..." -ForegroundColor Yellow
    try {
        Install-Module -Name $moduleName -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        Write-Host "Successfully installed module '$moduleName'." -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to install module '$moduleName': $_"
        Exit 1
    }
}
else {
    Write-Host "Module '$moduleName' is already installed." -ForegroundColor Green
}

# Connect to Entra (interactive)
try {
    Connect-Entra -Scopes 'Application.ReadWrite.All', 'DelegatedPermissionGrant.ReadWrite.All' -NoWelcome -ErrorAction Stop
}
catch {
    Write-Error "Failed to connect to Entra: $_"
    Exit 1
}

# Configuration - change these as needed for your environment
$apiAppName = "MessageCenterAgentSSO-api"
$clientAppName = "MessageCenterAgentSSO-client"
$teamsRedirectUri = 'https://teams.microsoft.com/api/platform/v1.0/oAuthRedirect'
# Backend redirect URI for local testing / Node.js backend (update for your dev-tunnel)
$backendRedirectUri = 'https://sx9lvcg1-3000.usw2.devtunnels.ms/auth/callback'

# Graph delegated permission required by the client app (to call Microsoft Graph)
$graphApiId = '00000003-0000-0000-c000-000000000000'
$delegatedPermission = 'ServiceMessage.Read.All'

Write-Host "Provisioning API app: $apiAppName and Client app: $clientAppName"

# --- Create or get API app (exposes scope) ---
$existingApiApp = Get-EntraApplication -Filter "DisplayName eq '$apiAppName'" -ErrorAction SilentlyContinue
if ($existingApiApp) {
    Write-Host "API app already exists: $($existingApiApp.AppId)" -ForegroundColor Yellow
    $apiApp = $existingApiApp
    $apiSp = Get-EntraServicePrincipal -Filter "AppId eq '$($apiApp.AppId)'"
} else {
    Write-Host "Creating API application: $apiAppName"
    $web = @{ redirectUris = @($backendRedirectUri) }
    $apiApp = New-EntraApplication -DisplayName $apiAppName -Web $web
    $apiSp = New-EntraServicePrincipal -AppId $apiApp.AppId

    # Expose an OAuth2 scope 'access_as_user' on the API app
    # Build a strongly-typed PermissionScope and ApiApplication so the Entra cmdlet receives
    # the exact CLR types it expects (List<PermissionScope>) instead of plain hashtables.
    $perm = New-Object Microsoft.Open.MSGraph.Model.PermissionScope
    $perm.AdminConsentDescription = 'Allow the application to access the backend API as the signed-in user'
    $perm.AdminConsentDisplayName = 'Access backend as user'
    $perm.Id = [guid]::NewGuid()
    $perm.IsEnabled = $true
    $perm.Type = 'User'
    $perm.Value = 'access_as_user'
    $perm.UserConsentDescription = 'Allow the application to access the backend API as the signed-in user'
    $perm.UserConsentDisplayName = 'Access backend as user'
    # AllowedMemberTypes defaults to ['User'] if omitted; do not assign directly to the CLR property here

    # Create ApiApplication and set its Oauth2PermissionScopes to a List<PermissionScope>
    $apiAppApi = New-Object Microsoft.Open.MSGraph.Model.ApiApplication
    $scopesList = New-Object 'System.Collections.Generic.List[Microsoft.Open.MSGraph.Model.PermissionScope]'
    $scopesList.Add($perm)
    $apiAppApi.Oauth2PermissionScopes = $scopesList

    # Update the application with the typed Api object
    Set-EntraApplication -ApplicationId $apiApp.Id -Api $apiAppApi

    Write-Host "Created API app: $($apiApp.AppId) and exposed scope 'access_as_user'"
}

# --- Create or get Client app (requests scopes and holds secret) ---
$existingClientApp = Get-EntraApplication -Filter "DisplayName eq '$clientAppName'" -ErrorAction SilentlyContinue
if ($existingClientApp) {
    Write-Host "Client app already exists: $($existingClientApp.AppId)" -ForegroundColor Yellow
    $clientApp = $existingClientApp
    $clientSp = Get-EntraServicePrincipal -Filter "AppId eq '$($clientApp.AppId)'"
} else {
    Write-Host "Creating Client application: $clientAppName"
    $web = @{ redirectUris = @($teamsRedirectUri, $backendRedirectUri) }
    $clientApp = New-EntraApplication -DisplayName $clientAppName -Web $web
    $clientSp = New-EntraServicePrincipal -AppId $clientApp.AppId

    # Create a client secret
    $secret = New-EntraApplicationPasswordCredential -ApplicationId $clientApp.Id -CustomKeyIdentifier "${clientAppName}-secret" -EndDate (Get-Date).AddYears(1)
    Write-Host "Created client secret for client app (secret value returned below)."
}

# --- Grant Microsoft Graph delegated permission to the client app ---
$graphServicePrincipal = Get-EntraServicePrincipal -Filter "AppId eq '$graphApiId'"
if (-not $graphServicePrincipal) {
    Write-Error "Could not find Microsoft Graph service principal."
} else {
    try {
        # Build RequiredResourceAccess object for Graph delegated permission
        $resourceAccessDelegated = New-Object Microsoft.Open.MSGraph.Model.ResourceAccess
        $resourceAccessDelegated.Id = ((Get-EntraServicePrincipal -ServicePrincipalId $graphServicePrincipal.Id).Oauth2PermissionScopes | Where-Object { $_.Value -eq $delegatedPermission }).Id
        $resourceAccessDelegated.Type = 'Scope'

        $requiredResourceAccessDelegated = New-Object Microsoft.Open.MSGraph.Model.RequiredResourceAccess
        $requiredResourceAccessDelegated.ResourceAppId = $graphApiId
        # ResourceAccess must be a List<ResourceAccess>
        $resourceAccessList = New-Object 'System.Collections.Generic.List[Microsoft.Open.MSGraph.Model.ResourceAccess]'
        $resourceAccessList.Add($resourceAccessDelegated)
        $requiredResourceAccessDelegated.ResourceAccess = $resourceAccessList

        # Wrap into a generic List<RequiredResourceAccess> because Set-EntraApplication expects that CLR type
        $reqList = New-Object 'System.Collections.Generic.List[Microsoft.Open.MSGraph.Model.RequiredResourceAccess]'
        $reqList.Add($requiredResourceAccessDelegated)

        Set-EntraApplication -ApplicationId $clientApp.Id -RequiredResourceAccess $reqList
        Write-Host "Assigned delegated Graph permission '$delegatedPermission' to client app"

        # Create an oauth2 permission grant (admin consent) for Graph permission
        $permissionGrantGraph = New-EntraOauth2PermissionGrant -ClientId $clientSp.Id -ConsentType 'AllPrincipals' -ResourceId $graphServicePrincipal.Id -Scope $delegatedPermission -ErrorAction SilentlyContinue
        if ($permissionGrantGraph) { Write-Host "Granted admin consent for Graph permission to client app" }
    }
    catch {
        Write-Warning "Failed to assign Graph permission: $_"
    }
}

# --- Grant client app admin consent to call API app scope (access_as_user) ---
try {
    $apiSp = Get-EntraServicePrincipal -Filter "AppId eq '$($apiApp.AppId)'"
    if ($apiSp) {
        $permissionGrantApi = New-EntraOauth2PermissionGrant -ClientId $clientSp.Id -ConsentType 'AllPrincipals' -ResourceId $apiSp.Id -Scope 'access_as_user' -ErrorAction SilentlyContinue
        if ($permissionGrantApi) { Write-Host "Granted admin consent for client to call API scope 'access_as_user'" }
    } else {
        Write-Warning "API service principal not found; cannot grant client access to API scope"
    }
}
catch {
    Write-Warning "Failed to create permission grant for API: $_"
}

# Output relevant details
$result = [ordered]@{
    apiAppId = $apiApp.AppId
    apiObjectId = $apiApp.Id
    clientAppId = $clientApp.AppId
    clientObjectId = $clientApp.Id
    tenantId = (Get-EntraTenantDetail).Id
}

# Attempt to include client secret value if created in this run
if ($secret) { $result.clientSecret = $secret.SecretText }

Write-Host (ConvertTo-Json $result -Depth 5)