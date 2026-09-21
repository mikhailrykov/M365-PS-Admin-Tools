<#
.SYNOPSIS
    Graph-first launcher for Remove-UnlicensedMailboxHolds.ps1.

.DESCRIPTION
    Establishes Microsoft Graph first, then Exchange Online with WAM disabled.
    This avoids the Microsoft.Identity/MSAL assembly conflict observed when the
    two modules authenticate in the opposite order. The existing remediation
    script is then invoked with the same parameters.

    Use this file while the original script remains unchanged. It intentionally
    leaves connections that existed before launch untouched and disconnects only
    connections created by this launcher.
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param (
    [switch]$DisableLitigationHold,
    [switch]$DisableInPlaceHold,
    [switch]$DisableAllHolds,
    [string]$ExportCsv,
    [string[]]$ReportTo,
    [string]$ReportFrom = 'copilot-noreply@ottawa.ca',
    [string]$ReportSubject = "Unlicensed Mailbox Holds Report — $(Get-Date -Format 'yyyy-MM-dd HH:mm')",
    [switch]$SkipConfirmation
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$scriptPath = Join-Path $PSScriptRoot 'Remove-UnlicensedMailboxHolds.ps1'
if (-not (Test-Path -LiteralPath $scriptPath -PathType Leaf)) {
    throw "The original script was not found: $scriptPath"
}

$graphConnectedByLauncher = $false
$exchangeConnectedByLauncher = $false

try {
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    Import-Module Microsoft.Graph.Users -ErrorAction Stop

    $graphContext = Get-MgContext -ErrorAction SilentlyContinue
    $hasGraphScope = $null -ne $graphContext -and
        -not [string]::IsNullOrWhiteSpace($graphContext.Account) -and
        @($graphContext.Scopes) -contains 'User.Read.All'

    if (-not $hasGraphScope) {
        if ($null -ne $graphContext -and -not [string]::IsNullOrWhiteSpace($graphContext.Account)) {
            throw "Microsoft Graph is already connected as '$($graphContext.Account)' without User.Read.All. Start a fresh PowerShell session or reconnect with User.Read.All."
        }

        Write-Host '[*] Connecting to Microsoft Graph first...' -ForegroundColor Cyan
        Connect-MgGraph -Scopes 'User.Read.All' -NoWelcome -ContextScope Process -ErrorAction Stop
        $graphConnectedByLauncher = $true
    }
    else {
        Write-Host "[*] Microsoft Graph: already connected as $($graphContext.Account)" -ForegroundColor Green
    }

    Import-Module ExchangeOnlineManagement -ErrorAction Stop

    $exchangeConnection = @(Get-ConnectionInformation -ErrorAction SilentlyContinue) |
        Where-Object { $_.State -eq 'Connected' } |
        Select-Object -First 1

    if ($exchangeConnection) {
        Write-Host "[*] Exchange Online: already connected as $($exchangeConnection.UserPrincipalName)" -ForegroundColor Green
    }
    else {
        $connectCommand = Get-Command Connect-ExchangeOnline -ErrorAction Stop
        $connectParameters = @{
            ShowBanner  = $false
            ErrorAction = 'Stop'
        }

        if ($connectCommand.Parameters.ContainsKey('DisableWAM')) {
            $connectParameters.DisableWAM = $true
        }
        else {
            Write-Warning 'Connect-ExchangeOnline does not support -DisableWAM. Update ExchangeOnlineManagement if authentication fails.'
        }

        Write-Host '[*] Connecting to Exchange Online after Microsoft Graph (WAM disabled)...' -ForegroundColor Cyan
        Connect-ExchangeOnline @connectParameters
        $exchangeConnectedByLauncher = $true
    }

    # Get-Mailbox is dynamically imported only after Exchange authentication.
    foreach ($commandName in @('Get-Mailbox', 'Set-Mailbox', 'Get-MailboxSearch', 'Set-MailboxSearch')) {
        if (-not (Get-Command -Name $commandName -ErrorAction SilentlyContinue)) {
            throw "Required Exchange command '$commandName' was not loaded after connecting to Exchange Online."
        }
    }

    $invokeParameters = @{}
    foreach ($parameterName in @(
        'DisableLitigationHold', 'DisableInPlaceHold', 'DisableAllHolds',
        'ExportCsv', 'ReportTo', 'ReportFrom', 'ReportSubject', 'SkipConfirmation'
    )) {
        if ($PSBoundParameters.ContainsKey($parameterName)) {
            $invokeParameters[$parameterName] = $PSBoundParameters[$parameterName]
        }
    }

    if ($WhatIfPreference) {
        $invokeParameters['WhatIf'] = $true
    }

    & $scriptPath @invokeParameters
}
finally {
    if ($exchangeConnectedByLauncher) {
        try {
            Disconnect-ExchangeOnline -Confirm:$false -ErrorAction Stop
            Write-Host '[*] Exchange Online: disconnected (launcher-created connection).' -ForegroundColor Gray
        }
        catch {
            Write-Warning "Unable to disconnect launcher-created Exchange Online connection: $($_.Exception.Message)"
        }
    }

    if ($graphConnectedByLauncher) {
        try {
            Disconnect-MgGraph -ErrorAction Stop | Out-Null
            Write-Host '[*] Microsoft Graph: disconnected (launcher-created connection).' -ForegroundColor Gray
        }
        catch {
            Write-Warning "Unable to disconnect launcher-created Microsoft Graph connection: $($_.Exception.Message)"
        }
    }
}
