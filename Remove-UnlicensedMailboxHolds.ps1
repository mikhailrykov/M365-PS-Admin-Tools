<#
.SYNOPSIS
    Finds unlicensed Exchange Online mailboxes with Litigation Hold or In-Place Holds.

.DESCRIPTION
    Connects to Microsoft Graph first, then Exchange Online with WAM disabled when
    supported, to avoid a known dependency assembly conflict between
    Microsoft.Graph.Authentication and ExchangeOnlineManagement.

    By default, both Litigation Hold and In-Place Hold remediation are enabled.
    Specify only -DisableLitigationHold or only -DisableInPlaceHold to limit
    remediation to that hold type. Use -DisableAllHolds:$false for report-only mode.

    The script:
      * Finds Litigation Hold mailboxes and active In-Place Holds.
      * Checks affected user accounts for assigned Microsoft 365 licenses through Graph.
      * Reports unlicensed accounts with holds.
      * Optionally disables Litigation Hold and/or removes mailbox membership from
        non-org-wide In-Place Holds.
      * Optionally exports and emails a CSV report when -ReportTo is supplied.

    Connection behavior:
      * Existing Graph and EXO connections are reused.
      * Only connections created by this script are disconnected at completion.
      * If an existing Graph connection lacks User.Read.All, the script stops rather
        than replacing or disconnecting the pre-existing connection.
      * Exchange authentication uses the Graph-first pattern and, when the installed
        module version supports it, passes -DisableWAM to Connect-ExchangeOnline.

.PARAMETER DisableLitigationHold
    Disable Litigation Hold for qualifying unlicensed mailboxes. When supplied by
    itself, In-Place Hold remediation and DisableAllHolds are disabled.

.PARAMETER DisableInPlaceHold
    Remove qualifying unlicensed mailboxes from non-org-wide In-Place Holds. When
    supplied by itself, Litigation Hold remediation and DisableAllHolds are disabled.

.PARAMETER DisableAllHolds
    Enable both Litigation Hold and In-Place Hold remediation. Defaults to true.
    Specify -DisableAllHolds:$false for report-only mode when neither individual
    remediation switch is supplied.

.PARAMETER ExportCsv
    Optional destination for the CSV report. If omitted while -ReportTo is provided,
    a temporary CSV is created, emailed, and removed.

.PARAMETER ReportTo
    One or more report recipients. Providing this parameter enables email reporting.
    Each entry must be a plain email address in the form user@example.com.

.PARAMETER ReportFrom
    Sender address for the SMTP relay. Must be a valid email address. When omitted,
    defaults to USERNAME@USERDNSDOMAIN from the current environment.

.PARAMETER ReportSubject
    Subject for the emailed report.

.PARAMETER SkipConfirmation
    Suppress confirmation prompts for changes. Intended for unattended execution.

.PARAMETER WhatIf
    Performs discovery, CSV generation, and optional email reporting, but makes no
    changes to Litigation Holds or In-Place Holds.

.EXAMPLE
    # Safely preview the default remediation of both hold types.
    .\Remove-UnlicensedMailboxHolds.ps1 -WhatIf

.EXAMPLE
    # Report only; no remediation.
    .\Remove-UnlicensedMailboxHolds.ps1 -DisableAllHolds:$false

.EXAMPLE
    # Preview Litigation Hold remediation only and email the report.
    .\Remove-UnlicensedMailboxHolds.ps1 -WhatIf -DisableLitigationHold `
        -ReportTo rykov@ottawa.ca

.EXAMPLE
    # Apply both remediation types without interactive confirmation and send a report.
    .\Remove-UnlicensedMailboxHolds.ps1 -SkipConfirmation -ReportTo rykov@ottawa.ca

.NOTES
    Required modules:
      Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
      Install-Module Microsoft.Graph.Users -Scope CurrentUser
      Install-Module ExchangeOnlineManagement -Scope CurrentUser

    Required permissions:
      Microsoft Graph: User.Read.All
      Exchange Online: Recipient Management and eDiscovery permissions appropriate
      for reviewing or changing holds.

    SMTP relay: appsmtp.ottawa.ca
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param (
    [switch]$DisableLitigationHold,
    [switch]$DisableInPlaceHold,
    [switch]$DisableAllHolds = $true,

    [string]$ExportCsv,

    [string[]]$ReportTo,

    [string]$ReportFrom,

    [string]$ReportSubject = "Unlicensed Mailbox Holds Report — $(Get-Date -Format 'yyyy-MM-dd HH:mm')",

    [switch]$SkipConfirmation
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

#region Prerequisite and connection helpers

function Assert-PowerShellRequirements {
    if ($PSVersionTable.PSEdition -ne 'Core') {
        throw "PowerShell 7.4 or later is required. Start the script using pwsh.exe."
    }

    if ($PSVersionTable.PSVersion -lt [version]'7.4.0') {
        throw "PowerShell 7.4 or later is required. Current version: $($PSVersionTable.PSVersion)"
    }

    Write-Verbose "PowerShell requirement satisfied: $($PSVersionTable.PSEdition) $($PSVersionTable.PSVersion)"
}

function Assert-ModuleAvailable {
    param(
        [Parameter(Mandatory)]
        [string]$ModuleName
    )

    $loadedModule = Get-Module -Name $ModuleName
    if ($loadedModule) {
        Write-Verbose "Module '$ModuleName' is already loaded: version $($loadedModule.Version)."
        return
    }

    $availableModule = Get-Module -ListAvailable -Name $ModuleName |
        Sort-Object Version -Descending |
        Select-Object -First 1

    if (-not $availableModule) {
        throw @"
Required module '$ModuleName' is not installed.

Install it with:
    Install-Module $ModuleName -Scope CurrentUser
"@
    }

    try {
        Write-Verbose "Importing module '$ModuleName' version $($availableModule.Version) from '$($availableModule.ModuleBase)'."
        Import-Module -Name $ModuleName -RequiredVersion $availableModule.Version -ErrorAction Stop
        Write-Host "[+] Module loaded: $ModuleName ($($availableModule.Version))" -ForegroundColor Green
    }
    catch {
        throw "Required module '$ModuleName' is installed but could not be loaded. $($_.Exception.Message)"
    }
}

function Assert-CommandAvailable {
    param(
        [Parameter(Mandatory)]
        [string]$CommandName,

        [Parameter(Mandatory)]
        [string]$ModuleName
    )

    if (-not (Get-Command -Name $CommandName -ErrorAction SilentlyContinue)) {
        throw "Required command '$CommandName' is unavailable after loading '$ModuleName'. Reinstall or update module '$ModuleName'."
    }
}

function Assert-ValidEmailAddress {
    param(
        [Parameter(Mandatory)]
        [string]$Address,

        [Parameter(Mandatory)]
        [string]$ParameterName
    )

    if ([string]::IsNullOrWhiteSpace($Address)) {
        throw "$ParameterName is empty. Provide a plain email address such as user@example.com."
    }

    $trimmed = $Address.Trim()

    try {
        $parsedAddress = [System.Net.Mail.MailAddress]::new($trimmed)
    }
    catch {
        throw "$ParameterName '$Address' is not a valid email address. Use a plain address such as user@example.com."
    }

    if ($parsedAddress.Address -ine $trimmed) {
        throw "$ParameterName '$Address' is not a valid plain email address. Display-name formats such as 'Name <user@example.com>' are not supported."
    }

    return $trimmed
}

function Assert-ValidEmailAddressList {
    param(
        [Parameter()]
        [string[]]$Addresses,

        [Parameter(Mandatory)]
        [string]$ParameterName
    )

    if ($null -eq $Addresses) {
        return @()
    }

    $normalized = [System.Collections.Generic.List[string]]::new()

    foreach ($candidate in $Addresses) {
        if ($null -eq $candidate) {
            continue
        }

        $trimmed = $candidate.Trim()
        if ([string]::IsNullOrWhiteSpace($trimmed)) {
            throw "$ParameterName contains an empty entry. Each address must be a plain email address such as user@example.com."
        }

        $normalized.Add((Assert-ValidEmailAddress -Address $trimmed -ParameterName $ParameterName))
    }

    return @($normalized)
}

function Get-GraphConnectionState {
    $state = @{
        AlreadyConnected = $false
        MissingScope     = $false
        Account          = $null
        Scopes           = @()
    }

    $context = Get-MgContext -ErrorAction SilentlyContinue
    if ($null -eq $context -or [string]::IsNullOrWhiteSpace($context.Account)) {
        return $state
    }

    $state.Account = $context.Account
    $state.Scopes  = @($context.Scopes)

    if ($context.Scopes -contains 'User.Read.All') {
        $state.AlreadyConnected = $true
    }
    else {
        $state.MissingScope = $true
    }

    return $state
}

function Connect-GraphService {
    param(
        [Parameter(Mandatory)]
        [hashtable]$State
    )

    if ($State.AlreadyConnected) {
        Write-Host "[*] Microsoft Graph: already connected as $($State.Account)" -ForegroundColor Green
        Write-Verbose "Reusing existing Graph connection with User.Read.All."
        return $false
    }

    if ($State.MissingScope) {
        throw @"
Microsoft Graph is already connected as '$($State.Account)', but the connection does not
have the required User.Read.All delegated scope.

The script will not disconnect or replace a pre-existing Graph connection. Start a new
PowerShell session, or manually reconnect Graph with User.Read.All before running it.

Current scopes:
$($State.Scopes -join ', ')
"@
    }

    Write-Host "[*] Connecting to Microsoft Graph..." -ForegroundColor Cyan
    Connect-MgGraph -Scopes 'User.Read.All' -NoWelcome -ContextScope Process -ErrorAction Stop

    $context = Get-MgContext -ErrorAction Stop
    if ($null -eq $context -or $context.Scopes -notcontains 'User.Read.All') {
        throw "Microsoft Graph connected, but User.Read.All is not present in the active context."
    }

    Write-Host "[+] Microsoft Graph: connected as $($context.Account)" -ForegroundColor Green
    return $true
}

function Get-ExchangeConnectionState {
    $state = @{
        AlreadyConnected = $false
        UserPrincipalName = $null
        Organization = $null
    }

    $connections = @(Get-ConnectionInformation -ErrorAction SilentlyContinue)
    $connection = $connections |
        Where-Object { $_.State -eq 'Connected' } |
        Select-Object -First 1

    if ($connection) {
        $state.AlreadyConnected  = $true
        $state.UserPrincipalName = $connection.UserPrincipalName
        $state.Organization      = $connection.Organization
    }

    return $state
}

function Connect-ExchangeService {
    param(
        [Parameter(Mandatory)]
        [hashtable]$State
    )

    if ($State.AlreadyConnected) {
        Write-Host "[*] Exchange Online: already connected as $($State.UserPrincipalName)" -ForegroundColor Green
        Write-Verbose "Reusing existing EXO connection for organization '$($State.Organization)'."
        return $false
    }

    $connectCommand = Get-Command Connect-ExchangeOnline -ErrorAction SilentlyContinue
    $connectParameters = @{
        ShowBanner = $false
        ErrorAction = 'Stop'
    }

    if ($null -ne $connectCommand -and $connectCommand.Parameters.ContainsKey('DisableWAM')) {
        $connectParameters['DisableWAM'] = $true
        Write-Verbose "Connect-ExchangeOnline supports -DisableWAM; using it to avoid the Graph/EXO auth assembly conflict."
    }
    else {
        Write-Warning "Connect-ExchangeOnline does not support -DisableWAM. Update ExchangeOnlineManagement if authentication fails. Continuing without it."
    }

    Write-Host "[*] Connecting to Exchange Online after Microsoft Graph..." -ForegroundColor Cyan
    Connect-ExchangeOnline @connectParameters
    Write-Host "[+] Exchange Online: connected." -ForegroundColor Green
    return $true
}

function Disconnect-ScriptServices {
    param(
        [bool]$GraphConnectedByScript,
        [bool]$ExchangeConnectedByScript
    )

    if ($ExchangeConnectedByScript) {
        try {
            Disconnect-ExchangeOnline -Confirm:$false -ErrorAction Stop
            Write-Host "[*] Exchange Online: disconnected (script-created connection)." -ForegroundColor Gray
        }
        catch {
            Write-Warning "Unable to disconnect script-created Exchange Online connection: $($_.Exception.Message)"
        }
    }
    else {
        Write-Verbose "Exchange Online was pre-existing or was never connected; leaving it unchanged."
    }

    if ($GraphConnectedByScript) {
        try {
            Disconnect-MgGraph -ErrorAction Stop | Out-Null
            Write-Host "[*] Microsoft Graph: disconnected (script-created connection)." -ForegroundColor Gray
        }
        catch {
            Write-Warning "Unable to disconnect script-created Microsoft Graph connection: $($_.Exception.Message)"
        }
    }
    else {
        Write-Verbose "Microsoft Graph was pre-existing or was never connected; leaving it unchanged."
    }
}

#endregion

#region Hold and report helpers

function Get-InPlaceHoldsForMailbox {
    param(
        [Parameter(Mandatory)]
        [string]$UserPrincipalName,

        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [object[]]$AllInPlaceHolds
    )

    $matchingHolds = @()

    foreach ($hold in $AllInPlaceHolds) {
        $sourceMailboxes = @($hold.SourceMailboxes)
        $isOrgWide = $sourceMailboxes.Count -eq 0

        if ($isOrgWide) {
            $matchingHolds += $hold
            continue
        }

        if ($sourceMailboxes | Where-Object { $_.ToString() -ieq $UserPrincipalName }) {
            $matchingHolds += $hold
        }
    }

    return $matchingHolds
}

function Remove-MailboxFromInPlaceHold {
    param(
        [Parameter(Mandatory)]
        [string]$UserPrincipalName,

        [Parameter(Mandatory)]
        [object]$Hold
    )

    $sourceMailboxes = @($Hold.SourceMailboxes)
    if ($sourceMailboxes.Count -eq 0) {
        Write-Warning "Hold '$($Hold.Name)' is org-wide. A single mailbox cannot be removed automatically."
        return [PSCustomObject]@{
            HoldName   = $Hold.Name
            HoldStatus = 'Skipped-OrgWide'
        }
    }

    $updatedSources = @(
        $sourceMailboxes |
            Where-Object { $_.ToString() -ine $UserPrincipalName }
    )

    $operation = if ($updatedSources.Count -gt 0) {
        "Remove '$UserPrincipalName' from In-Place Hold '$($Hold.Name)'"
    }
    else {
        "Disable In-Place Hold '$($Hold.Name)' because '$UserPrincipalName' is its only source mailbox"
    }

    if (-not $PSCmdlet.ShouldProcess($UserPrincipalName, $operation)) {
        return [PSCustomObject]@{
            HoldName   = $Hold.Name
            HoldStatus = 'WhatIf-WouldRemove'
        }
    }

    try {
        if ($updatedSources.Count -gt 0) {
            Set-MailboxSearch -Identity $Hold.Name -SourceMailboxes $updatedSources -ErrorAction Stop
            $status = 'RemovedFromHold'
        }
        else {
            Set-MailboxSearch -Identity $Hold.Name -InPlaceHoldEnabled $false -ErrorAction Stop
            $status = 'HoldDisabled-OnlyMailbox'
        }

        return [PSCustomObject]@{
            HoldName   = $Hold.Name
            HoldStatus = $status
        }
    }
    catch {
        return [PSCustomObject]@{
            HoldName   = $Hold.Name
            HoldStatus = "Failed: $($_.Exception.Message)"
        }
    }
}

function Send-HoldReport {
    param(
        [Parameter(Mandatory)]
        [string]$CsvPath,

        [Parameter(Mandatory)]
        [string[]]$To,

        [Parameter(Mandatory)]
        [string]$From,

        [Parameter(Mandatory)]
        [string]$Subject,

        [Parameter(Mandatory)]
        [string]$Mode,

        [Parameter(Mandatory)]
        [int]$ResultCount
    )

    $whatIfText = if ($WhatIfPreference) {
        "`r`nIMPORTANT: This was a WhatIf run. No account or hold changes were made."
    }
    else {
        ''
    }

    $body = @"
Unlicensed Mailbox Holds Report
================================
Generated : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
Mode      : $Mode
Results   : $ResultCount unlicensed mailbox(es) with hold(s) found.$whatIfText

The attached CSV contains the full report.
"@

    $emailSubject = if ($WhatIfPreference) {
        "[WhatIf] $Subject"
    }
    else {
        $Subject
    }

    Send-MailMessage `
        -SmtpServer 'appsmtp.ottawa.ca' `
        -From $From `
        -To $To `
        -Subject $emailSubject `
        -Body $body `
        -Attachments $CsvPath `
        -Encoding ([System.Text.Encoding]::UTF8) `
        -ErrorAction Stop
}

#endregion

#region Initialization

$litigationHoldWasSpecified = $PSBoundParameters.ContainsKey('DisableLitigationHold')
$inPlaceHoldWasSpecified = $PSBoundParameters.ContainsKey('DisableInPlaceHold')

if ($litigationHoldWasSpecified -xor $inPlaceHoldWasSpecified) {
    # An explicitly supplied individual switch overrides the all-holds default.
    $DisableAllHolds = $false

    if ($litigationHoldWasSpecified) {
        $DisableInPlaceHold = $false
    }
    else {
        $DisableLitigationHold = $false
    }
}
elseif ($litigationHoldWasSpecified -and $inPlaceHoldWasSpecified) {
    $DisableAllHolds = [bool]($DisableLitigationHold -and $DisableInPlaceHold)
}
elseif ($DisableAllHolds) {
    $DisableLitigationHold = $true
    $DisableInPlaceHold = $true
}
else {
    $DisableLitigationHold = $false
    $DisableInPlaceHold = $false
}

if ([string]::IsNullOrWhiteSpace($ReportFrom)) {
    if ([string]::IsNullOrWhiteSpace($env:USERNAME) -or
        [string]::IsNullOrWhiteSpace($env:USERDNSDOMAIN)) {
        throw "ReportFrom was not provided and USERNAME or USERDNSDOMAIN is unavailable. Specify -ReportFrom with a valid email address."
    }

    $ReportFrom = '{0}@{1}' -f $env:USERNAME.Trim(), $env:USERDNSDOMAIN.Trim()
}

$ReportFrom = Assert-ValidEmailAddress -Address $ReportFrom -ParameterName 'ReportFrom'

if ($PSBoundParameters.ContainsKey('ReportTo')) {
    $ReportTo = Assert-ValidEmailAddressList -Addresses $ReportTo -ParameterName 'ReportTo'
    if ($ReportTo.Count -eq 0) {
        throw "ReportTo was provided, but no valid recipient addresses were supplied. Use plain email addresses such as user@example.com."
    }
}

if ($SkipConfirmation) {
    $ConfirmPreference = 'None'
    Write-Verbose "SkipConfirmation enabled."
}

$sendReport = $PSBoundParameters.ContainsKey('ReportTo') -and $ReportTo.Count -gt 0

$temporaryCsv = $false
$effectiveCsv = $ExportCsv

if ($sendReport -and [string]::IsNullOrWhiteSpace($effectiveCsv)) {
    $effectiveCsv = Join-Path ([System.IO.Path]::GetTempPath()) (
        "UnlicensedMailboxHolds-{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')
    )
    $temporaryCsv = $true
}

$graphConnectedByScript = $false
$exchangeConnectedByScript = $false
$preflightGraphState = $null
$preflightExchangeState = $null
$results = [System.Collections.Generic.List[PSCustomObject]]::new()

#endregion

try {
    Write-Verbose "=== Script started: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ==="

    if ($WhatIfPreference) {
        Write-Host "`n[!] WhatIf mode: discovery, CSV export, and email reporting will run; no hold changes will be made.`n" -ForegroundColor DarkYellow
    }

    Assert-PowerShellRequirements

    # IMPORTANT:
    # Graph must be imported and connected BEFORE EXO is imported. EXO can load
    # conflicting Microsoft.Identity.Client / IdentityModel assemblies otherwise.
    Write-Host "[*] Loading Microsoft Graph components first..." -ForegroundColor Cyan
    Assert-ModuleAvailable -ModuleName 'Microsoft.Graph.Authentication'
    Assert-ModuleAvailable -ModuleName 'Microsoft.Graph.Users'

    Assert-CommandAvailable -CommandName 'Connect-MgGraph' -ModuleName 'Microsoft.Graph.Authentication'
    Assert-CommandAvailable -CommandName 'Get-MgUser' -ModuleName 'Microsoft.Graph.Users'

    Write-Host "[*] Checking existing Microsoft Graph connection..." -ForegroundColor Cyan
    $preflightGraphState = Get-GraphConnectionState
    $graphConnectedByScript = Connect-GraphService -State $preflightGraphState

    # Exchange Online is deliberately handled only after Graph authentication.
    Write-Host "[*] Loading Exchange Online components..." -ForegroundColor Cyan
    Assert-ModuleAvailable -ModuleName 'ExchangeOnlineManagement'

    Assert-CommandAvailable -CommandName 'Connect-ExchangeOnline' -ModuleName 'ExchangeOnlineManagement'

    Write-Host "[*] Checking existing Exchange Online connection..." -ForegroundColor Cyan
    $preflightExchangeState = Get-ExchangeConnectionState
    $exchangeConnectedByScript = Connect-ExchangeService -State $preflightExchangeState

    # Get-Mailbox and Get-MailboxSearch are generated by the Exchange session and
    # are not available until after Connect-ExchangeOnline succeeds.
    Assert-CommandAvailable -CommandName 'Get-Mailbox' -ModuleName 'ExchangeOnlineManagement'
    Assert-CommandAvailable -CommandName 'Get-MailboxSearch' -ModuleName 'ExchangeOnlineManagement'

    Write-Host "`n[*] Retrieving mailboxes with Litigation Hold enabled..." -ForegroundColor Cyan
    $litigationHoldMailboxes = @(
        Get-Mailbox -ResultSize Unlimited -Filter { LitigationHoldEnabled -eq $true } |
            Select-Object DisplayName, UserPrincipalName, LitigationHoldEnabled,
                LitigationHoldDate, LitigationHoldOwner
    )

    Write-Host "[*] Retrieving active In-Place Holds..." -ForegroundColor Cyan
    $allInPlaceHolds = @(
        Get-MailboxSearch -ResultSize Unlimited |
            Where-Object { $_.InPlaceHoldEnabled -eq $true }
    )

    Write-Host "[*] Retrieving mailbox list for In-Place Hold source resolution..." -ForegroundColor Cyan
    $allMailboxes = @(
        Get-Mailbox -ResultSize Unlimited |
            Select-Object DisplayName, UserPrincipalName
    )

    Write-Host "    Litigation Hold mailboxes : $($litigationHoldMailboxes.Count)"
    Write-Host "    Active In-Place Holds     : $($allInPlaceHolds.Count)"
    Write-Host "    Total mailboxes            : $($allMailboxes.Count)"

    $upnsToCheck = [System.Collections.Generic.HashSet[string]]::new(
        [System.StringComparer]::OrdinalIgnoreCase
    )

    foreach ($mailbox in $litigationHoldMailboxes) {
        if (-not [string]::IsNullOrWhiteSpace($mailbox.UserPrincipalName)) {
            [void]$upnsToCheck.Add($mailbox.UserPrincipalName)
        }
    }

    foreach ($hold in $allInPlaceHolds) {
        $sources = @($hold.SourceMailboxes)

        if ($sources.Count -eq 0) {
            Write-Verbose "In-Place Hold '$($hold.Name)' is org-wide; adding all mailboxes for license validation."
            foreach ($mailbox in $allMailboxes) {
                [void]$upnsToCheck.Add($mailbox.UserPrincipalName)
            }
            continue
        }

        foreach ($source in $sources) {
            $matchedMailbox = $allMailboxes |
                Where-Object {
                    $_.UserPrincipalName -ieq $source -or $_.DisplayName -ieq $source
                } |
                Select-Object -First 1

            if ($matchedMailbox) {
                [void]$upnsToCheck.Add($matchedMailbox.UserPrincipalName)
            }
            else {
                Write-Warning "Could not resolve In-Place Hold source '$source' in hold '$($hold.Name)'."
            }
        }
    }

    $graphUserCache = @{}
    $unlicensedUpns = [System.Collections.Generic.HashSet[string]]::new(
        [System.StringComparer]::OrdinalIgnoreCase
    )

    $upnList = @($upnsToCheck)
    Write-Host "[*] Checking M365 license status for $($upnList.Count) mailbox(es)..." -ForegroundColor Cyan

    $counter = 0
    foreach ($upn in $upnList) {
        $counter++
        $percent = if ($upnList.Count -gt 0) {
            [math]::Round(($counter / $upnList.Count) * 100)
        }
        else {
            100
        }

        Write-Progress -Id 1 `
            -Activity 'Checking Microsoft 365 license status' `
            -Status "[$counter / $($upnList.Count)] $upn" `
            -PercentComplete $percent

        try {
            $graphUser = Get-MgUser -UserId $upn `
                -Property 'assignedLicenses,displayName,userPrincipalName,id,accountEnabled' `
                -ErrorAction Stop

            $graphUserCache[$upn] = $graphUser
            $licenseCount = @($graphUser.AssignedLicenses).Count

            if ($licenseCount -eq 0) {
                [void]$unlicensedUpns.Add($upn)
            }
        }
        catch {
            Write-Warning "Could not retrieve Microsoft Graph user data for '$upn': $($_.Exception.Message)"
        }
    }

    Write-Progress -Id 1 -Activity 'Checking Microsoft 365 license status' -Completed

    $workList = @($unlicensedUpns)
    Write-Host "[*] Processing $($workList.Count) unlicensed mailbox(es) with hold(s)..." -ForegroundColor Cyan

    $counter = 0
    foreach ($upn in $workList) {
        $counter++
        $percent = if ($workList.Count -gt 0) {
            [math]::Round(($counter / $workList.Count) * 100)
        }
        else {
            100
        }

        Write-Progress -Id 2 `
            -Activity 'Processing unlicensed mailboxes' `
            -Status "[$counter / $($workList.Count)] $upn" `
            -PercentComplete $percent

        $graphUser = $graphUserCache[$upn]
        $litigationRow = $litigationHoldMailboxes |
            Where-Object { $_.UserPrincipalName -ieq $upn } |
            Select-Object -First 1

        $matchingInPlaceHolds = @(
            Get-InPlaceHoldsForMailbox -UserPrincipalName $upn -AllInPlaceHolds $allInPlaceHolds
        )

        $hasLitigationHold = $null -ne $litigationRow
        $hasInPlaceHold = $matchingInPlaceHolds.Count -gt 0

        $litigationAction = 'N/A'
        $litigationStatus = 'N/A'

        if ($hasLitigationHold) {
            if (-not $DisableLitigationHold) {
                $litigationAction = 'ReportOnly'
                $litigationStatus = 'No change made'
            }
            elseif ($PSCmdlet.ShouldProcess($upn, 'Disable Litigation Hold')) {
                try {
                    Set-Mailbox -Identity $upn -LitigationHoldEnabled $false -ErrorAction Stop
                    $litigationAction = 'LitigationHoldDisabled'
                    $litigationStatus = 'Success'
                }
                catch {
                    $litigationAction = 'LitigationHoldDisabled'
                    $litigationStatus = "Failed: $($_.Exception.Message)"
                }
            }
            else {
                $litigationAction = 'LitigationHoldDisabled'
                $litigationStatus = 'WhatIf-WouldDisable'
            }
        }

        $inPlaceActions = @()
        foreach ($hold in $matchingInPlaceHolds) {
            if (-not $DisableInPlaceHold) {
                $inPlaceActions += "$($hold.Name)=ReportOnly"
                continue
            }

            $holdResult = Remove-MailboxFromInPlaceHold -UserPrincipalName $upn -Hold $hold
            $inPlaceActions += "$($holdResult.HoldName)=$($holdResult.HoldStatus)"
        }

        $results.Add([PSCustomObject]@{
            DisplayName          = $graphUser.DisplayName
            UserPrincipalName    = $upn
            AccountEnabled       = $graphUser.AccountEnabled
            LicenseCount         = 0
            HasLitigationHold    = $hasLitigationHold
            LitigationHoldDate   = if ($litigationRow) { $litigationRow.LitigationHoldDate } else { $null }
            LitigationHoldOwner  = if ($litigationRow) { $litigationRow.LitigationHoldOwner } else { $null }
            LitigationHoldAction = $litigationAction
            LitigationHoldStatus = $litigationStatus
            HasInPlaceHold       = $hasInPlaceHold
            InPlaceHoldCount     = $matchingInPlaceHolds.Count
            InPlaceHoldNames     = ($matchingInPlaceHolds.Name -join '; ')
            InPlaceHoldActions   = ($inPlaceActions -join ' | ')
        })
    }

    Write-Progress -Id 2 -Activity 'Processing unlicensed mailboxes' -Completed

    $mode = switch ($true) {
        ($WhatIfPreference -and $DisableLitigationHold -and $DisableInPlaceHold) {
            'WhatIf — all holds would be remediated'
        }
        ($WhatIfPreference -and $DisableLitigationHold) {
            'WhatIf — Litigation Hold remediation only'
        }
        ($WhatIfPreference -and $DisableInPlaceHold) {
            'WhatIf — In-Place Hold remediation only'
        }
        ($WhatIfPreference) {
            'WhatIf — no changes made'
        }
        ($DisableLitigationHold -and $DisableInPlaceHold) {
            'Litigation and In-Place Hold remediation'
        }
        $DisableLitigationHold {
            'Litigation Hold remediation only'
        }
        $DisableInPlaceHold {
            'In-Place Hold remediation only'
        }
        default {
            'Report only'
        }
    }

    Write-Host "`n========== RESULTS ==========" -ForegroundColor Cyan
    Write-Host "[*] Mode      : $mode"
    Write-Host "[*] Mailboxes : $($results.Count)"

    if ($results.Count -gt 0) {
        $results | Format-Table DisplayName, UserPrincipalName, HasLitigationHold,
            HasInPlaceHold, InPlaceHoldCount, LitigationHoldStatus,
            InPlaceHoldActions -AutoSize
    }
    else {
        Write-Host "[+] No unlicensed mailboxes with holds were found." -ForegroundColor Green
    }

    if (-not [string]::IsNullOrWhiteSpace($effectiveCsv)) {
        $results | Export-Csv -Path $effectiveCsv -NoTypeInformation -Encoding utf8
        Write-Host "[+] CSV report created: $effectiveCsv" -ForegroundColor Green
    }

    if ($sendReport) {
        try {
            Send-HoldReport `
                -CsvPath $effectiveCsv `
                -To $ReportTo `
                -From $ReportFrom `
                -Subject $ReportSubject `
                -Mode $mode `
                -ResultCount $results.Count

            Write-Host "[+] Report emailed to: $($ReportTo -join ', ')" -ForegroundColor Green
        }
        catch {
            Write-Warning "Unable to send report email: $($_.Exception.Message)"
        }
    }

    return $results
}
finally {
    if ($temporaryCsv -and $effectiveCsv -and (Test-Path -LiteralPath $effectiveCsv)) {
        Remove-Item -LiteralPath $effectiveCsv -Force -ErrorAction SilentlyContinue
        Write-Verbose "Temporary CSV removed: $effectiveCsv"
    }

    Disconnect-ScriptServices `
        -GraphConnectedByScript $graphConnectedByScript `
        -ExchangeConnectedByScript $exchangeConnectedByScript

    Write-Verbose "=== Script completed: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ==="
}
