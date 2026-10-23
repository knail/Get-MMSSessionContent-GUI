<#
.SYNOPSIS
    GUI for discovering, filtering, and downloading Midwest Management Summit (MMS)
    session content from Sched.

.DESCRIPTION
    This script provides a graphical user interface for discovering,
    selecting, filtering, and downloading Midwest Management Summit (MMS)
    session content from Sched.

    This project is based on Get-MMSSessionContent.ps1 Version 1.7.4 from
    the CopyMMSFiles project and would not exist without the work of the
    original authors and contributors listed below.

    Original Project:
        Andrew Johnson
        https://github.com/AndrewJNet/CopyMMSFiles

    Original Contributors:

        Duncan Russell
        http://www.sysadmintechnotes.com

        Andrew Johnson
        https://github.com/AndrewJNet

        Evan Yeung
        https://www.forevanyeung.com

        Chris Kibble
        https://www.christopherkibble.com

        Jon Warnken
        https://www.mrbodean.net

        Oliver Baddeley
        Desert Edition Contributor

        Benjamin Reynolds
        https://sqlbenjamin.wordpress.com

        Jorge Suarez
        https://github.com/jorgeasaurus

        Nathan Ziehnert
        https://z-nerd.com

        Piotr Gardy
        https://garit.pro

    GUI Edition:

        Keith Nail
        Marion, North Carolina
        https://www.my02dcents.com

    GUI Enhancements Include:

        - Windows Forms graphical interface
        - Automatic MMS conference discovery
        - Conference accessibility validation
        - Conference and session filtering
        - Session selection with checkbox support
        - Bulk conference downloads
        - Individual session downloads
        - Credential prompt with optional anonymous mode
        - Download logging
        - Results reporting
        - Excel export support with CSV fallback
        - Preloaded conference and session inventory
        - Automatic STA relaunch when required

    Run with NO parameters:

        1. A sign-in window prompts for Sched credentials.
        2. Credentials may be skipped for public-only content.
        3. The script discovers known and available MMS conferences.
        4. Accessible conferences are loaded into the GUI.
        5. Sessions are retrieved and displayed.
        6. Users can:
             - Download all sessions from selected conferences
             - Download selected sessions only
             - Review session metadata before downloading
        7. Logs and reports are written to C:\Scripts.

    Requirements:

        - Windows
        - PowerShell 5.1 or PowerShell 7.x
        - Internet connectivity
        - Valid Sched credentials for protected content

.INPUTS
    None

.OUTPUTS

    Session Content:
        <DownloadLocation>\MMSContent\

    Default Download Location:
        C:\Scripts\MMSContent

    Log File:
        C:\Scripts\yyyy-MM-dd_HHmm-MMS-Download-Log.txt

    Results Export:
        C:\Scripts\yyyy-MM-dd_HHmm-MMS-Download-Results.xlsx

    CSV Fallback:
        C:\Scripts\yyyy-MM-dd_HHmm-MMS-Download-Results.csv

.NOTES

    Script Name:           Get-MMSSessionContent-GUI.ps1
    Version:               2.1.1-GUI
    Author:                Keith Nail
    Website:               https://www.my02dcents.com
    Created:               2026-10-23

    Based On:
        Get-MMSSessionContent.ps1 Version 1.7.4

    Original Project:
        https://github.com/AndrewJNet/CopyMMSFiles

    Purpose:
        Provide a graphical front-end for discovering,
        browsing, filtering, and downloading MMS session
        content without requiring command-line parameters.

    Version History

    Date          Version       Author          Notes
    ----------    -----------   -------------   -----------------------------------
    2026-10-23    2.1.1-GUI     Keith Nail      Initial GUI release based on
                                                Get-MMSSessionContent.ps1 v1.7.4

.EXAMPLE

    .\Get-MMSSessionContent-GUI.ps1

    Launches the GUI and prompts for Sched credentials.

.EXAMPLE

    powershell.exe -ExecutionPolicy Bypass `
        -File .\Get-MMSSessionContent-GUI.ps1

    Launches the GUI directly from PowerShell.

.LINK

    GUI Project Author
    https://www.my02dcents.com

.LINK

    Original CopyMMSFiles Project
    https://github.com/AndrewJNet/CopyMMSFiles

.LINK

    MMS Conference Website
    https://mmsmoa.sched.com

#>

# ==================================================
# MARK: Prerequisites and Assemblies
# ==================================================
if ($PSVersionTable.PSObject.Properties.Name -contains 'Platform' -and $PSVersionTable.Platform -and $PSVersionTable.Platform -ne 'Win32NT') {
    Write-Host 'This GUI requires Windows (Windows Forms).' -ForegroundColor Red
    return
}

# Windows Forms needs an STA thread. Relaunch automatically if required.
if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne [System.Threading.ApartmentState]::STA) {
    if ($PSCommandPath) {
        $HostExe = (Get-Process -Id $PID).Path
        Start-Process -FilePath $HostExe -ArgumentList @('-STA', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$PSCommandPath`"")
    }
    else {
        Write-Host 'Please run this script from a saved .ps1 file or start PowerShell with the -STA switch.' -ForegroundColor Red
    }
    return
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Web
[System.Windows.Forms.Application]::EnableVisualStyles()

try { [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12 } catch { }

## Hide Invoke-WebRequest progress bar. There's a bug that doesn't clear the bar after a request is finished.
$ProgressPreference = 'SilentlyContinue'
# ==================================================
# ^^^^^ Prerequisites and Assemblies ^^^^^
# ==================================================

# ==================================================
# MARK: Global Settings and State
# ==================================================
$Script:OutputRoot = 'C:\Scripts'
$Script:DefaultDownloadLocation = Join-Path -Path $Script:OutputRoot -ChildPath 'MMSContent'
$Script:RunStamp = Get-Date -Format 'yyyy-MM-dd_HHmm'
$Script:LogFile = Join-Path -Path $Script:OutputRoot -ChildPath ($Script:RunStamp + '-MMS-Download-Log.txt')

if (-not (Test-Path -Path $Script:OutputRoot)) {
    New-Item -ItemType Directory -Force -Path $Script:OutputRoot | Out-Null
}

$Script:RegexIgnoreCase = [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
$Script:NewLine = [System.Environment]::NewLine

$Script:Sessions = @{}          # ConferenceId -> WebRequestSession (authenticated or anonymous)
$Script:EventCache = @{}        # ConferenceId -> array of session (event) objects
$Script:SessionList = New-Object -TypeName System.Collections.Generic.List[object]   # All sessions from accessible conferences
$Script:Results = New-Object -TypeName System.Collections.Generic.List[object]       # Download results for export
$Script:Credential = $null
$Script:CancelRequested = $false
$Script:IsBusy = $false
$Script:SuppressEvents = $false
# ==================================================
# ^^^^^ Global Settings and State ^^^^^
# ==================================================

# ==================================================
# MARK: Conference Catalog
# ==================================================
function New-ConferenceEntry {
    param(
        [Parameter(Mandatory = $true)][string]$Id,
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][string]$Type
    )

    return [PSCustomObject]@{
        Id           = $Id
        Name         = $Name
        Type         = $Type          # Public / Private / Discovered
        Access       = 'Not queried'
        SessionCount = 0
        Checked      = $false
    }
}

# Known conferences (from $PublicContentYears / $PrivateContentYears in the original script)
$Script:ConferenceCatalog = New-Object -TypeName System.Collections.Generic.List[object]
$Script:ConferenceCatalog.Add((New-ConferenceEntry -Id '2015'      -Name 'MMS 2015'                        -Type 'Public'))
$Script:ConferenceCatalog.Add((New-ConferenceEntry -Id '2016'      -Name 'MMS 2016'                        -Type 'Public'))
$Script:ConferenceCatalog.Add((New-ConferenceEntry -Id '2017'      -Name 'MMS 2017'                        -Type 'Public'))
$Script:ConferenceCatalog.Add((New-ConferenceEntry -Id '2018'      -Name 'MMS 2018'                        -Type 'Private'))
$Script:ConferenceCatalog.Add((New-ConferenceEntry -Id 'de2018'    -Name 'MMS 2018 Desert Edition'         -Type 'Private'))
$Script:ConferenceCatalog.Add((New-ConferenceEntry -Id '2019'      -Name 'MMS 2019'                        -Type 'Public'))
$Script:ConferenceCatalog.Add((New-ConferenceEntry -Id 'jazz'      -Name 'MMS Jazz Edition'                -Type 'Public'))
$Script:ConferenceCatalog.Add((New-ConferenceEntry -Id 'miami'     -Name 'MMS Miami Edition'               -Type 'Public'))
$Script:ConferenceCatalog.Add((New-ConferenceEntry -Id '2022atmoa' -Name 'MMS 2022 at MOA'                 -Type 'Public'))
$Script:ConferenceCatalog.Add((New-ConferenceEntry -Id '2023atmoa' -Name 'MMS 2023 at MOA'                 -Type 'Public'))
$Script:ConferenceCatalog.Add((New-ConferenceEntry -Id '2023miami' -Name 'MMS 2023 Miami Edition'          -Type 'Public'))
$Script:ConferenceCatalog.Add((New-ConferenceEntry -Id '2024atmoa' -Name 'MMS 2024 at MOA'                 -Type 'Public'))
$Script:ConferenceCatalog.Add((New-ConferenceEntry -Id '2024fll'   -Name 'MMS 2024 Flamingo Edition (FLL)' -Type 'Public'))
$Script:ConferenceCatalog.Add((New-ConferenceEntry -Id '2025atmoa' -Name 'MMS 2025 at MOA'                 -Type 'Public'))
$Script:ConferenceCatalog.Add((New-ConferenceEntry -Id '2025music' -Name 'MMS 2025 Music City Edition'     -Type 'Public'))
$Script:ConferenceCatalog.Add((New-ConferenceEntry -Id '2026atmoa' -Name 'MMS 2026 at MOA'                 -Type 'Private'))

# Suffix patterns used to probe Sched for MMS conferences not in the list above
$Script:ProbeSuffixes = @('atmoa', 'miami', 'fll', 'music')
$Script:ProbeStartYear = 2022

function Get-ConferenceInfo {
    param([Parameter(Mandatory = $true)][string]$ConfId)

    foreach ($Conf in $Script:ConferenceCatalog) {
        if ($Conf.Id -ieq $ConfId) { return $Conf }
    }
    return $null
}
# ==================================================
# ^^^^^ Conference Catalog ^^^^^
# ==================================================

# ==================================================
# MARK: Logging Functions
# ==================================================
function Write-Log {
    param(
        [Parameter(Mandatory = $true)][string]$Message,
        [string]$Level = 'INFO'
    )

    $Line = '[{0}] [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd_HHmmss'), $Level, $Message
    try { Add-Content -Path $Script:LogFile -Value $Line -ErrorAction Stop } catch { }

    if ($Script:TxtLog -and -not $Script:TxtLog.IsDisposed) {
        $Script:TxtLog.AppendText($Line + $Script:NewLine)
    }
    [System.Windows.Forms.Application]::DoEvents()
}

function Set-Status {
    param([string]$Text)

    if ($Script:LblStatus -and -not $Script:LblStatus.IsDisposed) {
        $Script:LblStatus.Text = $Text
    }
    [System.Windows.Forms.Application]::DoEvents()
}

function Add-DownloadResult {
    param(
        [string]$ConfId,
        [string]$SessionTitle,
        [string]$FileName,
        [string]$Status,
        [string]$Path
    )

    $Script:Results.Add([PSCustomObject]@{
            Timestamp  = Get-Date -Format 'yyyy-MM-dd_HHmmss'
            Conference = $ConfId
            Session    = $SessionTitle
            File       = $FileName
            Status     = $Status
            Path       = $Path
        })
}
# ==================================================
# ^^^^^ Logging Functions ^^^^^
# ==================================================

# ==================================================
# MARK: Helper Functions
# ==================================================
function Wait-WithDoEvents {
    param([int]$Seconds = 1)

    $EndTime = (Get-Date).AddSeconds($Seconds)
    while ((Get-Date) -lt $EndTime -and -not $Script:CancelRequested) {
        Start-Sleep -Milliseconds 200
        [System.Windows.Forms.Application]::DoEvents()
    }
}

function Get-ErrorStatusCode {
    param($ErrorRecord)

    try {
        if ($ErrorRecord.Exception.Response) { return [int]$ErrorRecord.Exception.Response.StatusCode }
    }
    catch { }
    return 0
}

function Test-IsBlank {
    param($Value)
    return ([System.String]::IsNullOrWhiteSpace([string]$Value))
}

# Wrapper around Invoke-WebRequest that adds the 429 (Too Many Requests) wait-and-retry from the original script
function Invoke-SchedRequest {
    param(
        [Parameter(Mandatory = $true)][string]$Uri,
        [Parameter(Mandatory = $true)]$WebSession,
        [string]$Method = 'GET',
        [string]$Body,
        [hashtable]$Headers,
        [string]$OutFile
    )

    $Params = @{
        Uri             = $Uri
        WebSession      = $WebSession
        Method          = $Method
        UseBasicParsing = $true
        ErrorAction     = 'Stop'
    }
    if ($Body) { $Params['Body'] = $Body }
    if ($Headers) { $Params['Headers'] = $Headers }
    if ($OutFile) { $Params['OutFile'] = $OutFile }

    try {
        return (Invoke-WebRequest @Params)
    }
    catch {
        $Err = $_
        if ((Get-ErrorStatusCode -ErrorRecord $Err) -eq 429) {
            Write-Log -Message 'HTTP 429 (Too Many Requests) received. Waiting 20 seconds, then retrying once.' -Level 'WARN'
            Wait-WithDoEvents -Seconds 20
            return (Invoke-WebRequest @Params)
        }
        throw
    }
}

function Invoke-BasicHTMLParser {
    param(
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$Html,
        [Parameter(Mandatory = $true)][string]$SchedBaseURL
    )

    $Html = $Html.Replace('<br>', "`r`n").Replace('<br/>', "`r`n").Replace('<br />', "`r`n") # replace <br> with new line

    # Speaker Spacing
    $Html = $Html.Replace('<div class="sched-person-session">', "`r`n`r`n")

    # Link parsing
    $LinkRegex = '(?<texttoreplace><a.*?href="(?<link>.*?)".*?>(?<content>.*?)<\/a>)'
    $LinkMatches = [System.Text.RegularExpressions.Regex]::Matches($Html, $LinkRegex, $Script:RegexIgnoreCase)
    foreach ($LinkMatch in $LinkMatches) {
        $LinkTarget = $LinkMatch.Groups['link'].Value
        if (-not $LinkTarget.StartsWith('http')) { $LinkValue = $SchedBaseURL + '/' + $LinkTarget } else { $LinkValue = $LinkTarget }
        $OpenSquare = [System.Char]91
        $CloseSquare = [System.Char]93
        $ReplaceWith = ' ' + $OpenSquare + $LinkMatch.Groups['content'].Value + $CloseSquare + '(' + $LinkValue + ')'
        $Html = $Html.Replace($LinkMatch.Groups['texttoreplace'].Value, $ReplaceWith)
    }

    # List Parsing
    $ListRegex = '(?<texttoreplace><ul[^>]?>(?<content>.*?)<\/ul>)'
    $ListMatches = [System.Text.RegularExpressions.Regex]::Matches($Html, $ListRegex, $Script:RegexIgnoreCase)
    foreach ($ListMatch in $ListMatches) {
        $ListContent = $ListMatch.Groups['content'].Value.Replace('<li>', "`r`n* ").Replace('</li>', '')
        $Html = $Html.Replace($ListMatch.Groups['texttoreplace'].Value, $ListContent)
    }

    # General Cleanup
    $Html = $Html.Replace('&rarr;', '')
    $Html = $Html -replace '<div[^>]+>', "`r`n"
    $Html = $Html -replace '<[^>]+>', '' # Strip all HTML tags

    return $Html
}

function Get-EventLinkPairs {
    param(
        [Parameter(Mandatory = $true)]
        [object]$SchedulePage
    )

    $EventPairs = New-Object -TypeName System.Collections.Generic.List[object]
    $EventTitleByUrl = @{}

    if ($SchedulePage.Content) {
        $EventRegex = '<a[^>]+href="(?<url>event\/[^"#?]+)"[^>]*>(?<title>.*?)<\/a>'
        $FlatContent = $SchedulePage.Content.Replace("`r", '').Replace("`n", '')
        $EventMatches = [System.Text.RegularExpressions.Regex]::Matches($FlatContent, $EventRegex, $Script:RegexIgnoreCase)
        foreach ($EventMatch in $EventMatches) {
            $EventUrlValue = ([string]$EventMatch.Groups['url'].Value).Trim()
            $EventTitleValue = ($EventMatch.Groups['title'].Value -replace '<[^>]+>', '').Trim()
            if (-not (Test-IsBlank $EventUrlValue) -and -not (Test-IsBlank $EventTitleValue) -and -not $EventTitleByUrl.ContainsKey($EventUrlValue)) {
                $EventTitleByUrl[$EventUrlValue] = $EventTitleValue
            }
        }
    }

    # First attempt: use parsed Links collection
    if ($SchedulePage.Links) {
        foreach ($Lnk in $SchedulePage.Links) {
            if ($Lnk.href -like 'event/*') {
                $LnkUrl = [string]$Lnk.href
                $LnkTitle = [string]$Lnk.innerText

                if ((Test-IsBlank $LnkTitle) -and $Lnk.outerHTML) {
                    $LnkTitle = ([string]$Lnk.outerHTML -replace '<[^>]+>', '').Trim()
                }

                if ((Test-IsBlank $LnkTitle) -and $EventTitleByUrl.ContainsKey($LnkUrl)) {
                    $LnkTitle = [string]$EventTitleByUrl[$LnkUrl]
                }

                if ((Test-IsBlank $LnkTitle) -and $LnkUrl -match '^event\/[^\/]+\/(?<slug>[^\/?#]+)') {
                    $LnkTitle = [System.Uri]::UnescapeDataString($Matches['slug']) -replace '-', ' '
                }

                $EventPairs.Add([PSCustomObject]@{
                        Url   = $LnkUrl
                        Title = $LnkTitle
                    })
            }
        }
    }

    # Fallback: parse anchor tags from raw HTML when Links parsing is incomplete
    if ($EventPairs.Count -eq 0 -and $SchedulePage.Content) {
        foreach ($Entry in $EventTitleByUrl.GetEnumerator()) {
            $EventPairs.Add([PSCustomObject]@{
                    Url   = [string]$Entry.Key
                    Title = [string]$Entry.Value
                })
        }
    }

    # De-duplicate by URL while preserving order
    $Seen = @{}
    $UniquePairs = New-Object -TypeName System.Collections.Generic.List[object]
    foreach ($Pair in $EventPairs) {
        if (Test-IsBlank $Pair.Url) { continue }
        if (-not $Seen.ContainsKey($Pair.Url)) {
            $Seen[$Pair.Url] = $true
            $UniquePairs.Add($Pair)
        }
    }

    return $UniquePairs
}

function Get-CurrentCredential {
    if ((Test-IsBlank $Script:TxtUser.Text) -or [System.String]::IsNullOrEmpty($Script:TxtPass.Text)) { return $null }
    $SecurePass = ConvertTo-SecureString -String $Script:TxtPass.Text -AsPlainText -Force
    return (New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $Script:TxtUser.Text.Trim(), $SecurePass)
}
# ==================================================
# ^^^^^ Helper Functions ^^^^^
# ==================================================

# ==================================================
# MARK: Credential Prompt
# ==================================================
# Startup sign-in window. Returns a PSCredential, or $null if the user chooses public content only.
function Show-CredentialPrompt {
    $Dlg = New-Object -TypeName System.Windows.Forms.Form
    $Dlg.Text = 'MMS Sched Sign-in'
    $Dlg.Size = New-Object -TypeName System.Drawing.Size -ArgumentList 430, 250
    $Dlg.StartPosition = 'CenterScreen'
    $Dlg.FormBorderStyle = 'FixedDialog'
    $Dlg.MaximizeBox = $false
    $Dlg.MinimizeBox = $false
    $Dlg.TopMost = $true
    $Dlg.Font = New-Object -TypeName System.Drawing.Font -ArgumentList 'Segoe UI', 9

    $LblInfo = New-Object -TypeName System.Windows.Forms.Label
    $LblInfo.Text = 'Enter your Sched credentials. The script will query all MMS conferences and list the ones you can access. Choose Skip for public content only.'
    $LblInfo.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 15, 12
    $LblInfo.Size = New-Object -TypeName System.Drawing.Size -ArgumentList 390, 50
    $Dlg.Controls.Add($LblInfo)

    $LblDlgUser = New-Object -TypeName System.Windows.Forms.Label
    $LblDlgUser.Text = 'Username:'
    $LblDlgUser.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 15, 75
    $LblDlgUser.AutoSize = $true
    $Dlg.Controls.Add($LblDlgUser)

    $TxtDlgUser = New-Object -TypeName System.Windows.Forms.TextBox
    $TxtDlgUser.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 95, 72
    $TxtDlgUser.Size = New-Object -TypeName System.Drawing.Size -ArgumentList 300, 24
    $Dlg.Controls.Add($TxtDlgUser)

    $LblDlgPass = New-Object -TypeName System.Windows.Forms.Label
    $LblDlgPass.Text = 'Password:'
    $LblDlgPass.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 15, 110
    $LblDlgPass.AutoSize = $true
    $Dlg.Controls.Add($LblDlgPass)

    $TxtDlgPass = New-Object -TypeName System.Windows.Forms.TextBox
    $TxtDlgPass.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 95, 107
    $TxtDlgPass.Size = New-Object -TypeName System.Drawing.Size -ArgumentList 300, 24
    $TxtDlgPass.UseSystemPasswordChar = $true
    $Dlg.Controls.Add($TxtDlgPass)

    $BtnDlgOk = New-Object -TypeName System.Windows.Forms.Button
    $BtnDlgOk.Text = 'Sign in and Query'
    $BtnDlgOk.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 95, 150
    $BtnDlgOk.Size = New-Object -TypeName System.Drawing.Size -ArgumentList 140, 32
    $BtnDlgOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $Dlg.Controls.Add($BtnDlgOk)

    $BtnDlgSkip = New-Object -TypeName System.Windows.Forms.Button
    $BtnDlgSkip.Text = 'Skip (Public Only)'
    $BtnDlgSkip.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 245, 150
    $BtnDlgSkip.Size = New-Object -TypeName System.Drawing.Size -ArgumentList 150, 32
    $BtnDlgSkip.DialogResult = [System.Windows.Forms.DialogResult]::Ignore
    $Dlg.Controls.Add($BtnDlgSkip)

    $Dlg.AcceptButton = $BtnDlgOk
    $Dlg.CancelButton = $BtnDlgSkip
    $Dlg.Add_Shown({ $TxtDlgUser.Focus() | Out-Null })

    $DlgResult = $Dlg.ShowDialog()
    $ReturnCred = $null
    if ($DlgResult -eq [System.Windows.Forms.DialogResult]::OK -and -not (Test-IsBlank $TxtDlgUser.Text) -and -not [System.String]::IsNullOrEmpty($TxtDlgPass.Text)) {
        $SecurePass = ConvertTo-SecureString -String $TxtDlgPass.Text -AsPlainText -Force
        $ReturnCred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $TxtDlgUser.Text.Trim(), $SecurePass
    }
    $Dlg.Dispose()
    return $ReturnCred
}
# ==================================================
# ^^^^^ Credential Prompt ^^^^^
# ==================================================

# ==================================================
# MARK: Authentication and Conference Query Functions
# ==================================================
# Opens a web session for a conference.
# If credentials are supplied, a Sched login is attempted. Non-private conferences fall back to anonymous access.
function Connect-SchedConference {
    param(
        [Parameter(Mandatory = $true)][string]$ConfId,
        [Parameter(Mandatory = $true)][string]$Type,
        [System.Management.Automation.PSCredential]$Credential
    )

    $SchedBaseURL = 'https://mms' + $ConfId + '.sched.com'
    $SchedLoginURL = $SchedBaseURL + '/login'
    $LoginFormPattern = '<input[^>]+name=[''"]login[''"]'

    try {
        $WebSess = New-Object -TypeName Microsoft.PowerShell.Commands.WebRequestSession
        $Web = Invoke-SchedRequest -Uri $SchedLoginURL -WebSession $WebSess

        if (-not $Credential) {
            if ($Type -eq 'Private') {
                return [PSCustomObject]@{ Success = $false; Mode = 'Failed'; Message = 'Login required' }
            }
            $Script:Sessions[$ConfId] = $WebSess
            return [PSCustomObject]@{ Success = $true; Mode = 'Public'; Message = 'Anonymous session opened' }
        }

        $UserName = $Credential.UserName
        $Password = $Credential.GetNetworkCredential().Password

        # Updated POST body (v1.3 / v1.7.4 of the original script)
        $Body = 'landing_conf=' + [System.Uri]::EscapeDataString($SchedBaseURL) + '&username=' + [System.Uri]::EscapeDataString($UserName) + '&password=' + [System.Uri]::EscapeDataString($Password) + '&login='
        $Password = $null

        $Web = Invoke-SchedRequest -Uri $SchedLoginURL -WebSession $WebSess -Method 'POST' -Body $Body -Headers @{ Referer = $SchedLoginURL; Origin = $SchedBaseURL }
        $Body = $null

        # If the login form is still present the login did not work
        if ($Web.Content -match $LoginFormPattern) {
            if ($Type -ne 'Private') {
                $WebSess = New-Object -TypeName Microsoft.PowerShell.Commands.WebRequestSession
                $Web = Invoke-SchedRequest -Uri $SchedLoginURL -WebSession $WebSess
                $Script:Sessions[$ConfId] = $WebSess
                return [PSCustomObject]@{ Success = $true; Mode = 'Public'; Message = 'Login not accepted, using anonymous session' }
            }
            return [PSCustomObject]@{ Success = $false; Mode = 'Failed'; Message = 'Login failed / no access' }
        }

        $Script:Sessions[$ConfId] = $WebSess
        return [PSCustomObject]@{ Success = $true; Mode = 'Authenticated'; Message = 'Authenticated' }
    }
    catch {
        $Err = $_
        return [PSCustomObject]@{ Success = $false; Mode = 'Failed'; Message = 'Unavailable: ' + $Err.Exception.Message }
    }
}

# Retrieves (and caches) the list of sessions for a conference using the existing web session
function Get-ConferenceEvents {
    param(
        [Parameter(Mandatory = $true)][string]$ConfId,
        [switch]$Refresh
    )

    if (-not $Refresh -and $Script:EventCache.ContainsKey($ConfId)) { return @($Script:EventCache[$ConfId]) }

    $SchedBaseURL = 'https://mms' + $ConfId + '.sched.com'
    $WebSess = $Script:Sessions[$ConfId]

    $Sched = Invoke-SchedRequest -Uri ($SchedBaseURL + '/list/descriptions') -WebSession $WebSess

    $EventList = New-Object -TypeName System.Collections.Generic.List[object]
    $Pairs = Get-EventLinkPairs -SchedulePage $Sched

    foreach ($Pair in $Pairs) {
        $EventTitle = [string]$Pair.Title
        $EventUrl = [string]$Pair.Url
        if (Test-IsBlank $EventUrl) { continue }
        if (Test-IsBlank $EventTitle) { $EventTitle = $EventUrl }

        # Strip capacity labels (FULL, LIMITED, FILLING, etc.) that Sched appends to the anchor text
        $EventTitle = $EventTitle -replace '\s*(FULL|LIMITED|SOLD OUT|WAITLIST|FILLING)\s*$', ''
        $EventTitle = $EventTitle.Trim()

        # Skip social/networking/activity events that never have downloadable content
        if ($EventTitle -match '^Fishing with' -or $EventTitle -match '^Career Connections' -or
            $EventTitle -match '^Great Big Game Show' -or $EventTitle -match 'Escape Game') { continue }

        $EventList.Add([PSCustomObject]@{
                ConfId  = $ConfId
                Title   = $EventTitle
                Url     = $EventUrl
                Checked = $false
            })
    }

    $Script:EventCache[$ConfId] = $EventList.ToArray()
    return @($Script:EventCache[$ConfId])
}

# Tests one conference: connects, reads the schedule, updates Access / SessionCount, and loads its sessions
function Test-ConferenceAccess {
    param(
        [Parameter(Mandatory = $true)][object]$Conf,
        [System.Management.Automation.PSCredential]$Credential
    )

    $Conf.Access = 'Checking...'
    $Conf.SessionCount = 0

    $Script:Sessions.Remove($Conf.Id)
    $Script:EventCache.Remove($Conf.Id)

    $Connect = Connect-SchedConference -ConfId $Conf.Id -Type $Conf.Type -Credential $Credential
    if (-not $Connect.Success) {
        $Conf.Access = $Connect.Message
        Write-Log -Message ($Conf.Id + ': ' + $Connect.Message) -Level 'WARN'
        return
    }

    try {
        $Events = @(Get-ConferenceEvents -ConfId $Conf.Id -Refresh)
        $Conf.SessionCount = $Events.Count
        if ($Events.Count -gt 0) {
            if ($Connect.Mode -eq 'Authenticated') { $Conf.Access = 'Yes (Logged in)' } else { $Conf.Access = 'Yes (Public)' }
            foreach ($EventItem in $Events) { $Script:SessionList.Add($EventItem) }
        }
        else {
            $Conf.Access = 'No sessions found'
        }
        Write-Log -Message ('{0}: {1} - {2} session(s)' -f $Conf.Id, $Conf.Access, $Events.Count)
    }
    catch {
        $Err = $_
        $Conf.Access = 'Unable to read schedule'
        Write-Log -Message ($Conf.Id + ': unable to retrieve session list - ' + $Err.Exception.Message) -Level 'WARN'
    }
}

# Returns $true if mms<ConfId>.sched.com is a real Sched site (not a redirect to the Sched home page)
function Test-ConferenceExists {
    param([Parameter(Mandatory = $true)][string]$ConfId)

    $ExpectedHost = 'mms' + $ConfId + '.sched.com'
    $Url = 'https://' + $ExpectedHost + '/list/descriptions'
    try {
        $Resp = Invoke-WebRequest -Uri $Url -UseBasicParsing -TimeoutSec 15 -ErrorAction Stop
        $FinalHost = ''
        if ($Resp.BaseResponse.ResponseUri) { $FinalHost = [string]$Resp.BaseResponse.ResponseUri.Host }
        elseif ($Resp.BaseResponse.RequestMessage.RequestUri) { $FinalHost = [string]$Resp.BaseResponse.RequestMessage.RequestUri.Host }
        return ($FinalHost -ieq $ExpectedHost)
    }
    catch {
        return $false
    }
}

# Probes Sched for MMS conference IDs that are not in the known catalog
function Find-NewConferences {
    $LastYear = (Get-Date).Year + 1
    for ($Year = $Script:ProbeStartYear; $Year -le $LastYear; $Year++) {
        foreach ($Suffix in $Script:ProbeSuffixes) {
            if ($Script:CancelRequested) { return }
            $ProbeId = [string]$Year + $Suffix
            if (Get-ConferenceInfo -ConfId $ProbeId) { continue }

            Set-Status -Text ('Probing Sched for conference ID ' + $ProbeId + ' ...')
            if (Test-ConferenceExists -ConfId $ProbeId) {
                $Script:ConferenceCatalog.Add((New-ConferenceEntry -Id $ProbeId -Name ('MMS ' + $ProbeId + ' (discovered)') -Type 'Discovered'))
                Write-Log -Message ('Discovered conference ID: ' + $ProbeId)
            }
        }
    }
}

# Full discovery: probe for new IDs, then test every conference and load sessions for accessible ones
function Invoke-ConferenceDiscovery {
    if ($Script:IsBusy) { return }

    $Script:Credential = Get-CurrentCredential
    $Script:CancelRequested = $false
    Set-UiBusy -Busy $true

    $Script:SessionList.Clear()
    Update-SessionListView

    try {
        if ($Script:Credential) { Write-Log -Message ('Querying MMS conferences as ' + $Script:Credential.UserName + ' ...') }
        else { Write-Log -Message 'Querying MMS conferences (public content only) ...' }

        if ($Script:ChkProbe.Checked) { Find-NewConferences }

        $Total = $Script:ConferenceCatalog.Count
        $Script:PbMain.Minimum = 0
        $Script:PbMain.Maximum = [System.Math]::Max($Total, 1)
        $Script:PbMain.Value = 0

        for ($Index = 0; $Index -lt $Total; $Index++) {
            if ($Script:CancelRequested) { Write-Log -Message 'Query cancelled by user.' -Level 'WARN'; break }
            $Conf = $Script:ConferenceCatalog[$Index]
            Set-Status -Text ('Checking {0} of {1}: {2} ({3})' -f ($Index + 1), $Total, $Conf.Name, $Conf.Id)
            Test-ConferenceAccess -Conf $Conf -Credential $Script:Credential
            Update-ConferenceListView
            $Script:PbMain.Value = $Index + 1
        }

        $Accessible = @($Script:ConferenceCatalog | Where-Object { $_.Access -like 'Yes*' }).Count
        Update-SessionListView
        Set-Status -Text ('Query complete. You can access {0} conference(s) with {1} session(s). Check conferences to filter sessions, then download.' -f $Accessible, $Script:SessionList.Count)
        Write-Log -Message ('Query complete. Accessible conferences: {0} of {1}. Sessions loaded: {2}' -f $Accessible, $Total, $Script:SessionList.Count)
    }
    catch {
        $Err = $_
        Write-Log -Message ('Query error: ' + $Err.Exception.Message) -Level 'ERROR'
    }
    finally {
        Set-UiBusy -Busy $false
    }
}

# Makes sure a web session exists for the conference (reconnects with current credentials if needed)
function Confirm-ConferenceSession {
    param([Parameter(Mandatory = $true)][string]$ConfId)

    if ($Script:Sessions.ContainsKey($ConfId)) { return $true }

    $Conf = Get-ConferenceInfo -ConfId $ConfId
    $ConfType = 'Discovered'
    if ($Conf) { $ConfType = $Conf.Type }

    Write-Log -Message ('Opening session for ' + $ConfId + ' ...')
    $Connect = Connect-SchedConference -ConfId $ConfId -Type $ConfType -Credential $Script:Credential
    if (-not $Connect.Success) {
        Write-Log -Message ('FAILED: login to https://mms' + $ConfId + '.sched.com - ' + $Connect.Message) -Level 'ERROR'
        return $false
    }
    return $true
}
# ==================================================
# ^^^^^ Authentication and Conference Query Functions ^^^^^
# ==================================================

# ==================================================
# MARK: Download Functions
# ==================================================
# Downloads one file. Returns Downloaded / Updated / Skipped / Failed
function Save-SessionFile {
    param(
        [Parameter(Mandatory = $true)][string]$DownloadUrl,
        [Parameter(Mandatory = $true)][string]$OutputFilePath,
        [Parameter(Mandatory = $true)]$WebSession,
        [bool]$ReDownloadIfHashIsDifferent = $false
    )

    $FileLabel = Split-Path -Path $OutputFilePath -Leaf

    if ((Test-Path -LiteralPath $OutputFilePath) -eq $false) {
        Write-Log -Message ('DOWNLOADING: ' + $FileLabel)
        try {
            Invoke-SchedRequest -Uri $DownloadUrl -WebSession $WebSession -OutFile $OutputFilePath | Out-Null
            Unblock-File -LiteralPath $OutputFilePath -ErrorAction SilentlyContinue
            return 'Downloaded'
        }
        catch {
            Write-Log -Message ('FAILED: ' + $FileLabel) -Level 'ERROR'
            if (Test-Path -LiteralPath $OutputFilePath) { Remove-Item -LiteralPath $OutputFilePath -Force -ErrorAction SilentlyContinue }
            return 'Failed'
        }
    }

    if ($ReDownloadIfHashIsDifferent) {
        Write-Log -Message ('DOWNLOADING: ' + $FileLabel + ' (hash check)')
        $OldHash = (Get-FileHash -LiteralPath $OutputFilePath).Hash
        $TempPath = $OutputFilePath + '.new'
        try {
            Invoke-SchedRequest -Uri $DownloadUrl -WebSession $WebSession -OutFile $TempPath | Out-Null
            Unblock-File -LiteralPath $TempPath -ErrorAction SilentlyContinue
            $NewHash = (Get-FileHash -LiteralPath $TempPath).Hash
            if ($NewHash -ne $OldHash) {
                Move-Item -LiteralPath $TempPath -Destination $OutputFilePath -Force
                return 'Updated'
            }
            Write-Log -Message ('SKIPPING: ' + $FileLabel)
            Remove-Item -LiteralPath $TempPath -Force
            return 'Skipped'
        }
        catch {
            Write-Log -Message ('FAILED: ' + $FileLabel) -Level 'ERROR'
            if (Test-Path -LiteralPath $TempPath) { Remove-Item -LiteralPath $TempPath -Force -ErrorAction SilentlyContinue }
            return 'Failed'
        }
    }

    Write-Log -Message ('SKIPPING: ' + $FileLabel)
    return 'Skipped'
}

# Downloads every hosted file for one session (and writes Session Info.txt unless excluded)
function Save-ConferenceSession {
    param(
        [Parameter(Mandatory = $true)][string]$ConfId,
        [Parameter(Mandatory = $true)][object]$EventItem,
        [Parameter(Mandatory = $true)][string]$DownloadLocation,
        [bool]$ExcludeSessionDetails = $false,
        [bool]$ReDownloadIfHashIsDifferent = $false
    )

    $WebSess = $Script:Sessions[$ConfId]
    $SchedBaseURL = 'https://mms' + $ConfId + '.sched.com'
    $SessionDownloadPath = $DownloadLocation + '\mms' + $ConfId

    $EventTitle = [string]$EventItem.Title
    $EventUrl = [string]$EventItem.Url

    # Generate session info string
    $SessionInfoText = 'Session Title: ' + "`r`n" + $EventTitle + "`r`n`r`n"
    $DownloadTitle = $EventTitle -replace '[^A-Za-z0-9-_. ]', ''
    $DownloadTitle = $DownloadTitle.Trim()
    $DownloadTitle = $DownloadTitle -replace '\W+', '_'
    if (Test-IsBlank $DownloadTitle) {
        $DownloadTitle = [System.Uri]::UnescapeDataString($EventUrl) -replace '\W+', '_'
    }

    ## Set the download destination:
    $DownloadPath = $SessionDownloadPath + '\' + $DownloadTitle

    # Fetch the individual event page. Hosted files are discovered there.
    $SessionPage = $null
    try {
        $SessionPage = Invoke-SchedRequest -Uri ($SchedBaseURL + '/' + $EventUrl) -WebSession $WebSess
    }
    catch {
        $SessionPage = $null
    }

    if (-not $SessionPage) {
        Write-Log -Message ('FAILED: ' + $EventTitle) -Level 'ERROR'
        Add-DownloadResult -ConfId $ConfId -SessionTitle $EventTitle -FileName '' -Status 'Failed (session page)' -Path ''
        return
    }

    $SessionLinkInfo = $SessionPage.Content.Replace("`r", '').Replace("`n", '')

    ## Get session info if required:
    if (-not $ExcludeSessionDetails) {
        $DescriptionPattern = '<div class="tip-description">(?<description>.*?)(<div class="tip-roles">|<div class="sched-event-details-timeandplace">)'
        $DescriptionMatches = [System.Text.RegularExpressions.Regex]::Matches($SessionLinkInfo, $DescriptionPattern, $Script:RegexIgnoreCase)
        if ($DescriptionMatches.Count -gt 0) {
            $SessionInfoText += (Invoke-BasicHTMLParser -Html $DescriptionMatches[0].Groups['description'].Value -SchedBaseURL $SchedBaseURL) + "`r`n`r`n"
        }

        $RolesPattern = '<div class="tip-roles">(?<roles>.*?)<div class="sched-file">|<div class="sched-event-details-timeandplace">'
        $RolesMatches = [System.Text.RegularExpressions.Regex]::Matches($SessionLinkInfo, $RolesPattern, $Script:RegexIgnoreCase)
        if ($RolesMatches.Count -gt 0) {
            $SessionInfoText += (Invoke-BasicHTMLParser -Html $RolesMatches[0].Groups['roles'].Value -SchedBaseURL $SchedBaseURL) + "`r`n`r`n"
        }

        if ((Test-Path -LiteralPath $DownloadPath) -eq $false) { New-Item -ItemType Directory -Force -Path $DownloadPath | Out-Null }
        Out-File -LiteralPath ($DownloadPath + '\Session Info.txt') -InputObject $SessionInfoText -Force -Encoding UTF8
    }

    # Collect hosted-file links (de-duplicated by href)
    $Downloads = New-Object -TypeName System.Collections.Generic.List[object]
    $SeenHrefs = @{}
    if ($SessionPage.Links) {
        foreach ($Lnk in $SessionPage.Links) {
            $Href = [string]$Lnk.href
            if ($Href -like '*hosted-files*') {
                if (-not $SeenHrefs.ContainsKey($Href)) {
                    $SeenHrefs[$Href] = $true
                    $Downloads.Add([PSCustomObject]@{ href = $Href })
                }
            }
        }
    }

    # Fallback: parse hosted-files URLs directly from HTML when Links collection is sparse
    if ($Downloads.Count -eq 0 -and $SessionPage.Content) {
        $HostedFilesRegex = 'href="((?:https?:)?//hosted-files\.sched\.co/[^"]+)"'
        $HostedMatches = [System.Text.RegularExpressions.Regex]::Matches($SessionPage.Content, $HostedFilesRegex, $Script:RegexIgnoreCase)
        foreach ($HostedMatch in $HostedMatches) {
            $HostedUrl = $HostedMatch.Groups[1].Value
            if (-not $SeenHrefs.ContainsKey($HostedUrl)) {
                $SeenHrefs[$HostedUrl] = $true
                $Downloads.Add([PSCustomObject]@{ href = $HostedUrl })
            }
        }
    }

    if ($Downloads.Count -eq 0) {
        Write-Log -Message ('No files found for: ' + $EventTitle)
        Add-DownloadResult -ConfId $ConfId -SessionTitle $EventTitle -FileName '' -Status 'No files' -Path ''
        return
    }

    foreach ($Download in $Downloads) {
        if ($Script:CancelRequested) { return }

        $DownloadUrl = [string]$Download.href
        if ($DownloadUrl -like '//hosted-files.sched.co/*') {
            $DownloadUrl = 'https:' + $DownloadUrl
        }
        elseif ($DownloadUrl -notmatch '^https?://') {
            $DownloadUrl = $SchedBaseURL + '/' + $DownloadUrl
        }

        $FileName = Split-Path -Path $DownloadUrl -Leaf
        # Replace HTTP Encoding Characters (e.g. %20) with the proper equivalent.
        $FileName = [System.Web.HttpUtility]::UrlDecode($FileName)
        # Replace non-standard characters
        $FileName = $FileName -replace '[^A-Za-z0-9\.\-_ ]', ''

        $OutputFilePath = $DownloadPath + '\' + $FileName

        # Reduce Total Path to 255 characters.
        $OutputFilePathLen = $OutputFilePath.Length
        if ($OutputFilePathLen -ge 255) {
            $FileExt = [System.IO.Path]::GetExtension($OutputFilePath)
            $NewFileName = $OutputFilePath.Substring(0, ($OutputFilePathLen - $FileExt.Length))
            $NewFileName = $NewFileName.Substring(0, (255 - $FileExt.Length)).Trim()
            $OutputFilePath = $NewFileName + $FileExt
        }

        if ((Test-Path -LiteralPath $DownloadPath) -eq $false) { New-Item -ItemType Directory -Force -Path $DownloadPath | Out-Null }

        $FileStatus = Save-SessionFile -DownloadUrl $DownloadUrl -OutputFilePath $OutputFilePath -WebSession $WebSess -ReDownloadIfHashIsDifferent $ReDownloadIfHashIsDifferent
        Add-DownloadResult -ConfId $ConfId -SessionTitle $EventTitle -FileName $FileName -Status $FileStatus -Path $OutputFilePath
    }
}

# Runs a download job for a list of session objects (ConfId / Title / Url)
function Start-SessionDownload {
    param([Parameter(Mandatory = $true)][object[]]$Targets)

    if ($Script:IsBusy) { return }

    $DownloadLocation = $Script:TxtLocation.Text.Trim().TrimEnd('\')
    if (Test-IsBlank $DownloadLocation) { $DownloadLocation = $Script:DefaultDownloadLocation }
    $ExcludeDetails = [bool]$Script:ChkExclude.Checked
    $ReDownload = [bool]$Script:ChkRehash.Checked

    $Script:CancelRequested = $false
    $Script:Results.Clear()
    Set-UiBusy -Busy $true

    $Script:PbMain.Minimum = 0
    $Script:PbMain.Maximum = [System.Math]::Max($Targets.Count, 1)
    $Script:PbMain.Value = 0

    Write-Log -Message ('Base download folder is ' + $DownloadLocation)
    Write-Log -Message ('Starting download of ' + $Targets.Count + ' session(s).')

    try {
        for ($Index = 0; $Index -lt $Targets.Count; $Index++) {
            if ($Script:CancelRequested) { Write-Log -Message 'Download cancelled by user.' -Level 'WARN'; break }

            $Target = $Targets[$Index]
            Set-Status -Text ('Session {0} of {1}: [{2}] {3}' -f ($Index + 1), $Targets.Count, $Target.ConfId, $Target.Title)

            if (-not (Confirm-ConferenceSession -ConfId $Target.ConfId)) {
                Add-DownloadResult -ConfId $Target.ConfId -SessionTitle $Target.Title -FileName '' -Status 'Failed (login)' -Path ''
                $Script:PbMain.Value = $Index + 1
                continue
            }

            Save-ConferenceSession -ConfId $Target.ConfId -EventItem $Target -DownloadLocation $DownloadLocation -ExcludeSessionDetails $ExcludeDetails -ReDownloadIfHashIsDifferent $ReDownload
            $Script:PbMain.Value = $Index + 1
            [System.Windows.Forms.Application]::DoEvents()
        }
    }
    catch {
        $Err = $_
        Write-Log -Message ('Unexpected error: ' + $Err.Exception.Message) -Level 'ERROR'
    }
    finally {
        $Downloaded = @($Script:Results | Where-Object { $_.Status -eq 'Downloaded' }).Count
        $Updated = @($Script:Results | Where-Object { $_.Status -eq 'Updated' }).Count
        $Skipped = @($Script:Results | Where-Object { $_.Status -eq 'Skipped' }).Count
        $Failed = @($Script:Results | Where-Object { $_.Status -like 'Failed*' }).Count
        $NoFiles = @($Script:Results | Where-Object { $_.Status -eq 'No files' }).Count
        $Summary = 'Finished. Downloaded: {0} | Updated: {1} | Skipped: {2} | Failed: {3} | Sessions without files: {4}' -f $Downloaded, $Updated, $Skipped, $Failed, $NoFiles
        Write-Log -Message $Summary
        Set-Status -Text $Summary
        Export-DownloadResults
        Set-UiBusy -Busy $false
    }
}
# ==================================================
# ^^^^^ Download Functions ^^^^^
# ==================================================

# ==================================================
# MARK: Export Functions
# ==================================================
function Export-DownloadResults {
    if ($Script:Results.Count -eq 0) { return }

    $Stamp = Get-Date -Format 'yyyy-MM-dd_HHmm'
    $BaseName = $Stamp + '-MMS-Download-Results'
    $Data = @($Script:Results)
    $Exported = $false

    if (Get-Module -ListAvailable -Name ImportExcel) {
        try {
            Import-Module -Name ImportExcel -ErrorAction Stop
            $XlsxPath = Join-Path -Path $Script:OutputRoot -ChildPath ($BaseName + '.xlsx')
            $Data | Export-Excel -Path $XlsxPath -WorksheetName 'Results' -TableStyle Light21 -AutoSize -FreezeTopRow -ErrorAction Stop
            Write-Log -Message ('Results exported to ' + $XlsxPath)
            $Exported = $true
        }
        catch {
            $Err = $_
            Write-Log -Message ('Export-Excel failed (' + $Err.Exception.Message + '). Falling back to CSV.') -Level 'WARN'
        }
    }

    if (-not $Exported) {
        try {
            $CsvPath = Join-Path -Path $Script:OutputRoot -ChildPath ($BaseName + '.csv')
            $Data | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8
            Write-Log -Message ('Results exported to ' + $CsvPath)
        }
        catch {
            $Err = $_
            Write-Log -Message ('CSV export failed: ' + $Err.Exception.Message) -Level 'ERROR'
        }
    }
}
# ==================================================
# ^^^^^ Export Functions ^^^^^
# ==================================================

# ==================================================
# MARK: GUI State Functions
# ==================================================
function Sync-ConferenceChecks {
    if (-not $Script:LvConf) { return }
    foreach ($Item in $Script:LvConf.Items) { $Item.Tag.Checked = [bool]$Item.Checked }
}

function Sync-SessionChecks {
    if (-not $Script:LvSess) { return }
    foreach ($Item in $Script:LvSess.Items) { $Item.Tag.Checked = [bool]$Item.Checked }
}

# Shows only accessible conferences unless "Show unavailable conferences" is checked
function Update-ConferenceListView {
    if (-not $Script:LvConf -or $Script:LvConf.IsDisposed) { return }

    $Script:SuppressEvents = $true
    Sync-ConferenceChecks
    $Script:LvConf.BeginUpdate()
    $Script:LvConf.Items.Clear()
    foreach ($Conf in $Script:ConferenceCatalog) {
        $IsAccessible = ([string]$Conf.Access -like 'Yes*')
        if (-not $IsAccessible -and -not $Script:ChkShowAll.Checked) { continue }

        $Item = New-Object -TypeName System.Windows.Forms.ListViewItem -ArgumentList $Conf.Name
        [void]$Item.SubItems.Add($Conf.Id)
        [void]$Item.SubItems.Add([string]$Conf.Access)
        [void]$Item.SubItems.Add([string]$Conf.SessionCount)
        if (-not $IsAccessible) { $Item.ForeColor = [System.Drawing.Color]::Gray }
        $Item.Tag = $Conf
        $Item.Checked = [bool]$Conf.Checked
        [void]$Script:LvConf.Items.Add($Item)
    }
    $Script:LvConf.EndUpdate()
    $Script:SuppressEvents = $false
    [System.Windows.Forms.Application]::DoEvents()
}

# Shows sessions of the checked conferences (all loaded sessions if none are checked), then applies the text filter
function Update-SessionListView {
    if (-not $Script:LvSess -or $Script:LvSess.IsDisposed) { return }

    $Script:SuppressEvents = $true
    Sync-SessionChecks

    $CheckedIds = @{}
    foreach ($Conf in $Script:ConferenceCatalog) { if ($Conf.Checked) { $CheckedIds[$Conf.Id] = $true } }

    $Filter = $Script:TxtFilter.Text.Trim()
    $ShownCount = 0

    $Script:LvSess.BeginUpdate()
    $Script:LvSess.Items.Clear()
    foreach ($Sess in $Script:SessionList) {
        if ($CheckedIds.Count -gt 0 -and -not $CheckedIds.ContainsKey($Sess.ConfId)) { continue }
        if ($Filter) {
            $InTitle = ($Sess.Title.IndexOf($Filter, [System.StringComparison]::OrdinalIgnoreCase) -ge 0)
            $InConf = ($Sess.ConfId.IndexOf($Filter, [System.StringComparison]::OrdinalIgnoreCase) -ge 0)
            if (-not $InTitle -and -not $InConf) { continue }
        }
        $Item = New-Object -TypeName System.Windows.Forms.ListViewItem -ArgumentList $Sess.ConfId
        [void]$Item.SubItems.Add($Sess.Title)
        $Item.Tag = $Sess
        $Item.Checked = [bool]$Sess.Checked
        [void]$Script:LvSess.Items.Add($Item)
        $ShownCount++
    }
    $Script:LvSess.EndUpdate()
    $Script:SuppressEvents = $false

    $SelectedCount = @($Script:SessionList | Where-Object { $_.Checked }).Count
    $Script:LblSessCount.Text = 'Showing {0} of {1} session(s) | {2} selected' -f $ShownCount, $Script:SessionList.Count, $SelectedCount
    [System.Windows.Forms.Application]::DoEvents()
}

function Set-UiBusy {
    param([bool]$Busy)

    $Script:IsBusy = $Busy
    foreach ($Ctrl in @($Script:BtnQuery, $Script:BtnAddConf, $Script:BtnDlSelected, $Script:BtnDlAll, $Script:BtnBrowse)) {
        if ($Ctrl -and -not $Ctrl.IsDisposed) { $Ctrl.Enabled = -not $Busy }
    }
    if ($Script:BtnCancel -and -not $Script:BtnCancel.IsDisposed) { $Script:BtnCancel.Enabled = $Busy }
    [System.Windows.Forms.Application]::DoEvents()
}

# Returns the accessible conference objects that are checked
function Get-CheckedConferences {
    Sync-ConferenceChecks
    return @($Script:ConferenceCatalog | Where-Object { $_.Checked -and ([string]$_.Access -like 'Yes*') })
}
# ==================================================
# ^^^^^ GUI State Functions ^^^^^
# ==================================================

# ==================================================
# MARK: Build GUI
# ==================================================
$Script:Form = New-Object -TypeName System.Windows.Forms.Form
$Script:Form.Text = 'MMS Session Content Downloader (Sched)'
$Script:Form.Size = New-Object -TypeName System.Drawing.Size -ArgumentList 1200, 840
$Script:Form.MinimumSize = New-Object -TypeName System.Drawing.Size -ArgumentList 1000, 760
$Script:Form.StartPosition = 'CenterScreen'
$Script:Form.Font = New-Object -TypeName System.Drawing.Font -ArgumentList 'Segoe UI', 9

# ---------- Credentials / Query group ----------
$GrpCreds = New-Object -TypeName System.Windows.Forms.GroupBox
$GrpCreds.Text = 'Sched Credentials and Conference Discovery'
$GrpCreds.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 10, 8
$GrpCreds.Size = New-Object -TypeName System.Drawing.Size -ArgumentList 1160, 88
$GrpCreds.Anchor = 'Top, Left, Right'
$Script:Form.Controls.Add($GrpCreds)

$LblUser = New-Object -TypeName System.Windows.Forms.Label
$LblUser.Text = 'Username:'
$LblUser.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 12, 28
$LblUser.AutoSize = $true
$GrpCreds.Controls.Add($LblUser)

$Script:TxtUser = New-Object -TypeName System.Windows.Forms.TextBox
$Script:TxtUser.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 85, 25
$Script:TxtUser.Size = New-Object -TypeName System.Drawing.Size -ArgumentList 230, 24
$GrpCreds.Controls.Add($Script:TxtUser)

$LblPass = New-Object -TypeName System.Windows.Forms.Label
$LblPass.Text = 'Password:'
$LblPass.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 335, 28
$LblPass.AutoSize = $true
$GrpCreds.Controls.Add($LblPass)

$Script:TxtPass = New-Object -TypeName System.Windows.Forms.TextBox
$Script:TxtPass.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 405, 25
$Script:TxtPass.Size = New-Object -TypeName System.Drawing.Size -ArgumentList 230, 24
$Script:TxtPass.UseSystemPasswordChar = $true
$GrpCreds.Controls.Add($Script:TxtPass)

$Script:BtnQuery = New-Object -TypeName System.Windows.Forms.Button
$Script:BtnQuery.Text = 'Re-Query Conferences'
$Script:BtnQuery.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 655, 22
$Script:BtnQuery.Size = New-Object -TypeName System.Drawing.Size -ArgumentList 160, 30
$GrpCreds.Controls.Add($Script:BtnQuery)

$Script:ChkProbe = New-Object -TypeName System.Windows.Forms.CheckBox
$Script:ChkProbe.Text = 'Probe Sched for other MMS conference IDs'
$Script:ChkProbe.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 830, 27
$Script:ChkProbe.AutoSize = $true
$Script:ChkProbe.Checked = $true
$GrpCreds.Controls.Add($Script:ChkProbe)

$LblCustom = New-Object -TypeName System.Windows.Forms.Label
$LblCustom.Text = 'Add conference ID manually (e.g. 2027atmoa):'
$LblCustom.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 12, 61
$LblCustom.AutoSize = $true
$GrpCreds.Controls.Add($LblCustom)

$Script:TxtCustom = New-Object -TypeName System.Windows.Forms.TextBox
$Script:TxtCustom.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 290, 58
$Script:TxtCustom.Size = New-Object -TypeName System.Drawing.Size -ArgumentList 160, 24
$GrpCreds.Controls.Add($Script:TxtCustom)

$Script:BtnAddConf = New-Object -TypeName System.Windows.Forms.Button
$Script:BtnAddConf.Text = 'Add and Check'
$Script:BtnAddConf.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 460, 56
$Script:BtnAddConf.Size = New-Object -TypeName System.Drawing.Size -ArgumentList 110, 27
$GrpCreds.Controls.Add($Script:BtnAddConf)

# ---------- Conferences group ----------
$GrpConf = New-Object -TypeName System.Windows.Forms.GroupBox
$GrpConf.Text = 'Conferences You Can Access (check to filter sessions / download all)'
$GrpConf.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 10, 104
$GrpConf.Size = New-Object -TypeName System.Drawing.Size -ArgumentList 560, 330
$GrpConf.Anchor = 'Top, Left'
$Script:Form.Controls.Add($GrpConf)

$Script:LvConf = New-Object -TypeName System.Windows.Forms.ListView
$Script:LvConf.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 10, 22
$Script:LvConf.Size = New-Object -TypeName System.Drawing.Size -ArgumentList 540, 260
$Script:LvConf.View = 'Details'
$Script:LvConf.CheckBoxes = $true
$Script:LvConf.FullRowSelect = $true
$Script:LvConf.GridLines = $true
$Script:LvConf.HideSelection = $false
[void]$Script:LvConf.Columns.Add('Conference', 210)
[void]$Script:LvConf.Columns.Add('ID', 85)
[void]$Script:LvConf.Columns.Add('Access', 165)
[void]$Script:LvConf.Columns.Add('Sessions', 60)
$GrpConf.Controls.Add($Script:LvConf)

$BtnConfAll = New-Object -TypeName System.Windows.Forms.Button
$BtnConfAll.Text = 'Select All'
$BtnConfAll.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 10, 290
$BtnConfAll.Size = New-Object -TypeName System.Drawing.Size -ArgumentList 90, 28
$GrpConf.Controls.Add($BtnConfAll)

$BtnConfNone = New-Object -TypeName System.Windows.Forms.Button
$BtnConfNone.Text = 'Select None'
$BtnConfNone.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 106, 290
$BtnConfNone.Size = New-Object -TypeName System.Drawing.Size -ArgumentList 90, 28
$GrpConf.Controls.Add($BtnConfNone)

$Script:ChkShowAll = New-Object -TypeName System.Windows.Forms.CheckBox
$Script:ChkShowAll.Text = 'Show unavailable conferences'
$Script:ChkShowAll.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 210, 295
$Script:ChkShowAll.AutoSize = $true
$GrpConf.Controls.Add($Script:ChkShowAll)

# ---------- Sessions group ----------
$GrpSess = New-Object -TypeName System.Windows.Forms.GroupBox
$GrpSess.Text = 'Sessions (check the ones to download)'
$GrpSess.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 580, 104
$GrpSess.Size = New-Object -TypeName System.Drawing.Size -ArgumentList 590, 330
$GrpSess.Anchor = 'Top, Left, Right'
$Script:Form.Controls.Add($GrpSess)

$LblFilter = New-Object -TypeName System.Windows.Forms.Label
$LblFilter.Text = 'Filter:'
$LblFilter.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 10, 27
$LblFilter.AutoSize = $true
$GrpSess.Controls.Add($LblFilter)

$Script:TxtFilter = New-Object -TypeName System.Windows.Forms.TextBox
$Script:TxtFilter.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 55, 24
$Script:TxtFilter.Size = New-Object -TypeName System.Drawing.Size -ArgumentList 525, 24
$Script:TxtFilter.Anchor = 'Top, Left, Right'
$GrpSess.Controls.Add($Script:TxtFilter)

$Script:LvSess = New-Object -TypeName System.Windows.Forms.ListView
$Script:LvSess.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 10, 54
$Script:LvSess.Size = New-Object -TypeName System.Drawing.Size -ArgumentList 570, 228
$Script:LvSess.Anchor = 'Top, Left, Right'
$Script:LvSess.View = 'Details'
$Script:LvSess.CheckBoxes = $true
$Script:LvSess.FullRowSelect = $true
$Script:LvSess.GridLines = $true
$Script:LvSess.HideSelection = $false
[void]$Script:LvSess.Columns.Add('Conf', 85)
[void]$Script:LvSess.Columns.Add('Session Title', 460)
$GrpSess.Controls.Add($Script:LvSess)

$BtnSessAll = New-Object -TypeName System.Windows.Forms.Button
$BtnSessAll.Text = 'Select All Shown'
$BtnSessAll.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 10, 290
$BtnSessAll.Size = New-Object -TypeName System.Drawing.Size -ArgumentList 120, 28
$GrpSess.Controls.Add($BtnSessAll)

$BtnSessNone = New-Object -TypeName System.Windows.Forms.Button
$BtnSessNone.Text = 'Clear Selection'
$BtnSessNone.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 136, 290
$BtnSessNone.Size = New-Object -TypeName System.Drawing.Size -ArgumentList 110, 28
$GrpSess.Controls.Add($BtnSessNone)

$Script:LblSessCount = New-Object -TypeName System.Windows.Forms.Label
$Script:LblSessCount.Text = 'No sessions loaded yet'
$Script:LblSessCount.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 255, 296
$Script:LblSessCount.AutoSize = $true
$GrpSess.Controls.Add($Script:LblSessCount)

# ---------- Download options group ----------
$GrpDl = New-Object -TypeName System.Windows.Forms.GroupBox
$GrpDl.Text = 'Download'
$GrpDl.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 10, 442
$GrpDl.Size = New-Object -TypeName System.Drawing.Size -ArgumentList 1160, 110
$GrpDl.Anchor = 'Top, Left, Right'
$Script:Form.Controls.Add($GrpDl)

$LblLocation = New-Object -TypeName System.Windows.Forms.Label
$LblLocation.Text = 'Download folder:'
$LblLocation.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 12, 27
$LblLocation.AutoSize = $true
$GrpDl.Controls.Add($LblLocation)

$Script:TxtLocation = New-Object -TypeName System.Windows.Forms.TextBox
$Script:TxtLocation.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 110, 24
$Script:TxtLocation.Size = New-Object -TypeName System.Drawing.Size -ArgumentList 560, 24
$Script:TxtLocation.Text = $Script:DefaultDownloadLocation
$GrpDl.Controls.Add($Script:TxtLocation)

$Script:BtnBrowse = New-Object -TypeName System.Windows.Forms.Button
$Script:BtnBrowse.Text = 'Browse...'
$Script:BtnBrowse.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 680, 22
$Script:BtnBrowse.Size = New-Object -TypeName System.Drawing.Size -ArgumentList 80, 27
$GrpDl.Controls.Add($Script:BtnBrowse)

$BtnOpenFolder = New-Object -TypeName System.Windows.Forms.Button
$BtnOpenFolder.Text = 'Open Folder'
$BtnOpenFolder.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 766, 22
$BtnOpenFolder.Size = New-Object -TypeName System.Drawing.Size -ArgumentList 90, 27
$GrpDl.Controls.Add($BtnOpenFolder)

$Script:ChkExclude = New-Object -TypeName System.Windows.Forms.CheckBox
$Script:ChkExclude.Text = 'Exclude "Session Info.txt"'
$Script:ChkExclude.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 880, 25
$Script:ChkExclude.AutoSize = $true
$GrpDl.Controls.Add($Script:ChkExclude)

$Script:ChkRehash = New-Object -TypeName System.Windows.Forms.CheckBox
$Script:ChkRehash.Text = 'Re-download if file hash differs'
$Script:ChkRehash.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 880, 50
$Script:ChkRehash.AutoSize = $true
$GrpDl.Controls.Add($Script:ChkRehash)

$Script:BtnDlSelected = New-Object -TypeName System.Windows.Forms.Button
$Script:BtnDlSelected.Text = 'Download SELECTED Sessions'
$Script:BtnDlSelected.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 12, 62
$Script:BtnDlSelected.Size = New-Object -TypeName System.Drawing.Size -ArgumentList 230, 36
$GrpDl.Controls.Add($Script:BtnDlSelected)

$Script:BtnDlAll = New-Object -TypeName System.Windows.Forms.Button
$Script:BtnDlAll.Text = 'Download ALL Sessions of Checked Conferences'
$Script:BtnDlAll.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 250, 62
$Script:BtnDlAll.Size = New-Object -TypeName System.Drawing.Size -ArgumentList 330, 36
$GrpDl.Controls.Add($Script:BtnDlAll)

$Script:BtnCancel = New-Object -TypeName System.Windows.Forms.Button
$Script:BtnCancel.Text = 'Cancel'
$Script:BtnCancel.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 590, 62
$Script:BtnCancel.Size = New-Object -TypeName System.Drawing.Size -ArgumentList 90, 36
$Script:BtnCancel.Enabled = $false
$GrpDl.Controls.Add($Script:BtnCancel)

# ---------- Progress / status / log ----------
$Script:PbMain = New-Object -TypeName System.Windows.Forms.ProgressBar
$Script:PbMain.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 10, 562
$Script:PbMain.Size = New-Object -TypeName System.Drawing.Size -ArgumentList 1160, 20
$Script:PbMain.Anchor = 'Top, Left, Right'
$Script:Form.Controls.Add($Script:PbMain)

$Script:LblStatus = New-Object -TypeName System.Windows.Forms.Label
$Script:LblStatus.Text = 'Starting ...'
$Script:LblStatus.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 10, 588
$Script:LblStatus.Size = New-Object -TypeName System.Drawing.Size -ArgumentList 1160, 20
$Script:LblStatus.Anchor = 'Top, Left, Right'
$Script:Form.Controls.Add($Script:LblStatus)

$Script:TxtLog = New-Object -TypeName System.Windows.Forms.TextBox
$Script:TxtLog.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 10, 612
$Script:TxtLog.Size = New-Object -TypeName System.Drawing.Size -ArgumentList 1160, 180
$Script:TxtLog.Anchor = 'Top, Bottom, Left, Right'
$Script:TxtLog.Multiline = $true
$Script:TxtLog.ReadOnly = $true
$Script:TxtLog.ScrollBars = 'Vertical'
$Script:TxtLog.Font = New-Object -TypeName System.Drawing.Font -ArgumentList 'Consolas', 9
$Script:Form.Controls.Add($Script:TxtLog)
# ==================================================
# ^^^^^ Build GUI ^^^^^
# ==================================================

# ==================================================
# MARK: Event Handlers
# ==================================================
# ---------- Re-Query Conferences ----------
$Script:BtnQuery.Add_Click({ Invoke-ConferenceDiscovery })

# ---------- Add custom conference ID and check it immediately ----------
$Script:BtnAddConf.Add_Click({
        if ($Script:IsBusy) { return }
        $NewId = $Script:TxtCustom.Text.Trim().ToLower()
        if ($NewId -notmatch '^[a-z0-9]+$') {
            [void][System.Windows.Forms.MessageBox]::Show('Conference IDs may only contain letters and numbers (for example 2027atmoa).', 'MMS Session Downloader', 'OK', 'Warning')
            return
        }
        $Conf = Get-ConferenceInfo -ConfId $NewId
        if (-not $Conf) {
            $Conf = New-ConferenceEntry -Id $NewId -Name ('MMS ' + $NewId + ' (manual)') -Type 'Discovered'
            $Script:ConferenceCatalog.Add($Conf)
        }
        $Script:TxtCustom.Text = ''

        Set-UiBusy -Busy $true
        try {
            $Script:Credential = Get-CurrentCredential
            for ($Index = $Script:SessionList.Count - 1; $Index -ge 0; $Index--) {
                if ($Script:SessionList[$Index].ConfId -eq $NewId) { $Script:SessionList.RemoveAt($Index) }
            }
            Test-ConferenceAccess -Conf $Conf -Credential $Script:Credential
            Update-ConferenceListView
            Update-SessionListView
            Set-Status -Text ($NewId + ': ' + $Conf.Access)
        }
        finally {
            Set-UiBusy -Busy $false
        }
    })

# ---------- Conference checks drive the session list ----------
$Script:LvConf.Add_ItemChecked({
        if ($Script:SuppressEvents) { return }
        Sync-ConferenceChecks
        Update-SessionListView
    })

$BtnConfAll.Add_Click({
        $Script:SuppressEvents = $true
        foreach ($Item in $Script:LvConf.Items) { if ([string]$Item.Tag.Access -like 'Yes*') { $Item.Checked = $true } }
        $Script:SuppressEvents = $false
        Sync-ConferenceChecks
        Update-SessionListView
    })

$BtnConfNone.Add_Click({
        $Script:SuppressEvents = $true
        foreach ($Item in $Script:LvConf.Items) { $Item.Checked = $false }
        $Script:SuppressEvents = $false
        Sync-ConferenceChecks
        Update-SessionListView
    })

$Script:ChkShowAll.Add_CheckedChanged({ Update-ConferenceListView })

# ---------- Session list helpers ----------
$Script:TxtFilter.Add_TextChanged({ Update-SessionListView })

$Script:LvSess.Add_ItemChecked({
        if ($Script:SuppressEvents) { return }
        $_.Item.Tag.Checked = [bool]$_.Item.Checked
        $SelectedCount = @($Script:SessionList | Where-Object { $_.Checked }).Count
        $Script:LblSessCount.Text = 'Showing {0} of {1} session(s) | {2} selected' -f $Script:LvSess.Items.Count, $Script:SessionList.Count, $SelectedCount
    })

$BtnSessAll.Add_Click({
        $Script:SuppressEvents = $true
        foreach ($Item in $Script:LvSess.Items) { $Item.Checked = $true }
        $Script:SuppressEvents = $false
        Update-SessionListView
    })

$BtnSessNone.Add_Click({
        foreach ($Sess in $Script:SessionList) { $Sess.Checked = $false }
        $Script:SuppressEvents = $true
        foreach ($Item in $Script:LvSess.Items) { $Item.Checked = $false }
        $Script:SuppressEvents = $false
        Update-SessionListView
    })

# ---------- Browse / Open folder ----------
$Script:BtnBrowse.Add_Click({
        $Dialog = New-Object -TypeName System.Windows.Forms.FolderBrowserDialog
        $Dialog.Description = 'Select the base download folder'
        if (Test-Path -Path $Script:TxtLocation.Text) { $Dialog.SelectedPath = $Script:TxtLocation.Text }
        if ($Dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $Script:TxtLocation.Text = $Dialog.SelectedPath }
    })

$BtnOpenFolder.Add_Click({
        $FolderPath = $Script:TxtLocation.Text.Trim()
        if (Test-Path -Path $FolderPath) { Start-Process -FilePath 'explorer.exe' -ArgumentList ('"' + $FolderPath + '"') }
        else { [void][System.Windows.Forms.MessageBox]::Show('The download folder does not exist yet.', 'MMS Session Downloader', 'OK', 'Information') }
    })

# ---------- Download SELECTED sessions ----------
$Script:BtnDlSelected.Add_Click({
        if ($Script:IsBusy) { return }

        Sync-SessionChecks
        $Targets = @($Script:SessionList | Where-Object { $_.Checked })
        if ($Targets.Count -eq 0) {
            [void][System.Windows.Forms.MessageBox]::Show('No sessions are checked. Check the sessions you want in the Sessions list.', 'MMS Session Downloader', 'OK', 'Information')
            return
        }
        Start-SessionDownload -Targets $Targets
    })

# ---------- Download ALL sessions of checked conferences ----------
$Script:BtnDlAll.Add_Click({
        if ($Script:IsBusy) { return }

        $Confs = @(Get-CheckedConferences)
        if ($Confs.Count -eq 0) {
            [void][System.Windows.Forms.MessageBox]::Show('Check at least one accessible conference first.', 'MMS Session Downloader', 'OK', 'Information')
            return
        }

        $ConfIds = @{}
        foreach ($Conf in $Confs) { $ConfIds[$Conf.Id] = $true }
        $Targets = @($Script:SessionList | Where-Object { $ConfIds.ContainsKey($_.ConfId) })

        if ($Targets.Count -eq 0) {
            [void][System.Windows.Forms.MessageBox]::Show('No sessions were found for the checked conferences.', 'MMS Session Downloader', 'OK', 'Information')
            return
        }

        $Answer = [System.Windows.Forms.MessageBox]::Show(('Download ALL files for {0} session(s) across {1} conference(s)?' -f $Targets.Count, $Confs.Count), 'MMS Session Downloader', 'YesNo', 'Question')
        if ($Answer -eq 'Yes') { Start-SessionDownload -Targets $Targets }
    })

# ---------- Cancel ----------
$Script:BtnCancel.Add_Click({
        $Script:CancelRequested = $true
        Write-Log -Message 'Cancel requested. Stopping after the current operation ...' -Level 'WARN'
    })

# ---------- Form events ----------
$Script:Form.Add_FormClosing({
        $Script:CancelRequested = $true
    })

# Automatically discover conferences as soon as the window opens
$Script:Form.Add_Shown({
        Write-Log -Message ('MMS Session Content Downloader started. Log file: ' + $Script:LogFile)
        Update-ConferenceListView
        $Script:Form.Activate()
        Invoke-ConferenceDiscovery
    })
# ==================================================
# ^^^^^ Event Handlers ^^^^^
# ==================================================

# ==================================================
# MARK: Launch GUI
# ==================================================
$StartupCred = Show-CredentialPrompt
if ($StartupCred) {
    $Script:TxtUser.Text = $StartupCred.UserName
    $Script:TxtPass.Text = $StartupCred.GetNetworkCredential().Password
}
$StartupCred = $null

[void]$Script:Form.ShowDialog()
$Script:Form.Dispose()
# ==================================================
# ^^^^^ Launch GUI ^^^^^
# ==================================================