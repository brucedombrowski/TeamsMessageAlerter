# TeamsMessageAlerter.ps1
# Monitors Windows notification history for new Teams messages.
# Alerts by sending email from your work Outlook to your personal email.
# No installs required. Works on Windows PowerShell 5.1.
#
# FIRST RUN: script will prompt for your personal email address.
# Settings saved to TeamsMessageAlerter.config.json alongside this script.
# Delete the config file to re-run setup prompts.
#
# KNOWN LIMITATIONS
#   - Must run under Windows PowerShell 5.1 (powershell.exe), NOT PowerShell 7 (pwsh.exe).
#     In VS Code: Ctrl+Shift+P -> "Select Default Profile" -> "Windows PowerShell"
#   - Focus Assist / Do Not Disturb suppresses toast history entirely.
#     Ensure Focus Assist is OFF on the laptop for reliable alerting.
#   - Teams notification XML varies (chat vs channel vs reaction vs call).
#     Sender/message extraction is best-effort from the first two <text> nodes.
#
# TASK SCHEDULER (silent startup at login - set up after testing is confirmed working)
#   Trigger : At log on
#   Action  : powershell.exe
#   Args    : -WindowStyle Hidden -ExecutionPolicy Bypass -File "C:\path\to\TeamsMessageAlerter.ps1"
#   NOTE    : -WindowStyle Hidden required - omitting it shows a console at every login.
#
# LOG FILES  (written alongside this script)
#   TeamsMessageAlerter.log            - timestamped event log
#   TeamsMessageAlerter_transcript.log - full PowerShell transcript
#   TeamsMessageAlerter.config.json    - saved settings (delete to re-run setup)

# ==============================================================================
# SCRIPT CONFIG  -- safe defaults, do not normally need editing
# ==============================================================================

$script:PollSeconds        = 10     # seconds between notification checks
$script:CooldownSecs       = 30     # min seconds between alerts per sender

# Email subject/body limits - keeps alerts concise
$script:SmsMaxSubjectChars = 60
$script:SmsMaxBodyChars    = 160

# Truncation suffix - ASCII only
$script:TruncationSuffix   = "..."

# Teams AppIds to probe - varies by Teams version installed
$script:TeamsAppIds        = @(
    "MSTeams",
    "com.squirrel.Teams.Teams",
    "MicrosoftTeams_8wekyb3d8bbwe!MSTeams"
)

# Outlook COM constant for a mail item
# Ref: https://learn.microsoft.com/en-us/office/vba/api/outlook.olmailitem
$script:OlMailItem         = 0

# Fallback notification ID when XML hashing fails.
# Fixed constant is intentional - Get-Random would cause alert flood.
$script:FallbackNotifId    = "fallback-unparseable"

# Display strings used when notification fields cannot be extracted
$script:DefaultSender      = "Teams"
$script:DefaultMessage     = "(no preview)"
$script:ParseErrorMessage  = "(could not parse notification)"

# Timestamp format used in log entries
$script:LogTimestampFormat = "yyyy-MM-dd HH:mm:ss"

# ==============================================================================
# RUNTIME STATE -- do not edit
# ==============================================================================

$script:AlertEmailTo   = $null    # personal email, loaded from config
$script:WebhookUrl     = $null    # optional Teams webhook for test messages
$script:WorkingAppId   = $null    # cached Teams AppId
$script:Outlook        = $null    # Outlook COM object

$script:LogDir         = if ($PSScriptRoot) { $PSScriptRoot } else { "$env:USERPROFILE\TeamsAlerter" }
$script:LogFile        = "$($script:LogDir)\TeamsMessageAlerter.log"
$script:TranscriptFile = "$($script:LogDir)\TeamsMessageAlerter_transcript.log"
$script:ConfigFile     = "$($script:LogDir)\TeamsMessageAlerter.config.json"

$script:Seen           = [System.Collections.Generic.HashSet[string]]::new()
$script:Cooldown       = [System.Collections.Generic.Dictionary[string,datetime]]::new()

# ==============================================================================
# FUNCTIONS
# ==============================================================================

function Write-Log {
    param([string]$Message)
    $line = "$(Get-Date -f $script:LogTimestampFormat)  $Message"
    Write-Host $line
    try { Add-Content -Path $script:LogFile -Value $line -Encoding UTF8 } catch {}
}

# ------------------------------------------------------------------------------
# SETUP
# ------------------------------------------------------------------------------

function Get-ValidEmail {
    while ($true) {
        $input = (Read-Host "Enter your personal email address").Trim()
        if ($input -match '^[^@\s]+@[^@\s]+\.[^@\s]+$') { return $input }
        Write-Host "  Invalid email address. Try again."
    }
}

function Get-WebhookUrl {
    Write-Host ""
    Write-Host "Incoming Webhook URL for end-to-end testing (optional)."
    Write-Host "Sends a real Teams message to yourself to test the full alert pipeline."
    Write-Host ""
    Write-Host "HOW TO SET UP (Power Automate Workflows - Connectors retired Dec 2025):"
    Write-Host "  1. In Teams, click '...' next to any channel you own"
    Write-Host "  2. Select 'Workflows'"
    Write-Host "  3. Search for 'Post to a channel when a webhook request is received'"
    Write-Host "  4. Click it -> authenticate your account -> select your Team and Channel"
    Write-Host "  5. Click 'Add workflow'"
    Write-Host "  6. Copy the URL shown and paste it below"
    Write-Host ""
    Write-Host "  NOTE: If 'Workflows' is missing from the menu, IT has Power Automate"
    Write-Host "  blocked. Skip this step and ask a colleague to send a test message instead."
    Write-Host ""
    Write-Host "Press Enter to skip - you can add it later by deleting the config file."
    return (Read-Host "Paste webhook URL (or press Enter to skip)").Trim()
}

function Save-Config {
    $config = @{
        AlertEmailTo = $script:AlertEmailTo
        WebhookUrl   = $script:WebhookUrl
    } | ConvertTo-Json
    Set-Content -Path $script:ConfigFile -Value $config -Encoding UTF8
    Write-Log "Settings saved to: $script:ConfigFile"
}

function Load-Config {
    if (-not (Test-Path $script:ConfigFile)) { return $false }
    try {
        $config = Get-Content $script:ConfigFile -Raw | ConvertFrom-Json
        if ([string]::IsNullOrEmpty($config.AlertEmailTo)) { return $false }
        $script:AlertEmailTo = $config.AlertEmailTo
        $script:WebhookUrl   = $config.WebhookUrl
        return $true
    } catch {
        Write-Log "Warning: could not read config file, re-running setup. $_"
        return $false
    }
}

function Run-Setup {
    Write-Host ""
    Write-Host "========================================"
    Write-Host "  TeamsMessageAlerter - First Run Setup"
    Write-Host "========================================"
    Write-Host ""
    Write-Host "Alerts will be sent from your work Outlook to your personal email."
    Write-Host ""

    $script:AlertEmailTo = Get-ValidEmail
    $script:WebhookUrl   = Get-WebhookUrl

    Write-Host ""
    Write-Host "  Alerts will be sent to: $($script:AlertEmailTo)"

    Save-Config

    Write-Host ""
    Write-Host "Setup complete. Starting monitor..."
    Write-Host ""
}

# ------------------------------------------------------------------------------
# WEBHOOK TEST
# ------------------------------------------------------------------------------

function Send-TestTeamsMessage {
    # Posts a message to Teams via Power Automate Workflow webhook.
    # Uses Adaptive Card format required by the new workflow template.
    # Old MessageCard format {"text":"..."} no longer works after Connector retirement Dec 2025.
    param([string]$Text = "TeamsMessageAlerter test - if you receive an email alert the pipeline is working!")

    if ([string]::IsNullOrEmpty($script:WebhookUrl)) {
        Write-Log "Test skipped - no webhook URL configured. Delete config file and re-run setup to add one."
        return $false
    }
    try {
        $adaptiveCard = @{
            type    = "AdaptiveCard"
            version = "1.2"
            body    = @(
                @{ type = "TextBlock"; text = $Text; wrap = $true }
            )
        }
        $payload = @{
            attachments = @(
                @{
                    contentType = "application/vnd.microsoft.card.adaptive"
                    content     = $adaptiveCard
                }
            )
        } | ConvertTo-Json -Depth 10

        Invoke-RestMethod -Uri $script:WebhookUrl -Method Post -Body $payload -ContentType "application/json" | Out-Null
        Write-Log "Test message posted to Teams. Watch for the notification then check your email for the alert..."
        return $true
    } catch {
        Write-Log "Failed to post test message: $_"
        return $false
    }
}

# ------------------------------------------------------------------------------
# WINRT
# ------------------------------------------------------------------------------

function Load-WinRT {
    try {
        [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType=WindowsRuntime] | Out-Null
        return $true
    } catch {
        Write-Log "ERROR: Could not load WinRT. Must run under Windows PowerShell 5.1 (powershell.exe) not PowerShell 7 (pwsh.exe). $_"
        return $false
    }
}

function Get-TeamsNotifications {
    $history = [Windows.UI.Notifications.ToastNotificationManager]::History

    if ($script:WorkingAppId) {
        try {
            return $history.GetHistory($script:WorkingAppId)
        } catch {
            Write-Log "Warning: cached AppId '$($script:WorkingAppId)' failed, re-probing. $_"
            $script:WorkingAppId = $null
        }
    }

    foreach ($appId in $script:TeamsAppIds) {
        try {
            $items = $history.GetHistory($appId)
            $script:WorkingAppId = $appId
            Write-Log "Teams AppId confirmed: $appId ($($items.Count) current notifications)"
            return $items
        } catch {
            Write-Log "Warning: GetHistory failed for '$appId': $_"
        }
    }

    Write-Log "Warning: No working Teams AppId found. Is Teams running?"
    return @()
}

# ------------------------------------------------------------------------------
# NOTIFICATION PARSING
# ------------------------------------------------------------------------------

function Get-NotificationId {
    param([string]$Xml)
    $hash = $null
    try {
        $hash  = [System.Security.Cryptography.SHA256]::Create()
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($Xml)
        return [Convert]::ToBase64String($hash.ComputeHash($bytes))
    } catch {
        Write-Log "Warning: could not hash notification XML - using fallback ID. $_"
        return $script:FallbackNotifId
    } finally {
        if ($null -ne $hash) { $hash.Dispose() }
    }
}

function Get-TextContent {
    # PowerShell XML returns [string] for single node, [XmlElement] for multi-node.
    param($Node)
    if ($null -eq $Node)                   { return $null }
    if ($Node -is [string])                { return $Node }
    if ($Node -is [System.Xml.XmlElement]) { return $Node.InnerText }
    return $null
}

function Parse-Notification {
    param([string]$Xml)
    try {
        $doc     = [xml]$Xml
        $binding = @($doc.toast.visual.binding)[0]
        $texts   = @($binding.text)
        $sender  = Get-TextContent $texts[0]
        $message = Get-TextContent $texts[1]
        return @{
            Sender  = if ($sender)  { $sender }  else { $script:DefaultSender }
            Message = if ($message) { $message } else { $script:DefaultMessage }
        }
    } catch {
        Write-Log "Warning: failed to parse notification XML: $_"
        return @{ Sender = $script:DefaultSender; Message = $script:ParseErrorMessage }
    }
}

# ------------------------------------------------------------------------------
# HELPERS
# ------------------------------------------------------------------------------

function Limit-Length {
    param([string]$Text, [int]$Max)
    if ([string]::IsNullOrEmpty($Text))  { return "" }
    if ($Text.Length -le $Max)           { return $Text }
    $suffixLen = $script:TruncationSuffix.Length
    if ($Max -le $suffixLen)             { return $script:TruncationSuffix.Substring(0, $Max) }
    return $Text.Substring(0, $Max - $suffixLen) + $script:TruncationSuffix
}

# ------------------------------------------------------------------------------
# OUTLOOK
# ------------------------------------------------------------------------------

function Release-Outlook {
    if ($null -ne $script:Outlook) {
        try {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($script:Outlook) | Out-Null
        } catch {}
        $script:Outlook = $null
    }
}

function Initialize-Outlook {
    Release-Outlook
    try {
        $script:Outlook = New-Object -ComObject Outlook.Application
        Write-Log "Outlook COM initialized."
        return $true
    } catch {
        Write-Log "ERROR: Could not connect to Outlook. Ensure Outlook is open and signed in. $_"
        return $false
    }
}

function Send-MailItem {
    param([string]$Subject, [string]$Body)
    $mail         = $script:Outlook.CreateItem($script:OlMailItem)
    $mail.To      = $script:AlertEmailTo
    $mail.Subject = $Subject
    $mail.Body    = $Body
    $mail.Send()
}

function Send-Alert {
    # Returns $true on success, $false on failure.
    # Cooldown only advances on success so a failed send retries next poll.
    param([string]$Subject, [string]$Body)
    $Subject = Limit-Length $Subject $script:SmsMaxSubjectChars
    $Body    = Limit-Length $Body    $script:SmsMaxBodyChars
    try {
        Send-MailItem $Subject $Body
        Write-Log "Alert sent -> $Subject"
        return $true
    } catch {
        Write-Log "Outlook error, attempting reconnect: $_"
        if (Initialize-Outlook) {
            try {
                Send-MailItem $Subject $Body
                Write-Log "Alert sent after reconnect."
                return $true
            } catch {
                Write-Log "Send failed after reconnect: $_"
            }
        }
    }
    return $false
}

# ==============================================================================
# STARTUP
# ==============================================================================

if (-not (Test-Path $script:LogDir)) {
    New-Item -ItemType Directory -Path $script:LogDir -Force | Out-Null
}

Start-Transcript -Path $script:TranscriptFile -Append *>$null

try {
    # Load saved settings or run first-time setup
    if (-not (Load-Config)) { Run-Setup }

    Write-Log "TeamsMessageAlerter starting. Alerts -> $($script:AlertEmailTo)"

    if (-not (Load-WinRT))         { throw "WinRT unavailable - must use Windows PowerShell 5.1." }
    if (-not (Initialize-Outlook)) { throw "Outlook unavailable - ensure Outlook is open and signed in." }

    # Send startup confirmation so you know it's running
    Send-Alert "TeamsMessageAlerter" "Monitor started. Alerts will be sent to this address." | Out-Null

    # If webhook is configured, send a test Teams message for end-to-end validation
    if (-not [string]::IsNullOrEmpty($script:WebhookUrl)) {
        Write-Log "Webhook configured - sending test Teams message for end-to-end validation..."
        Send-TestTeamsMessage | Out-Null
    }

    # Seed dedup set with existing notifications so we don't alert on old messages
    foreach ($n in (Get-TeamsNotifications)) {
        try {
            $null = $script:Seen.Add((Get-NotificationId $n.Content.GetXml()))
        } catch {
            Write-Log "Warning: could not seed notification during startup: $_"
        }
    }
    Write-Log "Seeded $($script:Seen.Count) existing notifications. Watching for new ones..."

    # ── MAIN LOOP ──────────────────────────────────────────────────────────────
    while ($true) {
        Start-Sleep -Seconds $script:PollSeconds

        foreach ($n in (Get-TeamsNotifications)) {
            $xml = $null
            try {
                $xml = $n.Content.GetXml()
            } catch {
                Write-Log "Warning: could not read notification XML: $_"
                continue
            }

            $id = Get-NotificationId $xml

            if (-not $script:Seen.Contains($id)) {
                $null = $script:Seen.Add($id)

                $parsed  = Parse-Notification $xml
                $sender  = $parsed.Sender
                $message = $parsed.Message

                Write-Log "New Teams message from: $sender"

                $now      = [datetime]::Now
                $lastSent = if ($script:Cooldown.ContainsKey($sender)) {
                                $script:Cooldown[$sender]
                            } else {
                                [datetime]::MinValue
                            }

                if (($now - $lastSent).TotalSeconds -ge $script:CooldownSecs) {
                    if (Send-Alert "Teams: $sender" $message) {
                        $script:Cooldown[$sender] = $now
                    }
                } else {
                    Write-Log "Cooldown active for '$sender', skipping alert"
                }
            }
        }
    }

} finally {
    Write-Log "Monitor stopped. Releasing resources."
    Release-Outlook
    Stop-Transcript *>$null
}
