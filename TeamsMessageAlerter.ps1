# teams_monitor.ps1
# Monitors Windows notification history for new Teams messages.
# Alerts via Outlook COM -> email-to-SMS gateway.
# No installs required. No log parsing. Works with Teams 2.0 and classic Teams.
#
# KNOWN LIMITATIONS
#   - Focus Assist / Do Not Disturb suppresses toast history entirely.
#     Ensure Focus Assist is OFF on the laptop for reliable alerting.
#   - Teams notification XML varies (chat vs channel vs reaction vs call).
#     Sender/message extraction is best-effort from the first two <text> nodes.
#   - SMS carrier gateways may impose their own rate limits independent of $CooldownSecs.
#
# CARRIER SMS GATEWAYS  (use your 10-digit mobile number)
#   AT&T     : 5551234567@txt.att.net
#   Verizon  : 5551234567@vtext.com
#   T-Mobile : 5551234567@tmomail.net
#
# TASK SCHEDULER (silent startup at login)
#   Trigger : At log on
#   Action  : powershell.exe
#   Args    : -WindowStyle Hidden -ExecutionPolicy Bypass -File "C:\path\to\teams_monitor.ps1"
#   NOTE    : -WindowStyle Hidden is required — omitting it shows a console window at every login.
#
# LOG FILES  (written to $LogDir below)
#   teams_monitor.log            — timestamped event log
#   teams_monitor_transcript.log — full PowerShell transcript

# ══════════════════════════════════════════════════════════════════════════════
# USER CONFIG  — edit these values before first run
# ══════════════════════════════════════════════════════════════════════════════

$script:SmsTo              = "5551234567@vtext.com"  # your number@carrier gateway
$script:PollSeconds        = 10                       # seconds between notification checks
$script:CooldownSecs       = 30                       # min seconds between SMS per sender

# Log directory — must be a path you have write access to.
# Default: folder containing this script, falling back to %USERPROFILE%\teams_monitor
# if $PSScriptRoot is empty (can occur in some Task Scheduler configurations).
$script:LogDir             = if ($PSScriptRoot) { $PSScriptRoot } else { "$env:USERPROFILE\teams_monitor" }

# SMS field limits — carrier gateways enforce a 160-char GSM-7 total.
# Subject and body are independent fields; combined they should not exceed 160.
$script:SmsMaxSubjectChars = 30
$script:SmsMaxBodyChars    = 120

# Truncation suffix appended when a field exceeds its limit.
# Must be ASCII — Unicode chars (e.g. the ellipsis U+2026) force gateways into
# UCS-2 encoding, halving the effective limit from 160 to 70 chars.
$script:TruncationSuffix   = "..."

# Teams AppIds to probe — varies by Teams version installed.
# All are tried at startup; the first that responds without throwing is cached.
$script:TeamsAppIds        = @(
    "MSTeams",
    "com.squirrel.Teams.Teams",
    "MicrosoftTeams_8wekyb3d8bbwe!MSTeams"
)

# Outlook COM constant for a mail item type.
# Ref: https://learn.microsoft.com/en-us/office/vba/api/outlook.olmailitem
$script:OlMailItem         = 0

# Fallback notification ID used when XML hashing fails.
# Fixed constant is intentional — Get-Random would generate a new ID each poll,
# causing the same broken notification to fire an SMS every PollSeconds.
# With a constant, it is seeded once and subsequently deduped.
$script:FallbackNotifId    = "fallback-unparseable"

# Display strings used when notification fields cannot be extracted
$script:DefaultSender      = "Teams"
$script:DefaultMessage     = "(no preview)"
$script:ParseErrorMessage  = "(could not parse notification)"

# Timestamp format used in log entries
$script:LogTimestampFormat = "yyyy-MM-dd HH:mm:ss"

# ══════════════════════════════════════════════════════════════════════════════
# RUNTIME STATE — do not edit
# ══════════════════════════════════════════════════════════════════════════════

$script:WorkingAppId   = $null    # cached Teams AppId, populated after first successful probe
$script:Outlook        = $null    # Outlook COM object handle
$script:LogFile        = "$($script:LogDir)\teams_monitor.log"
$script:TranscriptFile = "$($script:LogDir)\teams_monitor_transcript.log"

# Notification dedup set — stores SHA256 hashes of seen notification XML
$script:Seen           = [System.Collections.Generic.HashSet[string]]::new()

# Per-sender cooldown tracker — stores last successful SMS timestamp per sender name
$script:Cooldown       = [System.Collections.Generic.Dictionary[string,datetime]]::new()

# ══════════════════════════════════════════════════════════════════════════════
# FUNCTIONS
# ══════════════════════════════════════════════════════════════════════════════

function Write-Log {
    param([string]$Message)
    $line = "$(Get-Date -f $script:LogTimestampFormat)  $Message"
    Write-Host $line
    # File write is best-effort — Write-Host above always executes first
    try { Add-Content -Path $script:LogFile -Value $line -Encoding UTF8 } catch {}
}

function Load-WinRT {
    try {
        [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType=WindowsRuntime] | Out-Null
        return $true
    } catch {
        Write-Log "ERROR: Could not load WinRT. Requires Windows 10/11. $_"
        return $false
    }
}

function Get-TeamsNotifications {
    $history = [Windows.UI.Notifications.ToastNotificationManager]::History

    # Fast path — use cached AppId
    if ($script:WorkingAppId) {
        try {
            return $history.GetHistory($script:WorkingAppId)
        } catch {
            Write-Log "Warning: cached AppId '$($script:WorkingAppId)' failed, re-probing. $_"
            $script:WorkingAppId = $null
        }
    }

    # Probe all AppIds. Cache the first one that responds without throwing.
    # An empty result (0 notifications) is valid — it means history is clear, not wrong AppId.
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

function Get-NotificationId {
    # Hashes notification XML to produce a stable unique dedup key.
    # Tag+Group fields are often both empty in Teams and cannot be relied upon.
    param([string]$Xml)
    $hash = $null
    try {
        $hash  = [System.Security.Cryptography.SHA256]::Create()
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($Xml)
        return [Convert]::ToBase64String($hash.ComputeHash($bytes))
    } catch {
        Write-Log "Warning: could not hash notification XML - using fallback ID. Subsequent notifications with the same parse failure will be skipped. $_"
        return $script:FallbackNotifId
    } finally {
        if ($null -ne $hash) { $hash.Dispose() }
    }
}

function Get-TextContent {
    # PowerShell XML shorthand returns a bare [string] for a single matching node,
    # or an [XmlElement] when selected from a multi-node collection.
    # Both must be handled explicitly — [string] has no InnerText property.
    param($Node)
    if ($null -eq $Node)                   { return $null }
    if ($Node -is [string])                { return $Node }
    if ($Node -is [System.Xml.XmlElement]) { return $Node.InnerText }
    return $null
}

function Parse-Notification {
    # Best-effort extraction of sender and message preview from toast XML.
    # Pins to first <binding> element — multiple bindings would cause .text to return null.
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

function Limit-Length {
    param([string]$Text, [int]$Max)
    if ([string]::IsNullOrEmpty($Text))  { return "" }
    if ($Text.Length -le $Max)           { return $Text }
    $suffixLen = $script:TruncationSuffix.Length
    # Guard: if Max is smaller than the suffix, return as much of the suffix as fits
    if ($Max -le $suffixLen)             { return $script:TruncationSuffix.Substring(0, $Max) }
    return $Text.Substring(0, $Max - $suffixLen) + $script:TruncationSuffix
}

function Release-Outlook {
    if ($null -ne $script:Outlook) {
        try {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($script:Outlook) | Out-Null
        } catch {}
        $script:Outlook = $null
    }
}

function Initialize-Outlook {
    Release-Outlook    # always release existing reference before creating a new one
    try {
        $script:Outlook = New-Object -ComObject Outlook.Application
        Write-Log "Outlook COM initialized."
        return $true
    } catch {
        Write-Log "ERROR: Could not connect to Outlook: $_"
        return $false
    }
}

function Send-MailItem {
    # Creates and sends a single Outlook mail item. Throws on any failure.
    param([string]$Subject, [string]$Body)
    $mail         = $script:Outlook.CreateItem($script:OlMailItem)
    $mail.To      = $script:SmsTo
    $mail.Subject = $Subject
    $mail.Body    = $Body
    $mail.Send()
}

function Send-Sms {
    # Returns $true on success, $false on failure.
    # Caller is responsible for only advancing the cooldown on $true.
    param([string]$Subject, [string]$Body)
    $Subject = Limit-Length $Subject $script:SmsMaxSubjectChars
    $Body    = Limit-Length $Body    $script:SmsMaxBodyChars
    try {
        Send-MailItem $Subject $Body
        Write-Log "SMS sent -> $Subject"
        return $true
    } catch {
        Write-Log "Outlook error, attempting reconnect: $_"
        if (Initialize-Outlook) {
            try {
                Send-MailItem $Subject $Body
                Write-Log "SMS sent after reconnect."
                return $true
            } catch {
                Write-Log "Send failed after reconnect: $_"
            }
        }
    }
    return $false
}

# ══════════════════════════════════════════════════════════════════════════════
# STARTUP
# ══════════════════════════════════════════════════════════════════════════════

# Create log directory once at startup — not inside Write-Log on every call
if (-not (Test-Path $script:LogDir)) {
    New-Item -ItemType Directory -Path $script:LogDir -Force | Out-Null
}

# -Append only — do not combine with -NoClobber.
# -NoClobber conflicts with -Append (NoClobber wins), causing the transcript
# to silently fail on every run after the first.
Start-Transcript -Path $script:TranscriptFile -Append *>$null

# Entire startup and runtime is inside one try/finally so Release-Outlook
# and Stop-Transcript always execute, even if startup throws.
try {
    if (-not (Load-WinRT))         { throw "WinRT unavailable - Windows 10/11 required." }
    if (-not (Initialize-Outlook)) { throw "Outlook unavailable - ensure Outlook is open and signed in." }

    Write-Log "Teams Monitor started. Logs: $script:LogDir"
    Send-Sms "Teams Monitor" "Monitor running." | Out-Null

    # Seed dedup set with notifications already in the Action Center so we do
    # not alert on messages that arrived before the monitor started.
    foreach ($n in (Get-TeamsNotifications)) {
        try {
            $xml = $n.Content.GetXml()
            $null = $script:Seen.Add((Get-NotificationId $xml))
        } catch {
            Write-Log "Warning: could not seed notification during startup: $_"
        }
    }
    Write-Log "Seeded $($script:Seen.Count) existing notifications. Watching for new ones..."

    # ── MAIN LOOP ─────────────────────────────────────────────────────────────
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
                    # Only advance cooldown if the SMS was actually delivered.
                    # On send failure the next poll retries immediately rather
                    # than waiting a full cooldown period with no alert sent.
                    if (Send-Sms "Teams: $sender" $message) {
                        $script:Cooldown[$sender] = $now
                    }
                } else {
                    Write-Log "Cooldown active for '$sender', skipping SMS"
                }
            }
        }
    }

} finally {
    Write-Log "Monitor stopped. Releasing resources."
    Release-Outlook
    Stop-Transcript *>$null
}
