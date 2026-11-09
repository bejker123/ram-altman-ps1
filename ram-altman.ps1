<#
.SYNOPSIS
 ram-altman.ps1 - track RAM usage and, at most once a day, tell Sam Altman
 when this PC has RAM to spare.

.EXAMPLE

 Usage:
   ram-altman.ps1            Log a sample and send the daily email if due
   ram-altman.ps1 -Status   Print current memory figures and exit
   ram-altman.ps1 -DryRun  Print the email that would be sent and exit

 Cron (every 5 minutes):
   */5 * * * * /path/to/ram-altman.ps1
#>

[CmdletBinding(DefaultParameterSetName = "NoArgs")]
param(
    [Parameter(ParameterSetName="Status",HelpMessage="Print current memory figures and exit")]
    [Switch]$Status,
    [Parameter(ParameterSetName="DryRun",HelpMessage="Print the email that would be sent and exit")]
    [Switch]$DryRun,
    [Parameter(ParameterSetName="Help",HelpMessage="Print help and exit")]
    [Switch]$Help
)
# ----- Settings: fill these in -----------------------------------------------
$RECIPIENT=$RAM_ALTMAN_TO ?? ""             # Who receives the email
$SENDER=$RAM_ALTMAN_FROM ?? ""              # A real address you can send from
$SENDER_NAME=$RAM_ALTMAN_FROM_NAME ?? ""    # Display name, optional
$THRESHOLD_PCT=$RAM_ALTMAN_THRESHOLD ?? 10 # Email when free is above this
$SMTP_SERVER = $RAM_ALTMAN_SMTP ?? ""
# -----------------------------------------------------------------------------

$STATE_DIR=(($RAM_ALTMAN_STATE_DIR ?? $XDG_STATE_HOME) ?? $HOME+"\.local/state")+"\ram-altman"
$USAGE_LOG="$STATE_DIR/usage.csv"
$LAST_SENT_FILE="$STATE_DIR/last_sent"
$MAX_LOG_BYTES=10 * 1MB   # Trim the log once it passes 10 MiB
$LOG_LINES_TO_KEEP=50000

function GetMemoryInfo{

    $script:OS = Get-CimInstance Win32_OperatingSystem
    $script:total_gb = $OS.TotalVisibleMemorySize / 1MB
    $script:free_gb = $OS.FreePhysicalMemory /1MB
    $script:used_gb = $total_gb - $free_gb

    $script:free_pct = 100 * $free_gb / $total_gb
}

function IsOverThreshold{
    ($script:free_pct) -gt $THRESHOLD_PCT
}

function FormatGiB($a){
    "$([math]::Round($a,1)) GiB" 
}

function PrintStatus {
    Write-Host 'Total:      ' "$(FormatGiB($total_gb))"
    Write-Host 'Used:       ' "$(FormatGiB($used_gb))"
    Write-Host 'Free:  ' "$(FormatGiB($free_gb))" "($([math]::Round($free_pct,2))%)"
}

function LogSample{
    $firstLine = "timestamp,total_gb,used_gb,free_gb,free_pct"
    if (-not(Test-Path $USAGE_LOG)){
        Set-Content -Path $USAGE_LOG -Value $firstLine
        Write-Debug "Created log file ($(Resolve-Path $USAGE_LOG))"
    }
    $content = Get-Content $USAGE_LOG | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    if (($content.Length -eq 0)-or ($content[0] -ne $firstLine)){
        Write-Debug "Repaired first line of log file ($(Resolve-Path $USAGE_LOG))"
        Set-Content -Path $USAGE_LOG -Value $firstLine
        Add-Content -Path $USAGE_LOG -Value $content
    }
    $timestamp = Get-Date -UFormat "+%Y-%m-%dT%H:%M:%S%Z"
    $out = ""
    @($total_gb,$used_gb,$free_gb,$free_pct) | ForEach-Object { $out += ",$([math]::Round($_,2))" } 

    Add-Content -Path $USAGE_LOG -Value "$timestamp$out"

    if((Get-ChildItem $USAGE_LOG).Length -gt $MAX_LOG_BYTES){
        $log = Get-Content $USAGE_LOG
        Set-Content -Path $USAGE_LOG -Value $log[0]+$log[[math]::Max(1,$log.Length - $LOG_LINES_TO_KEEP),$log.Length]
    }
}
function Pick{
    param (
        [string]$Name,
        [int]$Count
    )

    # Dry run check
    if ($DryRun) {
        return (Get-Random -Minimum 0 -Maximum $Count)
    }

    $bag = Join-Path $STATE_DIR "bag_$Name"
    $bagLast = "$bag.last"

    # If the file doesn't exist or is empty
    if (-not (Test-Path $bag) -or (Get-Item $bag).Length -eq 0) {
        # Generate 0 to count-1 and shuffle
        $shuffled = 0..($Count - 1) | Get-Random -Count $Count
        
        # Stop the last pick of the old round opening the new one
        $last = if (Test-Path $bagLast) { Get-Content $bagLast -Raw } else { $null }
        
        if ($shuffled[0] -eq $last -and $Count -1) {
            # Move the first item to the end
            $shuffled = $shuffled[1..($shuffled.Length - 1)] + $shuffled[0]
        }
        $shuffled | Set-Content $bag
    }

    # Read the content of the bag
    $content = Get-Content $bag
    
    # Handle single element vs array behavior in PowerShell
    if ($content -is [array]) {
        $next = $content[0]
        $rest = $content[1..($content.Length - 1)]
    } else {
        $next = $content
        $rest = @()
    }

    # Update the bag file on disk
    if ($rest) {
        $rest | Set-Content $bag
    } else {
        Clear-Content $bag
    }
    
    $next | Set-Content $bagLast

    # Covers the template lists being shortened after the bag was dealt
    if ($next -ge $Count) {
        $next = Get-Random -Minimum 0 -Maximum $Count
    }

    return $next
}

function ComposeEmail{
    $total=FormatGiB($total_gb)
    $free=FormatGiB($free_gb)
    $pct=$([math]::Round($free_pct,2))

    $subjects=@(
        "Good news: we found some RAM",
        "Unused RAM detected in the wild",
        "RAM sighting on my PC",
        "$free of RAM, just sitting there",
        "You missed a spot",
        "Daily RAM census: some survived",
        "$pct% of my RAM remains at large"
        "Still not all of it, I'm afraid"
    )

    $bodies=(
        "I'm writing to report that you appear to have missed some. This PC has $total of RAM and $free of it ($pct%) is currently available. I assume this is an oversight.",
        "A routine inspection of my PC found $free of unclaimed RAM out of $total. That's $pct% of the machine doing absolutely nothing. I thought you'd want to know before someone tries to train something on it.",
        "Daily status report:

  Total RAM:     $total
  Available:     $free ($pct%)
  Threat level:  someone at OpenAI might notice",
        "Even with every browser tab I own open, $free of my $total of RAM is still available. I'm as surprised as you are.",
        "Please don't be alarmed, but there is $free of RAM on my desk that isn't in a data centre. That's $pct% of what's installed here, and as far as I can tell it has no plans today.",
        "Before you ask: yes, I counted the page cache as available. The kernel gives it back the moment anyone asks, which leaves $free of $total up for grabs.",
        "Today's forecast for my PC: clear skies, light workloads and $pct% available memory. Expect $free of the installed $total to stay idle through the evening.",
        "I hear RAM is hard to come by these days, so it feels a bit rude to be sitting on $free of it. Total installed: $total. It's not for sale, I just like knowing it's there.",
        "I have a confession. I have $free of RAM that I'm not using, which is $pct% of the $total in this PC. I've told nobody but you and a cron job.",
        "If $free of RAM sits idle in a PC and nobody allocates it, does it still count as capacity? Asking on behalf of the $pct% of my memory with nothing to do."
    )

    $signoffs=(
        "Yours in abundance,",
        "Regards from the last free RAM in the county,",
        "Hoarding responsibly,",
        "Stay allocated,",
        "Yours, with memory to spare,",
        "Kind regards,"
    )

    $script:email_subject=$subjects[$(Pick -Name "subject" -Count $($subjects.Length))]
    $script:email_body="Hi Sam,

$($bodies[$(Pick -Name "body" -Count $($bodies.Length))])

$($signoffs[$(Pick -Name signoff -Count $($bodies.Length))])
$($SENDER_NAME ?? "A concerned RAM owner")

--
Sent automatically, at most once a day, by a small script on one PC.
Reply if you'd like it to stop."

}
function RenderMessage{
    $from=$SENDER
    if($SENDER_NAME){
        $from = "$SENDER_NAME <$SENDER>"
    }
    
    @(
        "From: $from",
        "To: $RECIPIENT",
        "Subject: $email_subject",
        "Date: $(Get-Date -Format "ddd, d MMM yyyy HH:mm:ss zz")",
        "MIME-Version: 1.0",
        "Content-Type: text/plain; charset=UTF-8",
        "Auto-Submitted: auto-generated`n",
        $email_body
    )
}

function SendEmail{
    if (Get-Command Send-MailMessage -ErrorAction SilentlyContinue) {
    # Render your body and subject (assuming these variables are set)
    $emailBody = Get-RenderMessage 
    
    Send-MailMessage -To $RECIPIENT `
                     -Subject $EMAIL_SUBJECT `
                     -Body $emailBody `
                     -SmtpServer $SMTP_SERVER
    }
    elseif (Get-Command sendmail -ErrorAction SilentlyContinue) {
        Get-RenderMessage | sendmail -t -i
    }
    elseif (Get-Command msmtp -ErrorAction SilentlyContinue) {
        Get-RenderMessage | msmtp -t
    }
    elseif (Get-Command mail -ErrorAction SilentlyContinue) {
        $EMAIL_BODY | mail -s $EMAIL_SUBJECT $RECIPIENT
    }
    # 3. Error Catch-all
    else {
        Write-Error "ram-altman: no native Send-MailMessage, sendmail, msmtp or mail command found"
        return 1
    }
}

function PrintHelp {
    Get-Help $script:MyInvocation.MyCommand.Definition
}
if($Status){
    GetMemoryInfo
    PrintStatus
    exit
} elseif ($DryRun) {
    GetMemoryInfo
    PrintStatus
    if(!(IsOverThreshold)){
        Write-Host "No email: free RAM is not above $THRESHOLD_PCT%."
    }
    ComposeEmail
    RenderMessage
    exit
}elseif($Help){
    PrintHelp
    exit
}
else{
    $quit = $false
    if([String]::IsNullOrWhiteSpace($RECIPIENT)){
        Write-Error "RECIPIENT is not set"
        $quit = $true
    }
    
    if([String]::IsNullOrWhiteSpace($SENDER)){
        Write-Error "SENDER is not set"
        $quit = $true
    }
    if($quit){
        exit 1
    }

    New-Item -ItemType Directory -Path $STATE_DIR -Force | Out-Null
    
    GetMemoryInfo
    LogSample

    if(-not(IsOverThreshold)){
        exit
    }

    $today = $(Get-Date -UFormat "+%F")
    if ((Test-Path $LAST_SENT_FILE) -and ($today -eq ((Get-Content $LAST_SENT_FILE) ?? ""))){
        exit
    }

    ComposeEmail
    if (-not (SendEmail)){
        Write-Error "email failed to send; will try again on the next run"
        exit 1
    }

    # Recorded only after a successful send, so a failure gets retried.
    Set-Content -Path $LAST_SENT_FILE -Value $today
}
