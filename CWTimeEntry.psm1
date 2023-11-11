$script:config   = Get-Content $PSScriptRoot\config.json | ConvertFrom-Json
$script:template = Get-Content $PSScriptRoot\template.txt -Raw
$script:AuthString  = ($script:config.CWconfig.companyId + '+' + $script:config.CWconfig.API.publickey + ':' + $script:config.CWconfig.API.privatekey)
$script:EncodedAuth = ([Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes($script:AuthString)))

class TimeEntry
{
    [string]$TicketId
    [datetime]$TimeStart
    [datetime]$TimeEnd
    [string]$Notes
    [string]$BillableOption
}

function New-CWTime
{
    [cmdletbinding()]
    param(
        [Parameter(mandatory=$true)]
        [TimeEntry]$TimeEntry
    )

    $Body = @{
        chargeToType    = 'ServiceTicket'
        chargeToId      = $TimeEntry.TicketId
        timeStart       = $TimeEntry.TimeStart.ToString("yyyy-MM-ddTHH:mm:ssZ")
        timeEnd         = $TimeEntry.TimeEnd.ToString("yyyy-MM-ddTHH:mm:ssZ")
        notes           = $TimeEntry.Notes
        member          = @{identifier = $script:config.CWconfig.memberIdentifier}
        workRole        = @{name       = $script:config.CWconfig.workRole}
        billableOption = $TimeEntry.BillableOption
    } | ConvertTo-Json
    
    $Headers= @{
        "Authorization" = "Basic $script:EncodedAuth"
        "clientId"      = $script:config.CWconfig.API.clientId
        "Content-Type"  = "application/json"
    }

    $Response = Invoke-RestMethod -Method Post -Uri 'https://na.myconnectwise.net/v4_6_release/apis/3.0/time/entries' -Headers $Headers -Body $Body

    return [PSCustomObject]@{
        'BillableOption' = $Response.billableOption
        'TimeStart'      = ([datetime]$Response.timeStart).ToLocalTime()
        'TimeEnd'        = ([datetime]$Response.timeEnd).ToLocalTime()
    }
}

function Get-CWTime
{
    [cmdletbinding()]
    param(
        [datetime]$date = (Get-Date)
    )
    
    $TodayDateTime       = (Get-Date -Year $date.Year -Month $date.Month -Day $date.Day -Hour 0 -Minute 0 -Second 0).ToUniversalTime()
    $StartDateTimeString = $TodayDateTime.ToString("yyyy-MM-ddTHH:mm:ssZ")
    $EndDateTimeString   = $TodayDateTime.AddDays(1).ToString("yyyy-MM-ddTHH:mm:ssZ")
    
    $MemberIdentifier = $script:config.CWconfig.memberIdentifier
    $Uri = "https://na.myconnectwise.net/v4_6_release/apis/3.0/time/entries?pageSize=1000&conditions=member/identifier='$MemberIdentifier' and timeStart >= [$StartDateTimeString] and timeStart <= [$EndDateTimeString]"
    
    $Headers = @{
        "Authorization" = "Basic $script:EncodedAuth"
        "Content-Type"  = "application/json"
        "clientId"      = $script:config.CWconfig.API.clientId
    }
    
    $TargetTimeEntryList = Invoke-RestMethod -Method Get -Uri $Uri -Headers $Headers
    
    return ($TargetTimeEntryList | Select-Object id,billableOption,@{n='timeStart';e={([datetime]$_.timeStart).ToLocalTime()}},@{n='timeEnd';e={([datetime]$_.timeEnd).ToLocalTime()}},notes)
}

function Clear-CWTime
{
    [cmdletbinding()]
    param(
        [datetime]$date = (Get-Date)
    )
    
    $TimeEntryList = Get-CWTime $date

    foreach($TimeEntry in $TimeEntryList)
    {
        $TimeEntry | Select-Object timeStart,timeEnd,notes
        Remove-CWTime $TimeEntry.id
    }
}
function Remove-CWTime
{
    [cmdletbinding()]
    param(
        [Parameter(mandatory=$true)]
        [string]$TimeEntryId
    )
    
    $Headers = @{
        "Authorization" = "Basic $script:EncodedAuth"
        "Content-Type"  = "application/json"
        "clientId"      = $script:config.CWconfig.API.clientId
    }
    
    $Response = Invoke-RestMethod -Method Delete -Uri "https://na.myconnectwise.net/v4_6_release/apis/3.0/time/entries/$TimeEntryId" -Headers $Headers

    return
}

function Send-CWTimeNote
{
    param(
        [datetime]$Date = (Get-Date)
    )
    
    $datestring = $date.ToString("yyyy-MM-dd")
    $TimeNotePath = "$PSScriptRoot\TimeNotes\$datestring.txt"
    $TimeNoteText = Get-Content -Path $TimeNotePath -Raw
    
    $TimeEntryTextList = ($TimeNoteText -split "\r\n\r\n(?:\r\n)*")
    
    $TimeEntryList = [System.Collections.Generic.List[TimeEntry]]@()
    foreach($TimeEntryText in $TimeEntryTextList[0])
    {
        $TimeEntrySplit = $TimeEntryText -split "`n"
        $TimeText = $TimeEntrySplit[0].Trim()
        $TicketText = $TimeEntrySplit[1].Trim()
        $NotesText = $TimeEntrySplit[2..($TimeEntrySplit.Count)] -join "`n"

        if($TimeText -notmatch '^(?<StartTime>([1-9]|10|11|12)(:(15|30|45))?(a|p)) - (?<EndTime>([1-9]|10|11|12)(:(15|30|45))?(a|p)) ?(?<BillableOption>(NB|NC))?$')
        {
            Write-Warning "Invalid time entry. Time is not formatted properly.`n$TimeEntryText"
            return
        }

        $TimeStart      = ([datetime]"$datestring $($Matches.StartTime)m").ToUniversalTime()
        $TimeEnd        = ([datetime]"$datestring $($Matches.EndTime)m").ToUniversalTime()
        $BillableOption = switch ($Matches.BillableOption) {
            'NB'    {'DoNotBill'}
            'NC'    {'NoCharge'}
            default {'Billable'}
        }

        if($TicketText -notmatch '^(?<TicketId>\d+) (?<Client>[^-]+) - (?<TicketDescription>[\s\S]+)')
        {
            Write-Warning "Invalid time entry. Ticket is not formatted properly.`n$TimeEntryText"
            return
        }
    
        $TicketId       = $Matches.TicketId
        $Notes          = "$($Matches.TicketDescription)`n$($NotesText)"
        
        if($TimeEnd -le $TimeStart)
        {
            Write-Warning "Invalid time entry. Start time must be after end time.`n$TimeEntryText"
            return
        }
    
        if(($TimeEnd - $TimeStart).TotalHours -gt 8)
        {
            Write-Warning "WARNING: Time entry is over 8 hours long.`n$TimeEntryText"
        }
    
        $TimeEntryList.Add([TimeEntry]@{
            'TicketId'       = $TicketId
            'TimeStart'      = $TimeStart
            'TimeEnd'        = $TimeEnd
            'Notes'          = $Notes
            'BillableOption' = $BillableOption
        })
    }

    foreach($TimeEntry in $TimeEntryList)
    {
        New-CWTime -TimeEntry $TimeEntry
    }
}

function Show-CWTimeNote
{
    [CmdletBinding()]
    param(
        [datetime]$Date = (Get-Date)
    )

    $datestring = $date.ToString("yyyy-MM-dd")
    $TimeNotePath = "$PSScriptRoot\TimeNotes\$datestring.txt"

    if(-not (Test-Path $TimeNotePath))
    {
        New-Item -ItemType File -Path $TimeNotePath -Force | Out-Null
        $script:template > $TimeNotePath
    }

    Start-Process -FilePath $script:config.TextEditor -ArgumentList $TimeNotePath
}

function Show-CWConfig
{
    [CmdletBinding()]
    param()

    Start-Process -FilePath $script:config.TextEditor -ArgumentList "$PSScriptRoot\config.json"
}

function Edit-CWConfig
{
    while ([string]::IsNullOrEmpty($UserId))
    {
        $UserId = Read-Host "Please enter your CW userId, e.g. jdoe"
    }

    while ([string]::IsNullOrEmpty($WorkRole))
    {
        $WorkRole = Read-Host "Please enter your CW Work Role"
    }

    while ([string]::IsNullOrEmpty($companyId))
    {
        $companyId = Read-Host "Please enter your CW Company Id"
    }

    while ([string]::IsNullOrEmpty($publickey))
    {
        $publickey = Read-Host "Please enter your CW API Public Key"
    }

    while ([string]::IsNullOrEmpty($clientId))
    {
        $clientId = Read-Host "Please enter your CW API Client Id"
    }

    while ([string]::IsNullOrEmpty($privatekey))
    {
        $privatekey = Read-Host "Please enter your CW API Private Key"
    }

    # Setup config file
    try
    {
        $config = Get-Content "$PSScriptRoot\config.json" | ConvertFrom-Json

        Write-Host "Creating config file..."
        $config.CWconfig.memberIdentifier = $UserId
        $config.CWconfig.workRole         = $WorkRole
        $config.CWconfig.companyId        = $companyId
        $config.CWconfig.API.publickey    = $publickey
        $config.CWconfig.API.clientId     = $clientId
        $config.CWconfig.API.privatekey   = $privatekey

        # Write the config file
        $config | ConvertTo-Json -Depth 50 > "$PSScriptRoot\config.json"
    }
    catch
    {
        Write-Warning "Error creating config file."
        throw $_
    }

    Write-Host "Installation complete. Please Restart PowerShell session for changes to take effect."
}

function Show-CWTickets
{
    [cmdletbinding()]
    param()

    Start-Process -FilePath $script:config.TextEditor -ArgumentList "$PSScriptRoot\tickets.txt"
}

function Show-CWTemplate
{
    [cmdletbinding()]
    param()

    Start-Process -FilePath $script:config.TextEditor -ArgumentList "$PSScriptRoot\template.txt"
}

function Measure-CWTimeNote
{
    [cmdletbinding()]
    param
    (
        [datetime]$date = (get-date)
    )

    $datestring = $date.ToString("yyyy-MM-dd")
    $TimeNotePath = "$PSScriptRoot\TimeNotes\$datestring.txt"
    $TimeNoteText = Get-Content -Path $TimeNotePath -Raw
    
    $TimeEntryTextList = ($TimeNoteText -split "\r\n\r\n(?:\r\n)*")
    
    $TimeEntryList = [System.Collections.Generic.List[TimeEntry]]@()
    foreach($TimeEntryText in $TimeEntryTextList[0])
    {
        $TimeEntrySplit = $TimeEntryText -split "`n"
        $TimeText = $TimeEntrySplit[0].Trim()
        $TicketText = $TimeEntrySplit[1].Trim()
        $NotesText = $TimeEntrySplit[2..($TimeEntrySplit.Count)] -join "`n"

        if($TimeText -notmatch '^(?<StartTime>([1-9]|10|11|12)(:(15|30|45))?(a|p)) - (?<EndTime>([1-9]|10|11|12)(:(15|30|45))?(a|p)) ?(?<BillableOption>(NB|NC))?$')
        {
            Write-Warning "Invalid time entry. Time is not formatted properly.`n$TimeEntryText"
            return
        }

        $TimeStart      = ([datetime]"$datestring $($Matches.StartTime)m").ToUniversalTime()
        $TimeEnd        = ([datetime]"$datestring $($Matches.EndTime)m").ToUniversalTime()
        $BillableOption = switch ($Matches.BillableOption) {
            'NB'    {'DoNotBill'}
            'NC'    {'NoCharge'}
            default {'Billable'}
        }

        if($TicketText -notmatch '^(?<TicketId>\d+) (?<Client>[^-]+) - (?<TicketDescription>[\s\S]+)')
        {
            Write-Warning "Invalid time entry. Ticket is not formatted properly.`n$TimeEntryText"
            return
        }
    
        $TicketId       = $Matches.TicketId
        $Notes          = "$($Matches.TicketDescription)`n$($NotesText)"
        
        if($TimeEnd -le $TimeStart)
        {
            Write-Warning "Invalid time entry. Start time must be after end time.`n$TimeEntryText"
            return
        }
    
        if(($TimeEnd - $TimeStart).TotalHours -gt 8)
        {
            Write-Warning "WARNING: Time entry is over 8 hours long.`n$TimeEntryText"
        }
    
        $TimeEntryList.Add([TimeEntry]@{
            'TicketId'       = $TicketId
            'TimeStart'      = $TimeStart
            'TimeEnd'        = $TimeEnd
            'Notes'          = $Notes
            'BillableOption' = $BillableOption
        })
    }

    return [PSCustomObject]@{
        'Total'       = ($TimeEntryList | Select-Object -Property @{e={($_.TimeEnd - $_.TimeStart).TotalHours}} | Measure-Object -sum -Property *).Sum
        'Billable'    = ($TimeEntryList | Where-Object {$_.BillableOption -like 'Billable'} | Select-Object -Property @{e={($_.TimeEnd - $_.TimeStart).TotalHours}} | Measure-Object -sum -Property *).Sum
        'Non-Billable' = ($TimeEntryList | Where-Object {$_.BillableOption -like 'DoNotBill'} | Select-Object -Property @{e={($_.TimeEnd - $_.TimeStart).TotalHours}} | Measure-Object -sum -Property *).Sum
        'NoCharge' = ($TimeEntryList | Where-Object {$_.BillableOption -like 'NoCharge'} | Select-Object -Property @{e={($_.TimeEnd - $_.TimeStart).TotalHours}} | Measure-Object -sum -Property *).Sum
    }
}