$script:config   = Get-Content $PSScriptRoot\config.json | ConvertFrom-Json
$script:template = Get-Content $PSScriptRoot\template.txt -Raw


function New-CWTime
{
    [CmdletBinding()]
    param()

    Write-Information "Remove!"
}

function Remove-CWTime
{
    [CmdletBinding()]
    param()

    Write-Information "Remove!"
}

function Get-CWTime
{
    [CmdletBinding()]
    param(
        [datetime]$Date = (Get-Date)
    )

    Write-Information "Get!"
}

function Clear-CWTime
{
    [CmdletBinding()]
    param(
        [datetime]$Date = (Get-Date)
    )

    Write-Information "Remove!"
}

function Send-CWTimeNote
{
    [CmdletBinding()]
    param(
        [datetime]$Date = (Get-Date)
    )

    Write-Host $Date

    Write-Information "Get!"
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

function Show-CWLog
{
    [CmdletBinding()]
    param()

    Start-Process -FilePath $script:config.TextEditor -ArgumentList "$PSScriptRoot\log.txt"
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