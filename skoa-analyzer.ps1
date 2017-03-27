# SkyKick Outlook Assistant Analyzer


# Declare an array to hold the result information
$results = @()

# Log file path
$logfile = "C:\Office365\skoa-analyzer-results.csv"

# Read input file of computer names
$computers =  Get-Content "C:\Office365\computers.txt"


function ComputerIsAccessible
{
    Param([string]$computer)
    If (Test-Path "\\$computer\c$\Windows") { return $true} Else { return $false }
}

Function GetSkyKickProfile
{
    Param([string]$computer)
    # Look for an Outlook profile name ending in "_sk.ost"
    $skProfileName = Get-ChildItem -Path "\\$computer\C$\Users\*\AppData\Local\Microsoft\Outlook\*_sk.ost" -Recurse -Name
    return $skProfileName
}

Function GetSklquery86
{
    Param([string]$computer)
    $target = Get-ChildItem -Path "\\$computer\C$\Program Files (x86)\SkyKick Inc\sklquery\sklquery.exe" -ErrorAction Ignore |
    Select-Object -Property Name, LastWriteTime
    return $target
}

Function GetSklquery64
{
    Param([string]$computer)
    $target = Get-ChildItem -Path "\\$computer\C$\Program Files\SkyKick Inc\sklquery\sklquery.exe" -ErrorAction Ignore |
    Select-Object -Property Name, LastWriteTime
    return $target
}


foreach ($computer in $computers) {

    $result = [PSCustomObject]@{
        ComputerName = $computer
        Online = $false
        SkyKickProfileName = $null
        Sklquery86 = $null
        Sklquery86Modified = $null
        Sklquery64 = $null
        Sklquery64Modified = $null
    }

    If (ComputerIsAccessible $computer)
    {
        $skProfileName = GetSkyKickProfile $computer
        $x86 = GetSklquery86 $computer
        $x64 = GetSklquery64 $computer

        $result.Online = $true
        $result.SkyKickProfileName = $skProfileName
        $result.Sklquery86 = $($x86.Name)
        $result.Sklquery86Modified = $($x86.LastWriteTime)
        $result.Sklquery64 = $($x64.Name)
        $result.Sklquery64Modified = $($x64.LastWriteTime)
    }

    Write-Host $result

    # Add the result to the array of results
    $results += $result
}

# Write the array to a .csv file
$results | Export-Csv $logfile -NoTypeInformation
