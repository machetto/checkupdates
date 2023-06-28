#requires -version 3
<#
.SYNOPSIS
    The script scans remote computers (via COM's call to 'Microsoft.Update.Session') in specified Active Directory OUs for Windows Updates and reports whether they are compliant.
    Computers can be excluded from a scan (use 'exception_list.txt') and/or included in a scan (use 'inclusion_list.txt').
    Reports will saved into a '.\reports' subfolder and also optionaly emailed if receipents are specified.
.EXAMPLE 
    C:\Windows\System32\WindowsPowerShell\v1.0\powershell -NoProfile -Command .\Check-Servers-Updates.ps1 *> logfile.log
#>

$number_of_parallel_jobs = 10 # optimal for 2 x vCPU and 8GB RAM VM

# default input variables
$search_bases = "OU=Servers,OU=YOUROU,DC=YOURDC,DC=,DC=CORP", "OU=Domain Controllers,DC=YOURDC,DC=CORP"

#specify an email here or in the *-Variables.ps1 file to receive the report in an email
$recepient_email = ""

function Get-ScriptDirectory {
    if ($psise) {Split-Path $psise.CurrentFile.FullPath}
    else {
         if ($global:PSScriptRoot) { $global:PSScriptRoot }
         else { $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath('.\') }
    }
}

$exec_folder = Get-ScriptDirectory

#Write-Host -NoNewLine 'Press any key to continue...';
#$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');

# here we dot source external variables that will override local ones
If ( Test-Path .\Check-Servers-Updates-Variables.ps1 ) {
 . $exec_folder\Check-Servers-Updates-Variables.ps1
}

If ( ! (Test-Path .\Reports) ) {
    New-Item -Path $exec_folder -Name "Reports" -ItemType "directory"
}

function Get-PatchTue 
{ 
    param(
    $month = (get-date).month, 
    $year = (get-date).year
    ) 

    $firstdayofmonth = Get-Date -Day 1 -Month $month -Year $year
    (0..30 | ForEach-Object {$firstdayofmonth.adddays($_) } | Where-Object {$_.dayofweek -like "Tue*"})[1]
}

if ( $(Get-Date) -gt $(Get-PatchTue) )
{
    $date_colours = @{
                        "$(Get-Date ((Get-Date).AddDays(-0)) -format "yyyy-MM")"="green"
                        "$(Get-Date ((Get-Date).AddDays(-30)) -format "yyyy-MM")"="black"
                        "$(Get-Date ((Get-Date).AddDays(-60)) -format "yyyy-MM")"="red"
                    }
}
else
{
    $date_colours = @{
                        "$(Get-Date ((Get-Date).AddDays(-0)) -format "yyyy-MM")"="green"
                        "$(Get-Date ((Get-Date).AddDays(-30)) -format "yyyy-MM")"="green"
                        "$(Get-Date ((Get-Date).AddDays(-60)) -format "yyyy-MM")"="black"
                        "$(Get-Date ((Get-Date).AddDays(-90)) -format "yyyy-MM")"="red"
                    }
}

$servers =  $search_bases | ForEach-Object { Get-ADComputer -prop * -SearchBase $_ -filter { (OperatingSystem -Like "*Windows*") -and (Enabled -eq $True) } } | select-object -ExpandProperty Name

$servers = $servers.ToUpper() | Sort-Object

If ( Test-Path $exec_folder\exception_list.txt ) {
$exception_list = Get-Content $exec_folder\exception_list.txt
}

If ( Test-Path $exec_folder\inclusion_list.txt ) {
$inclusion_list = Get-Content $exec_folder\inclusion_list.txt 
}

$all_servers = $servers + $inclusion_list
$all_servers = $all_servers | Where-Object { -not ($exception_list -contains $_) }

if ($null -eq $all_servers) {
	Write-Output "$(Get-Date -format "yyyy-MM-dd HH:mm:ss") No computers found. Exiting."
	Exit(1)
}

Write-Output "$(Get-Date -format "yyyy-MM-dd HH:mm:ss") $($all_servers.Count) computers found in $search_bases and will be checked"

$job_script =
{
    param( [Parameter(Mandatory = $true)]
        [string] $ComputerName
    )
    
    $pattern1 = "Monthly Quality Rollup for Windows"
    #example: 2019-08 Security Monthly Quality Rollup for Windows Server 2012 R2 for x64-based Systems (KB4512488)

    $pattern2 = "Cumulative Update for Windows"
    #example: 2019-08 Cumulative Update for Windows 10 Version 1709 for x64-based Systems (KB4512516)

    $domain_name = (Get-ADDomain).DNSRoot

    function Convert-WuaResultCodeToName {
        param( [Parameter(Mandatory = $true)]
            [int] $ResultCode
        )
        $Result = $ResultCode
        switch ($ResultCode) {
            0 {
                $Result = "Not Started"
            }
            1 {
                $Result = "In Progress"
            }
            2 {
                $Result = "Succeeded"
            }
            3 {
                $Result = "Succeeded With Errors"
            }
            4 {
                $Result = "Failed"
            }
            5 {
                $Result = "Aborted"
            }
        }
        return $Result
    }

    function Get-WuaHistory {
        param( [Parameter(Mandatory = $true)]
            [string] $ComputerName
        )

        # make sure the remote PC has its ports open. Otherwise enable "COM+ Remote Administration (DCOM-In)" firewall rule
        Try
        {
            Write-Verbose "Creating COM object for WSUS Session"
            $objSession = [activator]::CreateInstance([type]::GetTypeFromProgID("Microsoft.Update.Session", $ComputerName))
            $objSearcher = $objSession.CreateUpdateSearcher()
            $totalupdates = $objSearcher.GetTotalHistoryCount()
            $history = $objSearcher.QueryHistory(0, $totalupdates)
        }

        Catch
        {
            Write-Warning "$($Error[0])"
            Write-Verbose "Unable to create a COM object for WSUS Session, exiting"
            exit 1
        }

        #or
        #$session = (New-Object -ComObject 'Microsoft.Update.Session')
        #$history = $session.QueryHistory("",0,50)
        #or
        #$history = $objSearcher.GetTotalHistoryCount()

        $history | Where-Object { ![string]::IsNullOrWhitespace($_.title) } | ForEach-Object {

            $update_object = New-Object -TypeName PSObject
            $update_object | Add-Member -Name 'Result' -MemberType Noteproperty -Value ''
            $update_object | Add-Member -Name 'Date' -MemberType Noteproperty -Value ''
            $update_object | Add-Member -Name 'Title' -MemberType Noteproperty -Value ''
            $update_object | Add-Member -Name 'SupportUrl' -MemberType Noteproperty -Value ''
            $update_object | Add-Member -Name 'Product' -MemberType Noteproperty -Value ''
            $update_object | Add-Member -Name 'UpdateId' -MemberType Noteproperty -Value ''
            $update_object | Add-Member -Name 'RevisionNumber' -MemberType Noteproperty -Value ''

            $update_object.Result = Convert-WuaResultCodeToName -ResultCode $_.ResultCode
            
            $update_object.UpdateId = $_.UpdateIdentity.UpdateId
            
            $update_object.RevisionNumber = $_.UpdateIdentity.RevisionNumber
            
            $update_object.SupportUrl = $_.SupportUrl

            $update_object.Title = $_.Title

            $update_object.Date = $_.Date
            
            $update_object.Product = $_.Categories | Where-Object { $_.Type -eq 'Product' } | Select-Object -First 1 -ExpandProperty Name
            
            $update_object
        }
    }

    $computer_object = New-Object -TypeName PSObject
    $computer_object | Add-Member -Name 'Name' -MemberType Noteproperty -Value $ComputerName
    $computer_object | Add-Member -Name 'PoweredOn' -MemberType Noteproperty -Value ''
    $computer_object | Add-Member -Name 'Updated' -MemberType Noteproperty -Value ''
    $computer_object | Add-Member -Name 'ScannedOn' -MemberType Noteproperty -Value ''

    $computer_object.ScannedOn = Get-Date
    $computer_object.PoweredOn = "Online"

    If ( ! (Test-Connection $ComputerName -Count 1 -Quiet) )
    {
        $computer_object.PoweredOn = "Offline"
        return $computer_object
    }

    $latest_succssess_updates = Get-WuaHistory -ComputerName "$ComputerName.$domain_name" | ? { ($_.Date -gt ((Get-Date).AddMonths(-6))) -and ($_.Result -eq "Succeeded") }
    
    #$latest_succssess_updates | % { $match = [regex]::Match($_.Title,"([\d-]+) $pattern1.*$").Groups[1]; if ($match.Success -eq $True ) { $match.Value } }
    #$latest_succssess_updates | % { [regex]::Match($_.Title,"^(\d\d\d\d-\d\d).*Monthly Quality Rollup for Windows Server.*$").Groups[1].Value }
    #$latest_succssess_updates | % { $m = [regex]::Match($_.Title,"^(\d\d\d\d-\d\d).*Monthly Quality Rollup for Windows Server.*$").Groups[1].Captures; if ($m.Success -eq $True ) { $m.Value } }
    
    $detected_dates = @()
    $detected_dates += $latest_succssess_updates | ForEach-Object { $m = [regex]::Match($_.Title, "^(\d\d\d\d-\d\d).*$($pattern1).*$").Groups[1]; if ($m.Success -eq $True ) { $m.Value } }
    $detected_dates += $latest_succssess_updates | ForEach-Object { $m = [regex]::Match($_.Title, "^(\d\d\d\d-\d\d).*$($pattern2).*$").Groups[1]; if ($m.Success -eq $True ) { $m.Value } }
    $computer_object.Updated = $detected_dates | Sort-Object | Select-Object -Last 1

    $computer_object
}

#remove older jobs if they exist
Get-Job | Remove-Job -Force

#run a job per each server
$all_servers | ForEach-Object {

    Start-Job -Name $_ -ScriptBlock $job_script -ArgumentList $_

    while ( (Get-Job | Where-Object { $_.State -eq 'Running' }).Count -ge $number_of_parallel_jobs ) {
        Start-Sleep -Seconds 10
        Write-Output "$(Get-Date -format "yyyy-MM-dd HH:mm:ss") Pausing to get the number of jobs below $number_of_parallel_jobs"
    }
}

#report jobs' status every 5 seconds
Do {
    Start-Sleep -Seconds 5
    $jobcount = (Get-Job | Where-Object { $_.State -eq 'Running' }).Count
    Write-Output "$(Get-Date -format "yyyy-MM-dd HH:mm:ss") Number of jobs currently running: $($jobcount)"
} Until ($jobcount -le 10)

#track long running outliers
Do {
    Start-Sleep -Seconds 5
    $jobnames = Get-Job | Where-Object { $_.State -eq 'Running' } | Select-Object Name -ExpandProperty Name
    $jobcount = (Get-Job | Where-Object { $_.State -eq 'Running' }).Count
    Write-Output "$(Get-Date -format "yyyy-MM-dd HH:mm:ss") Jobs currently running: $($jobnames)"
} Until ($jobcount -eq 0)

$jobs = Get-Job
$jobs_status = $jobs | Select-Object -Property Name, @{L = 'Totaltime'; E = { [math]::Round(($_.psendtime – $_.psbegintime).TotalSeconds) } } | Sort-Object TotalTime -Descending | Select-Object -first 5
Write-Output "Top 5 longest running jobs, in seconds"
$jobs_status | Format-Table
Write-Output "$(Get-Date -format "yyyy-MM-dd HH:mm:ss") $($jobs.Count) jobs completed"

$received_data = Get-Job | Receive-Job

#
# Build the HTML output
#
$Head = "<title>Servers Status Report</title>"
$Body = @()
$Style = @"

<style type="text/css">
body {
font-family: Verdana, Geneva, Arial, Helvetica, sans-serif;
}
       tr:nth-child(odd) {background-color: #f2f2f2;}

       table{
       border-collapse: collapse;
       border: 1px solid black;
       font: 10pt Verdana, Geneva, Arial, Helvetica, sans-serif;
       color: black;
       margin-bottom: 10px;
}

       p {
       font: 10pt Verdana, Geneva, Arial, Helvetica, sans-serif;
       color: black;
       margin-bottom: 10px;
}

    table td{
       font-size: 12px;
       border: 1px solid black;
       padding-left: 10px;
       padding-right: 10px;
       padding-bottom: 5px;
       padding-top: 5px;

}

    table th {
       background-color: LIGHTGRAY;
       font-size: 12px;
       border: 1px solid black;
       font-weight: bold;
       padding-left: 10px;
       padding-right: 10px;
       padding-bottom: 10px;
       padding-top: 10px;
       text-align: left;
}
</style>

"@

$Head += $Style
$Body += "<table><tr><th>Date/Time</th><th>Server Name</th><th>Connectivity status</th><th>Latest Windows Update</th></tr>"
$Body += $received_data | ForEach-Object {
    If ($_.PoweredOn -eq "Online") {
        "<tr><td>$($_.ScannedOn)  </td><td>$($_.Name)</td><td><font color='green'>Online</font></td>"
    }
    Else {
        "<tr><td>$($_.ScannedOn)  </td><td>$($_.Name)</td><td><font color='red'>Offline</font></td>"
    }

    $colour = $date_colours[$_.Updated]
    If ( $colour -eq $null ) { $colour = "red" }

    "<td><font color=$colour>$($_.Updated)</font></td></tr>"
}


$Body += "</table>"

$Html = ConvertTo-Html -Body $Body -Head $Head

$Html > "$($exec_folder)\Reports\Report-$(Get-Date -format "yyyy-MM-dd-HH-mm-ss").html"

$messageParameters = @{                         
    Subject    = "Servers' Status Report / $(Get-Date -format "yyyy-MM-dd HH:mm:ss")"
    Body       = [string]$Html
    From       = "windowsupdates@yourcorp.com"                         
    To         = $recepient_email
    SmtpServer = "SMTP.YOURCORP.CORP"
}    

if ( $recepient_email -ne "" ) {
Send-MailMessage @messageParameters -BodyAsHtml
}
