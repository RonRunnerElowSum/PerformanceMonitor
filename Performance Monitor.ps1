function PunchIt {

    #CPU utilization warning percentage
    $CPUWarningThreshhold = "90"
    #RAM utilization warning threshhold in GB
    $RAMWarningThreshhold = "1"
    #Network interface warning threshhold in Mbps
    $NetInterfaceDownloadWarningThreshhold = "30"
    $NetInterfaceUploadWarningThreshhold = "15"

<#---------------------------------------------------------------------------#>

    if(!(Get-Module -Name "SqlServer")){
        Write-MSPLog -LogSource "MSP Performance Monitor" -LogType "Information" -LogMessage "SqlServer module is not installed...installing now..."
        InstallSqlServerPSModule
    }
    if(!(Test-Path -Path "C:\MSP\secret.txt")){
        Write-MSPLog -LogSource "MSP Performance Monitor" -LogType "Error" -LogMessage "C:\MSP\secret.txt does not exist...exiting..."
        EXIT
    }

    Write-MSPLog -LogSource "MSP Performance Monitor" -LogType "Information" -LogMessage "Starting performance monitor..."

    $CPUPerformanceErrorCounter = @()
    $RAMPerformanceErrorCounter = @()
    $NetworkUploadPerformanceErrorCounter = @()
    $NetworkDownloadPerformanceErrorCounter = @()
    
    do{
    #CPU check
        $CPUUtilization = [math]::Round((Get-Counter -Counter '\processor(_total)\% processor time').CounterSamples.CookedValue,1)
        $AvailableRAMInGB = [math]::Round(((Get-Counter -Counter '\*Memory\Available Bytes').CounterSamples.CookedValue) / 1000000000,1)
        $NetInterfaceUploadUtilizationInMbps = [math]::Round(((Get-Counter -Counter "\Network interface(*)\Bytes sent/sec").CounterSamples.CookedValue | Sort-Object -Descending | Select-Object -First 1) / 125000,1)
        $NetInterfaceDownloadUtilizationInMbps = [math]::Round(((Get-Counter -Counter "\Network interface(*)\Bytes received/sec").CounterSamples.CookedValue | Sort-Object -Descending | Select-Object -First 1) / 125000,1)
        if($CPUUtilization -ge $CPUWarningThreshhold){
            $CPUPerformanceErrorCounter += 1
            if($CPUPerformanceErrorCounter.Count -ge "2"){
                [string]$Top10ProcessesUsingCPU = (Get-Counter -Counter "\Process(*)\% Processor Time").CounterSamples | Select-Object -First 10 | Sort-Object -Property CookedValue -Descending | Format-Table -Property InstanceName, CookedValue -AutoSize | Out-String
                Get-HistoricalCPUPerfData
                Write-MSPLog -LogSource "MSP Performance Monitor" -LogType "Warning" -LogMessage "$Env:ComputerName's CPU utilization has peaked $CPUWarningThreshhold`% for roughly 60 seconds (utilizing $CPUUtilization`%) at $(Get-Date)`r`n`r`nTop 10 processes utilizing CPU:`r`n$Top10ProcessesUsingCPU`r`n`r`nNumber of CPU performance errors today: $NumberOfCPUUtilizationErrorsToday`r`nNumber of CPU performance errors last 30 days: $NumberOfCPUUtilizationErrorsInPast30Days`r`nNumber of CPU performance errors last 60 days: $NumberOfCPUUtilizationErrorsInPast60Days`r`nNumber of CPU performance errors last 90 days: $NumberOfCPUUtilizationErrorsInPast90Days"
                $EndpointCPUStatus = "Warning: High Utilization"
            }
            else{
                $EndpointCPUStatus = "Healthy"
                Get-HistoricalCPUPerfData
            }
        }
        else{
            $CPUPerformanceErrorCounter = @()
            $EndpointCPUStatus = "Healthy"
            Get-HistoricalCPUPerfData
        }
    #RAM check
        if($AvailableRAMInGB -le $RAMWarningThreshhold){
            $RAMPerformanceErrorCounter += 1
            if($RAMPerformanceErrorCounter.Count -ge "3"){
                $OSArch = (Get-WmiObject Win32_OperatingSystem).OSArchitecture
                if($OSArch -eq "32-bit"){
                    [string]$Top20ProcessesUsingRAM = Get-Process | Sort-Object -Descending WorkingSet | Select-Object Name,@{Name='RAM Used (MB)';Expression={($_.WorkingSet/1MB)}} | Select-Object -First 20 | Out-String
                }
                elseif($OSArch -eq "64-bit"){
                    [string]$Top20ProcessesUsingRAM = Get-Process | Sort-Object -Descending WorkingSet64 | Select-Object Name,@{Name='RAM Used (MB)';Expression={($_.WorkingSet64/1MB)}} | Select-Object -First 20 | Out-String
                }
                Get-HistoricalRAMPerfData
                Write-MSPLog -LogSource "MSP Performance Monitor" -LogType "Warning" -LogMessage "$Env:ComputerName's available RAM has gone below $RAMWarningThreshhold`GB and has remained below $RAMWarningThreshhold`GB for roughly 120 seconds (available RAM: $AvailableRAMInGB`GB) at $(Get-Date)`r`n`r`nTop 10 processes utilizing RAM:`r`n$Top20ProcessesUsingRAM`r`n`r`nNumber of RAM performance errors today: $NumberOfRAMUtilizationErrorsToday`r`nNumber of RAM performance errors last 30 days: $NumberOfRAMUtilizationErrorsInPast30Days`r`nNumber of RAM performance errors last 60 days: $NumberOfRAMUtilizationErrorsInPast60Days`r`nNumber of RAM performance errors last 90 days: $NumberOfRAMUtilizationErrorsInPast90Days"
                $EndpointRAMStatus = "Warning: High Utilization"
            }
            else{
                $EndpointRAMStatus = "Healthy"
                Get-HistoricalRAMPerfData
            }
        }
        else{
            $RAMPerformanceErrorCounter = @()
            $EndpointRAMStatus = "Healthy"
            Get-HistoricalRAMPerfData
        }
    #Network interface upload check
        if($NetInterfaceUploadUtilizationInMbps -ge $NetInterfaceUploadWarningThreshhold){
            $NetworkUploadPerformanceErrorCounter += 1
            if($NetworkUploadPerformanceErrorCounter.Count -ge "3"){
                Get-HistoricalNetworkUploadPerfData
                Write-MSPLog -LogSource "MSP Performance Monitor" -LogType "Warning" -LogMessage "$Env:ComputerName's network interface upload utilization has peaked $NetInterfaceUploadWarningThreshhold`Mbps for roughly 120 seconds (utilizing $NetInterfaceUploadUtilizationInMbps`Mbps) at $(Get-Date)`r`n`r`nNumber of network upload performance errors today: $NumberOfNetworkUtilizationUploadErrorsToday`r`nNumber of network upload performance errors last 30 days: $NumberOfNetworkUtilizationUploadErrorsInPast30Days`r`nNumber of network upload performance errors last 60 days: $NumberOfNetworkUtilizationUploadErrorsInPast60Days`r`nNumber of network upload performance errors last 90 days: $NumberOfNetworkUtilizationUploadErrorsInPast90Days"
                $EndpointNetIntUploadStatus = "Warning: High Upload Utilization"
            }
            else{
                $EndpointNetIntUploadStatus = "Healthy"
                Get-HistoricalNetworkUploadPerfData
            }
        }
        else{
            $NetworkUploadPerformanceErrorCounter = @()
            $EndpointNetIntUploadStatus = "Healthy"
            Get-HistoricalNetworkUploadPerfData
        }
    #Network interface download check
        if($NetInterfaceDownloadUtilizationInMbps -ge $NetInterfaceDownloadWarningThreshhold){
            $NetworkDownloadPerformanceErrorCounter += 1
            if($NetworkDownloadPerformanceErrorCounter.Count -ge "3"){
                Get-HistoricalNetworkDownloadPerfData
                Write-MSPLog -LogSource "MSP Performance Monitor" -LogType "Warning" -LogMessage "$Env:ComputerName's network interface download utilization has peaked $NetInterfaceDownloadWarningThreshhold`Mbps for roughly 120 seconds (utilizing $NetInterfaceDownloadUtilizationInMbps`Mbps) at $(Get-Date)`r`n`r`nNumber of network download performance errors today: $NumberOfNetworkUtilizationDownloadErrorsToday`r`nNumber of network download performance errors last 30 days: $NumberOfNetworkUtilizationDownloadErrorsInPast30Days`r`nNumber of network download performance errors last 60 days: $NumberOfNetworkUtilizationDownloadErrorsInPast60Days`r`nNumber of network download performance errors last 90 days: $NumberOfNetworkUtilizationDownloadErrorsInPast90Days"
                $EndpointNetIntDownloadStatus = "Warning: High Download Utilization"
            }
            else{
                $EndpointNetIntDownloadStatus = "Healthy"
                Get-HistoricalNetworkDownloadPerfData
            }
        }
        else{
            $NetworkDownloadPerformanceErrorCounter = @()
            $EndpointNetIntDownloadStatus = "Healthy"
            Get-HistoricalNetworkDownloadPerfData
        }

        if(($EndpointCPUStatus -eq "Healthy") -and ($EndpointRAMStatus -eq "Healthy") -and ($EndpointNetIntUploadStatus -eq "Healthy") -and ($EndpointNetIntDownloadStatus -eq "Healthy")){
            Script:$EndpointHasPerformanceIssues = "False"
        }
        else{
            Script:$EndpointHasPerformanceIssues = "True"
        }
    
        PostPerformanceData
        Start-Sleep -Seconds 45
    }
    while(
        $True
    )
}

function Get-HistoricalCPUPerfData {
    Script:$NumberOfCPUUtilizationErrorsToday = (Get-EventLog -LogName MSP-IT -EntryType Warning -Source "MSP Performance Monitor" | Where-Object {$_.TimeWritten | Select-String "$(Get-Date -Format "MM/dd/yyyy")"} | Where-Object {$_.Message -Like "*CPU utilization*"}).Count
    Script:$NumberOfCPUUtilizationErrorsInPast30Days = (Get-EventLog -LogName MSP-IT -EntryType Warning -Source "MSP Performance Monitor" | Where-Object {[datetime]$_.TimeWritten -ge [datetime]$(Get-Date).AddDays(-30)} | Where-Object {$_.Message -Like "*CPU utilization*"}).Count
    Script:$NumberOfCPUUtilizationErrorsInPast60Days = (Get-EventLog -LogName MSP-IT -EntryType Warning -Source "MSP Performance Monitor" | Where-Object {[datetime]$_.TimeWritten -ge [datetime]$(Get-Date).AddDays(-60)} | Where-Object {$_.Message -Like "*CPU utilization*"}).Count
    Script:$NumberOfCPUUtilizationErrorsInPast90Days = (Get-EventLog -LogName MSP-IT -EntryType Warning -Source "MSP Performance Monitor" | Where-Object {[datetime]$_.TimeWritten -ge [datetime]$(Get-Date).AddDays(-90)} | Where-Object {$_.Message -Like "*CPU utilization*"}).Count
}

function Get-HistoricalRAMPerfData {
    Script:$NumberOfRAMUtilizationErrorsToday = (Get-EventLog -LogName MSP-IT -EntryType Warning -Source "MSP Performance Monitor" | Where-Object {$_.TimeWritten | Select-String "$(Get-Date -Format "MM/dd/yyyy")"} | Where-Object {$_.Message -Like "*available RAM*"}).Count
    Script:$NumberOfRAMUtilizationErrorsInPast30Days = (Get-EventLog -LogName MSP-IT -EntryType Warning -Source "MSP Performance Monitor" | Where-Object {[datetime]$_.TimeWritten -ge [datetime]$(Get-Date).AddDays(-30)} | Where-Object {$_.Message -Like "*available RAM*"}).Count
    Script:$NumberOfRAMUtilizationErrorsInPast60Days = (Get-EventLog -LogName MSP-IT -EntryType Warning -Source "MSP Performance Monitor" | Where-Object {[datetime]$_.TimeWritten -ge [datetime]$(Get-Date).AddDays(-60)} | Where-Object {$_.Message -Like "*available RAM*"}).Count
    Script:$NumberOfRAMUtilizationErrorsInPast90Days = (Get-EventLog -LogName MSP-IT -EntryType Warning -Source "MSP Performance Monitor" | Where-Object {[datetime]$_.TimeWritten -ge [datetime]$(Get-Date).AddDays(-90)} | Where-Object {$_.Message -Like "*available RAM*"}).Count
}

function Get-HistoricalNetworkUploadPerfData {
    Script:$NumberOfNetworkUtilizationUploadErrorsToday = (Get-EventLog -LogName MSP-IT -EntryType Warning -Source "MSP Performance Monitor" | Where-Object {$_.TimeWritten | Select-String "$(Get-Date -Format "MM/dd/yyyy")"} | Where-Object {$_.Message -Like "*network interface upload*"}).Count
    Script:$NumberOfNetworkUtilizationUploadErrorsInPast30Days = (Get-EventLog -LogName MSP-IT -EntryType Warning -Source "MSP Performance Monitor" | Where-Object {[datetime]$_.TimeWritten -ge [datetime]$(Get-Date).AddDays(-30)} | Where-Object {$_.Message -Like "*network interface upload*"}).Count
    Script:$NumberOfNetworkUtilizationUploadErrorsInPast60Days = (Get-EventLog -LogName MSP-IT -EntryType Warning -Source "MSP Performance Monitor" | Where-Object {[datetime]$_.TimeWritten -ge [datetime]$(Get-Date).AddDays(-60)} | Where-Object {$_.Message -Like "*network interface upload*"}).Count
    Script:$NumberOfNetworkUtilizationUploadErrorsInPast90Days = (Get-EventLog -LogName MSP-IT -EntryType Warning -Source "MSP Performance Monitor" | Where-Object {[datetime]$_.TimeWritten -ge [datetime]$(Get-Date).AddDays(-90)} | Where-Object {$_.Message -Like "*network interface upload*"}).Count
}

function Get-HistoricalNetworkDownloadPerfData {
    Script:$NumberOfNetworkUtilizationDownloadErrorsToday = (Get-EventLog -LogName MSP-IT -EntryType Warning -Source "MSP Performance Monitor" | Where-Object {$_.TimeWritten | Select-String "$(Get-Date -Format "MM/dd/yyyy")"} | Where-Object {$_.Message -Like "*network interface download*"}).Count
    Script:$NumberOfNetworkUtilizationDownloadErrorsInPast30Days = (Get-EventLog -LogName MSP-IT -EntryType Warning -Source "MSP Performance Monitor" | Where-Object {[datetime]$_.TimeWritten -ge [datetime]$(Get-Date).AddDays(-30)} | Where-Object {$_.Message -Like "*network interface download*"}).Count
    Script:$NumberOfNetworkUtilizationDownloadErrorsInPast60Days = (Get-EventLog -LogName MSP-IT -EntryType Warning -Source "MSP Performance Monitor" | Where-Object {[datetime]$_.TimeWritten -ge [datetime]$(Get-Date).AddDays(-60)} | Where-Object {$_.Message -Like "*network interface download*"}).Count
    Script:$NumberOfNetworkUtilizationDownloadErrorsInPast90Days = (Get-EventLog -LogName MSP-IT -EntryType Warning -Source "MSP Performance Monitor" | Where-Object {[datetime]$_.TimeWritten -ge [datetime]$(Get-Date).AddDays(-90)} | Where-Object {$_.Message -Like "*network interface download*"}).Count
}

function Write-MSPLog {
    Param
    (
         [Parameter(Mandatory=$true, Position=0)]
         [ValidateSet('MSP Performance Monitor')]
         [string] $LogSource,
         [Parameter(Mandatory=$true, Position=1)]
         [ValidateSet('Information','Warning','Error')]
         [string] $LogType,
         [Parameter(Mandatory=$true, Position=2)]
         [string] $LogMessage
    )

    New-EventLog -LogName MSP-IT -Source 'MSP' -ErrorAction SilentlyContinue
    if(!(Get-EventLog -LogName MSP-IT -Source 'MSP Performance Monitor' -ErrorAction SilentlyContinue)){
        New-EventLog -LogName MSP-IT -Source 'MSP Performance Monitor' -ErrorAction SilentlyContinue
    }
    Write-EventLog -Log MSP-IT -Source "MSP Performance Monitor" -EventID 0 -EntryType $LogType -Message "$LogMessage"
}

function InstallSqlServerPSModule () {
    Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Scope CurrentUser -Force -ErrorAction SilentlyContinue
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Install-Module SqlServer -Force -AllowClobber | Out-Null
    Import-Module SqlServer -Force | Out-Null
    if(!(Get-Module -Name "SqlServer")){
        Write-MSPLog -LogSource "MSP Performance Monitor" -LogType "Warning" -LogMessage "The module SqlServer failed to install..."
        EXIT
    }
}

function PostPerformanceData {

    $EndpointSerial = (Get-CimInstance win32_bios).SerialNumber
    $EndpointComputerName = [System.Net.Dns]::GetHostByName($Env:ComputerName).HostName
    $EndpointOS = (Get-WmiObject -class Win32_OperatingSystem).Caption
    if($EndpointOS | Select-String "Server"){
        $EndpointType = "Server"
    }
    else{
        $EndpointType = "Workstation"
    }
    if(($Null -eq $EndpointSerial) -or ($EndpointSerial -eq "To be filled by O.E.M.")){
        $EndpointSerial = $EndpointComputerName
    }
    if(Test-Path "C:\ComSys\EndpointSiteName.txt" -ErrorAction SilentlyContinue){
        $EndpointSiteName = Get-Content -Path "C:\ComSys\EndpointSiteName.txt"
    }
    elseif(Test-Path "C:\Program Files\SAAZOD\ApplicationLog\zSCCLog\zDCMGetSitecode.log" -ErrorAction SilentlyContinue){
        $EndpointSiteName = (Get-Content -Path "C:\Program Files\SAAZOD\ApplicationLog\zSCCLog\zDCMGetSitecode.log" | Select-String "sitename=" | Select-Object -Last 1) -Replace "^.*?="
    }
    elseif(Test-Path "C:\Program Files (x86)\SAAZOD\ApplicationLog\zSCCLog\zDCMGetSitecode.log" -ErrorAction SilentlyContinue){
        $EndpointSiteName = (Get-Content -Path "C:\Program Files (x86)\SAAZOD\ApplicationLog\zSCCLog\zDCMGetSitecode.log" | Select-String "sitename=" | Select-Object -Last 1) -Replace "^.*?="
    }
    if($Null -eq $EndpointSiteName){
        if(Test-Path "C:\Program Files\SAAZOD\ApplicationLog\WebPost\WebPostDLL\WebPostComp-tfr_wpdcmgetsitecode.Log" -ErrorAction SilentlyContinue){
            $EndpointSiteName = (((((Get-Content -Path "C:\Program Files\SAAZOD\ApplicationLog\WebPost\WebPostDLL\WebPostComp-tfr_wpdcmgetsitecode.Log" | Select-String "<sitename>" | Select-Object -Last 1) -Replace "^.*?!") -Replace "^.*?!") -Replace "^.*?!") -Replace '\[CDATA\[','') -Split '\]\]'[-1] | Select-Object -First 1
        }
        if(Test-Path "C:\Program Files (x86)\SAAZOD\ApplicationLog\WebPost\WebPostDLL\WebPostComp-tfr_wpdcmgetsitecode.Log" -ErrorAction SilentlyContinue){
            $EndpointSiteName = (((((Get-Content -Path "C:\Program Files (x86)\SAAZOD\ApplicationLog\WebPost\WebPostDLL\WebPostComp-tfr_wpdcmgetsitecode.Log" | Select-String "<sitename>" | Select-Object -Last 1) -Replace "^.*?!") -Replace "^.*?!") -Replace "^.*?!") -Replace '\[CDATA\[','') -Split '\]\]'[-1] | Select-Object -First 1
        }
    }

    try{
        if(!(Test-Path -Path "C:\MSP\secret.txt")){
            Write-MSPLog -LogSource "MSP Performance Monitor" -LogType "Error" -LogMessage "C:\MSP\secret.txt does not exist...exiting..."
            EXIT
        }
        $SensitiveString = Get-Content -Path "C:\MSP\secret.txt" | ConvertTo-SecureString
        $Marshal = [System.Runtime.InteropServices.Marshal]
        $Bstr = $Marshal::SecureStringToBSTR($SensitiveString)
        $DecryptedString = $Marshal::PtrToStringAuto($Bstr)
        $Marshal::ZeroFreeBSTR($Bstr)
        $SQLServer = $DecryptedString -split ";" | Select-Object -Index 0
        $SQLDatabase = $DecryptedString -split ";" | Select-Object -Index 1
        $SQLUsername = $DecryptedString -split ";" | Select-Object -Index 2
        $SQLPassword = $DecryptedString -split ";" | Select-Object -Index 3
    }
    catch{
        Write-MSPLog -LogSource "MSP Performance Monitor" -LogType "Error" -LogMessage "Failed to decrypt SQL connection info...exiting..."
    }

    $DateTime = Get-Date -Format "MM/dd/yyyy HH:mm"
    #Write-MSPLog -LogSource "MSP Performance Monitor" -LogType "Information" -LogMessage "Posting the following information:`r`n`r`nSerial: $EndpointSerial`r`nComputer Name: $EndpointComputerName`r`nOS: $EndpointOS`r`nType: $EndpointType`r`nSite: $EndpointSiteName`r`nCPU Status: $EndpointCPUStatus`r`nCPU Utilization: $CPUUtilization`%`r`nRAM Status: $EndpointRAMStatus`r`nAvailable RAM (GB): $AvailableRAMInGB`GB`r`nUpload status: $EndpointNetIntUploadStatus`r`nUpload Utilization: $NetInterfaceUploadUtilizationInMbps`Mbps`r`nDownload status: $EndpointNetIntDownloadStatus`r`nDownload utilization: $NetInterfaceDownloadUtilizationInMbps`Mbps`r`nDate/Time: $DateTime'"

$SQLCommand = @"
if exists(SELECT * from Table_CustomerPerformanceData where EndpointSerial='$EndpointSerial')
BEGIN            
UPDATE Table_CustomerPerformanceData SET EndpointComputerName='$EndpointComputerName',EndpointOS='$EndpointOS',EndpointType='$EndpointType',EndpointSiteName='$EndpointSiteName',EndpointCPUStatus='$EndpointCPUStatus',EndpointCPUUtilization='$CPUUtilization',EndpointCPUPerfErrorsCountToday='$NumberOfCPUUtilizationErrorsToday',EndpointCPUPerfErrorsCountLast30Days='$NumberOfCPUUtilizationErrorsInPast30Days',EndpointCPUPerfErrorsCountLast60Days='$NumberOfCPUUtilizationErrorsInPast60Days',EndpointCPUPerfErrorsCountLast90Days='$NumberOfCPUUtilizationErrorsInPast90Days',EndpointRAMStatus='$EndpointRAMStatus',EndpointAvailableRAMInGB='$AvailableRAMInGB',EndpointRAMPerfErrorsCountToday='$NumberOfRAMUtilizationErrorsToday',EndpointRAMPerfErrorsCountLast30Days='$NumberOfRAMUtilizationErrorsInPast30Days',EndpointRAMPerfErrorsCountLast60Days='$NumberOfRAMUtilizationErrorsInPast60Days',EndpointRAMPerfErrorsCountLast90Days='$NumberOfRAMUtilizationErrorsInPast90Days',EndpointNetInterfaceUploadStatus='$EndpointNetIntUploadStatus',EndpointNetInterfaceUploadUtilizationInMbps='$NetInterfaceUploadUtilizationInMBps',EndpointNetUploadPerfErrorsCountToday='$NumberOfNetworkUtilizationUploadErrorsToday',EndpointNetUploadPerfErrorsCountLast30Days='$NumberOfNetworkUtilizationUploadErrorsInPast30Days',EndpointNetUploadPerfErrorsCountLast60Days='$NumberOfNetworkUtilizationUploadErrorsInPast60Days',EndpointNetUploadPerfErrorsCountLast90Days='$NumberOfNetworkUtilizationUploadErrorsInPast90Days',EndpointNetInterfaceDownloadStatus='$EndpointNetIntDownloadStatus',EndpointNetInterfaceDownloadUtilizationInMbps='$NetInterfaceDownloadUtilizationInMBps',EndpointNetDownloadPerfErrorsCountToday='$NumberOfNetworkUtilizationDownloadErrorsToday',EndpointNetDownloadPerfErrorsCountLast30Days='$NumberOfNetworkUtilizationDownloadErrorsPast30Days',EndpointNetDownloadPerfErrorsCountLast60Days='$NumberOfNetworkUtilizationDownloadErrorsPast60Days',EndpointNetDownloadPerfErrorsCountLast90Days='$NumberOfNetworkUtilizationDownloadErrorsPast90Days',EndpointHasPerformanceIssues='$EndpointHasPerformanceIssues',LastPostDate='$DateTime' WHERE (EndpointSerial = '$EndpointSerial')
END
else            
BEGIN
INSERT INTO [$SQLDatabase].[dbo].[Table_CustomerPerformanceData](EndpointSerial,EndpointComputerName,EndpointOS,EndpointType,EndpointSiteName,EndpointCPUStatus,EndpointCPUUtilization,EndpointCPUPerfErrorsCountToday,EndpointCPUPerfErrorsCountLast30Days,EndpointCPUPerfErrorsCountLast60Days,EndpointCPUPerfErrorsCountLast90Days,EndpointRAMStatus,EndpointAvailableRAMInGB,EndpointRAMPerfErrorsCountToday,EndpointRAMPerfErrorsCountLast30Days,EndpointRAMPerfErrorsCountLast60Days,EndpointRAMPerfErrorsCountLast90Days,EndpointNetInterfaceUploadStatus,EndpointNetInterfaceUploadUtilizationInMbps,EndpointNetUploadPerfErrorsCountToday,EndpointNetUploadPerfErrorsCountLast30Days,EndpointNetUploadPerfErrorsCountLast60Days,EndpointNetUploadPerfErrorsCountLast90Days,EndpointNetInterfaceDownloadStatus,EndpointNetInterfaceDownloadUtilizationInMbps,EndpointNetDownloadPerfErrorsCountToday,EndpointNetDownloadPerfErrorsCountLast30Days,EndpointNetDownloadPerfErrorsCountLast60Days,EndpointNetDownloadPerfErrorsCountLast90Days,EndpointHasPerformanceIssues,LastPostDate)
VALUES ('$EndpointSerial','$EndpointComputerName','$EndpointOS','$EndpointType','$EndpointSiteName','$EndpointCPUStatus','$CPUUtilization','$NumberOfCPUUtilizationErrorsToday','$NumberOfCPUUtilizationErrorsInPast30Days','$NumberOfCPUUtilizationErrorsInPast60Days','$NumberOfCPUUtilizationErrorsInPast90Days','$EndpointRAMStatus','$AvailableRAMInGB','$NumberOfRAMUtilizationErrorsToday','$NumberOfRAMUtilizationErrorsInPast30Days','$NumberOfRAMUtilizationErrorsInPast60Days','$NumberOfRAMUtilizationErrorsInPast90Days','$EndpointNetIntUploadStatus','$NetInterfaceUploadUtilizationInMBps','$NumberOfNetworkUtilizationUploadErrorsToday','$NumberOfNetworkUtilizationUploadErrorsInPast30Days','$NumberOfNetworkUtilizationUploadErrorsInPast60Days','$NumberOfNetworkUtilizationUploadErrorsInPast90Days','$EndpointNetIntDownloadStatus','$NetInterfaceDownloadUtilizationInMBps','$NumberOfNetworkUtilizationDownloadErrorsToday','$NumberOfNetworkUtilizationDownloadErrorsInPast30Days','$NumberOfNetworkUtilizationDownloadErrorsInPast60Days','$NumberOfNetworkUtilizationDownloadErrorsInPast90Days','$EndpointHasPerformanceIssues','$DateTime')
END
"@      

    $Params = @{
        'ServerInstance'=$SQLServer;
        'Database'=$SQLDatabase;
        'Username'=$SQLUsername;
        'Password'=$SQLPassword
    }
    
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Invoke-SqlCmd @Params -Query $SQLCommand -EncryptConnection
}
