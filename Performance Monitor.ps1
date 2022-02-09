#Warning  utilization percentage is greater than or equal to this amount
$CPUWarningThreshhold = "90"
#Warning when this amount or less of RAM is remaining
$RAMWarningThreshhold = "1"
#Network interface warning threshhold in Mbps
$NetInterfaceUploadWarningThreshhold = "15"
$NetInterfaceDownloadWarningThreshhold = "30"

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
    Write-MSPLog -LogSource "MSP Performance Monitor" -LogType "Information" -LogMessage "Posting the following information:`r`n`r`nSerial: $EndpointSerial`r`nComputer Name: $EndpointComputerName`r`nOS: $EndpointOS`r`nType: $EndpointType`r`nSite: $EndpointSiteName`r`nCPU Status: $EndpointCPUStatus`r`nCPU Utilization: $CPUUtilization`%`r`nRAM Status: $EndpointRAMStatus`r`nAvailable RAM (GB): $AvailableRAMInGB`GB`r`nUpload status: $EndpointNetIntUploadStatus`r`nUpload Utilization: $NetInterfaceUploadUtilizationInMBps`r`nDownload status: $EndpointNetIntDownloadStatus`r`nDownload utilization: $NetInterfaceDownloadUtilizationInMBps`r`nDate/Time: $DateTime'"

$SQLCommand = @"
if exists(SELECT * from Table_CustomerPerformanceData where EndpointSerial='$EndpointSerial')
BEGIN            
UPDATE Table_CustomerPerformanceData SET EndpointComputerName='$EndpointComputerName',EndpointOS='$EndpointOS',EndpointType='$EndpointType',EndpointSiteName='$EndpointSiteName',EndpointCPUStatus='$EndpointCPUStatus',EndpointCPUUtilization='$CPUUtilization',EndpointRAMStatus='$EndpointRAMStatus',EndpointAvailableRAMInGB='$AvailableRAMInGB',EndpointNetInterfaceUploadStatus='$EndpointNetIntUploadStatus',EndpointNetInterfaceUploadUtilizationInMBps='$NetInterfaceUploadUtilizationInMBps',EndpointNetInterfaceDownloadStatus='$EndpointNetIntDownloadStatus',EndpointNetInterfaceDownloadUtilizationInMBps='$NetInterfaceDownloadUtilizationInMBps',EndpointHasPerformanceIssues='$EndpointHasPerformanceIssues',LastPostDate='$DateTime' WHERE (EndpointSerial = '$EndpointSerial')
END                  
else            
BEGIN
INSERT INTO [$SQLDatabase].[dbo].[Table_CustomerPerformanceData](EndpointSerial,EndpointComputerName,EndpointOS,EndpointType,EndpointSiteName,EndpointCPUStatus,EndpointCPUUtilization,EndpointRAMStatus,EndpointAvailableRAMInGB,EndpointNetInterfaceUploadStatus,EndpointNetInterfaceUploadUtilizationInMBps,EndpointNetInterfaceDownloadStatus,EndpointNetInterfaceDownloadUtilizationInMBps,EndpointHasPerformanceIssues,LastPostDate)
VALUES ('$EndpointSerial','$EndpointComputerName','$EndpointOS','$EndpointType','$EndpointSiteName','$EndpointCPUStatus','$CPUUtilization','$EndpointRAMStatus','$AvailableRAMInGB','$EndpointNetIntUploadStatus','$NetInterfaceUploadUtilizationInMBps','$EndpointNetIntDownloadStatus','$NetInterfaceDownloadUtilizationInMBps','$EndpointHasPerformanceIssues','$DateTime')
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

if(!(Test-Path -Path "C:\MSP\secret.txt")){
    Write-MSPLog -LogSource "MSP Performance Monitor" -LogType "Error" -LogMessage "C:\MSP\secret.txt does not exist...exiting..."
    EXIT
}

do{
    $CPUUtilization = [math]::Round((Get-Counter -Counter '\processor(_total)\% processor time').CounterSamples.CookedValue,1)
    $AvailableRAMInGB = [math]::Round(((Get-Counter -Counter '\*Memory\Available Bytes').CounterSamples.CookedValue) / 1000000000,1)
    $NetInterfaceUploadUtilizationInMbps = [math]::Round(((Get-Counter -Counter "\Network interface(*)\Bytes sent/sec").CounterSamples.CookedValue | Sort-Object -Descending | Select-Object -First 1) / 125000,1)
    $NetInterfaceDownloadUtilizationInMbps = [math]::Round(((Get-Counter -Counter "\Network interface(*)\Bytes received/sec").CounterSamples.CookedValue | Sort-Object -Descending | Select-Object -First 1) / 125000,1)
    if($CPUUtilization -ge $CPUWarningThreshhold){
        [string]$Top10ProcessesUsingCPU = (Get-Counter -Counter "\Process(*)\% Processor Time").CounterSamples | Select-Object -First 10 | Sort-Object -Property CookedValue -Descending | Format-Table -Property InstanceName, CookedValue -AutoSize | Out-String
        Write-MSPLog -LogSource "MSP Performance Monitor" -LogType "Warning" -LogMessage "$Env:ComputerName's CPU utilization has peaked $CPUWarningThreshhold`% (utilizing $CPUUtilization`%) at $(Get-Date)`r`n`r`nTop 10 processes utilizing CPU:`r`n$Top10ProcessesUsingCPU"
        $EndpointCPUStatus = "Warning: High Utilization"
    }
    else{
        $EndpointCPUStatus = "Healthy"
    }
    if($AvailableRAMInGB -le $RAMWarningThreshhold){
        $OSArch = (Get-WmiObject Win32_OperatingSystem).OSArchitecture
        if($OSArch -eq "32-bit"){
            [string]$Top20ProcessesUsingRAM = Get-Process | Sort-Object -Descending WorkingSet | Select-Object Name,@{Name='RAM Used (MB)';Expression={($_.WorkingSet/1MB)}} | Select-Object -First 20 | Out-String
        }
        elseif($OSArch -eq "64-bit"){
            [string]$Top20ProcessesUsingRAM = Get-Process | Sort-Object -Descending WorkingSet64 | Select-Object Name,@{Name='RAM Used (MB)';Expression={($_.WorkingSet64/1MB)}} | Select-Object -First 20 | Out-String
        }
        Write-MSPLog -LogSource "MSP Performance Monitor" -LogType "Warning" -LogMessage "$Env:ComputerName's available RAM has gone below $RAMWarningThreshhold`GB (available RAM: $AvailableRAMInGB`GB) at $(Get-Date)`r`n`r`nTop 10 processes utilizing RAM:`r`n$Top20ProcessesUsingRAM"
        $EndpointRAMStatus = "Warning: High Utilization"
    }
    else{
        $EndpointRAMStatus = "Healthy"
    }
    if($NetInterfaceUploadUtilizationInMbps -ge $NetInterfaceUploadWarningThreshhold){
        Write-MSPLog -LogSource "MSP Performance Monitor" -LogType "Warning" -LogMessage "$Env:ComputerName's network interface upload utilization has peaked $NetInterfaceUploadWarningThreshhold`Mbps (utilizing $NetInterfaceUploadUtilizationInMbps`Mbps) at $(Get-Date)"
        $EndpointNetIntUploadStatus = "Warning: High Upload Utilization"
    }
    else{
        $EndpointNetIntUploadStatus = "Healthy"
    }
    if($NetInterfaceDownloadUtilizationInMbps -ge $NetInterfaceDownloadWarningThreshhold){
        Write-MSPLog -LogSource "MSP Performance Monitor" -LogType "Warning" -LogMessage "$Env:ComputerName's network interface download utilization has peaked $NetInterfaceDownloadWarningThreshhold`Mbps (utilizing $NetInterfaceDownloadUtilizationInMbps`Mbps) at $(Get-Date)"
        $EndpointNetIntDownloadStatus = "Warning: High Download Utilization"
    }
    else{
        $EndpointNetIntDownloadStatus = "Healthy"
    }
    if(($EndpointCPUStatus = "Healthy") -and ($EndpointRAMStatus = "Healthy") -and ($EndpointNetIntUploadStatus = "Healthy") -and ($EndpointNetIntDownloadStatus = "Healthy")){
        $EndpointHasPerformanceIssues = "False"
    }
    else{
        $EndpointHasPerformanceIssues = "True"
    }

    PostPerformanceData
    Start-Sleep -Seconds 30
}
while(
    $True
)
