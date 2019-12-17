Param(
    [Parameter(Mandatory = $false, Position = 0)][string] $Computer = '127.0.0.1'
)
$ErrorActionPreference = "SilentlyContinue"

function Test-Port{
    Param(
        [Parameter(Mandatory = $True, Position = 0)][string] $Computer,
        $port = 135,
        $timeout = 500,
        [switch]$showDebug
    )
     
    # Test-Port.ps1
    # Does a TCP connection on specified port (135 by default)
     
    $ErrorActionPreference = "SilentlyContinue"
     
    # Create TCP Client
    $tcpclient = new-Object system.Net.Sockets.TcpClient
     
    # Tell TCP Client to connect to machine on Port
    $iar = $tcpclient.BeginConnect($Computer, $port, $null, $null)
     
    # Set the wait time
    $wait = $iar.AsyncWaitHandle.WaitOne($timeout, $false)
     
    # Check to see if the connection is done
    if (!$wait) {
        # Close the connection and report timeout
        $tcpclient.Close()
        if ($showDebug) {Write-Host "Connection Timeout"}
        Return $false
    }
    else {
        # Close the connection and report the error if there is one
        $error.Clear()
        $tcpclient.EndConnect($iar) | out-Null
        if (!$?) {if ($showDebug) {write-host $error[0]}; $failed = $true}
        $tcpclient.Close()
    }
     
    # Return $true if connection Establish else $False
    
    return (-not $failed)
    
}

if (-not (Test-Port $Computer)) {
    Write-Host -ForegroundColor White -BackgroundColor Red -NoNewline "Erro"
    Write-Host -NoNewline " Cannot Connet to Computer "
    Write-Host -ForegroundColor Yellow "${Computer}"
    Pause
    Exit
}

$PCInfo = New-Object PSObject
$wmiOS = get-wmiobject Win32_OperatingSystem -computername $Computer
$ubr = ([Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine,$Computer)).OpenSubKey('SOFTWARE\Microsoft\Windows NT\CurrentVersion\').GetValue('ubr')
    
$wmiMemArr = Get-WmiObject Win32_PhysicalMemoryArray -computername $Computer
$wmiSYS = get-wmiobject Win32_ComputerSystem -computername $Computer

$PCInfo | Add-Member NoteProperty PCName $wmiOS.CSName    
$PCInfo | Add-Member NoteProperty OS "$($wmiOS.Caption)`t$($wmiOS.OSArchitecture)`t$($wmiOS.Version).${ubr}"
$PCInfo | Add-Member NoteProperty CPU (get-wmiobject Win32_Processor -computername $Computer).name
$PCInfo | Add-Member NoteProperty GPU (get-wmiobject win32_VideoController -computername $Computer).VideoProcessor
$PCInfo | Add-Member NoteProperty GPURAM ([math]::Round((get-wmiobject win32_VideoController -computername $Computer).AdapterRAM / 1MB))
$PCInfo | Add-Member NoteProperty RAM ([math]::Round($wmiSYS.TotalPhysicalMemory / 1MB))
$PCInfo | Add-Member NoteProperty RAMDETAIL (Get-WmiObject Win32_PhysicalMemory -computername $Computer | select-object @{name = "r1"; expression = { $_.PartNumber + " " + $_.Speed + "  " + ([math]::Round($_.Capacity / 1MB)) + "MB" } }).r1
$PCInfo | Add-Member NoteProperty RAMSLOT $wmiMemArr.MemoryDevices
$PCInfo | Add-Member NoteProperty RAMMAX ([math]::Round($wmiMemArr.MaxCapacity / 1KB))
$PCInfo | Add-Member NoteProperty DISK (Get-WmiObject -Class Win32_DiskDrive -ComputerName $Computer | select-object @{name = "r2"; expression = { $_.Caption + " " + $_.SerialNumber } }).r2
$PCInfo | Add-Member NoteProperty DISKMOUNT (Get-WmiObject -Class win32_logicalDisk -ComputerName $Computer | where-object { $_.DriveType -eq 3 } | select-object @{name = "r3"; expression = { $_.DeviceID + "  " + ([math]::Round($_.Size / 1GB)) + "GB" + "  " + (-not $_.QuotasDisabled).toString() } }).r3
$PCInfo | Add-Member NoteProperty MODEL $wmiSYS.model
$PCInfo | Add-Member NoteProperty SN (Get-WmiObject win32_bios -computername $Computer).SerialNumber
$PCInfo | Add-Member NoteProperty MAC (Get-WmiObject -Class "Win32_NetworkAdapter" -computername $Computer | where-object { $_.PhysicalAdapter -eq $true -and $_.MACAddress -ne $null -and $_.Manufacturer -notlike "VM*" } | Select-Object Name, MACAddress)
$PCINfo | Add-Member NoteProperty MONITOR ((((Get-WmiObject WmiMonitorID -Namespace root\wmi -ComputerName $computer | Select-Object @{name = "name"; expression = { $($_.UserFriendlyName -notmatch 0) + "," + " " + $($_.SerialNumberID -notmatch 0) + "`n" } }) | ForEach-Object { $_.name }) | ForEach-Object { [char]$_ }) -join "")

$PCInfo

