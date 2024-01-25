$ErrorActionPreference = 'Continue'

$VBKSourceRoot = 'F:\VBKSourceRoot'
$DestPath = "E:\VMDK_DestPath"
$VMName = 'VMName'

Import-Module "${PSScriptRoot}\Export-VeeamToVMDK.ps1"

# 開啟連線
$RestoreServer = $null
# 先檢查是否已連線
try {
    $RestoreServer = (Get-VBRServer -ErrorAction SilentlyContinue | Where-Object type -eq 'Local')
}
catch {
    Connect-VBRServer -Server localhost
    $RestoreServer = (Get-VBRServer -ErrorAction SilentlyContinue .\.affinityilentlyContinue | Where-Object type -eq 'Local')
}

# 如果沒有建立連線，回覆錯誤並結束
if ($null -eq $RestoreServer) {
    Write-Host "Connection error"
    Out-LogFile -Level Error -Content "Connection error" -LogFilePath "${env:Temp}\Restore_${VMName}.log"
}
else {
    Out-LogFile -Level Info -Content "Start Export 【${VMName}】" -LogFilePath "${env:Temp}\Restore_${VMName}.log"

    # 檢查要匯出的主機是否在本機備份清單中
    if ((Get-VBRJobObject -Job $(Get-VBRJob)).Name -contains ${VMName}) {
        # 有就直接匯出
        $result = Export-VeeamToVMDK -VMName ${VMName} -OutPath ${DestPath}
    }
    else {
        # 如果沒有，跑匯入備份檔並匯出的流程

        # 找到 VBK 檔
        $VBKSource = Get-ChildItem -Recurse -File "$VBKSourceRoot" | Where-Object { $_.Extension -eq '.vbk' -and $_.Name -like "$($VMName).*" } | Sort-Object LastWriteTime -Descending | Select-Object Name, Fullname, LastWriteTime -First 1

        # 如果有找到再繼續
        if (-not [string]::IsNullOrEmpty($VBKSource.Name)) {
            # 匯入 VBK
            try {
                Out-LogFile -Level Info -Content "import $($VBKSource.Name)" -LogFilePath "${env:Temp}\Restore_${VMName}.log"
                Import-VBRBackup -Server $RestoreServer -FileName "$($VBKSource.Fullname)"
            }
            catch {
                Out-LogFile -Level Error -Content "import ${VMName} fail" -LogFilePath "${env:Temp}\Restore_${VMName}.log"
            }

            # 開始匯出
            $result = Export-VeeamToVMDK -VMName ${VMName} -OutPath ${DestPath}
            if ($result) {
                # 成功匯出後移除匯入檔
                #$imported = Get-VBRBackup -Name '*imported'
                #if ($imported.count -gt 0) {
                    #if ((Get-VBRRestorePoint -Backup $imported -Name $VMName).Count -gt 0) {
                        #"$(Get-Date -Format s) !!!! [WRAN] 【${VMName}】 imported" | Out-File -append "${DestPath}\Restore_${VMName}.log"
                    #}
                    #Remove-VBRBackup -Backup $imported -Confirm
                #}
            }
            else {
                
            }
        }
        else {
            Out-LogFile -Level Error -Content "${VMName} Not found" -LogFilePath "${env:Temp}\Restore_${VMName}.log"
        }
    }

    # 關閉連線
    Disconnect-VBRServer
}

if (Test-Path "${env:Temp}\Restore_${VMName}.log") {
    New-Item -ItemType Directory "${DestPath}\XXX-Log" -Force -ErrorAction SilentlyContinue | Out-Null
    Move-Item "${env:Temp}\Restore_${VMName}.log" "${DestPath}\XXX-Log" -Force
    Rename-Item "${DestPath}\XXX-Log\Restore_${VMName}.log" "Restore_${VMName}-$(Get-Date -Format 'yyyyMMddHHmmss').log" -Force
}
