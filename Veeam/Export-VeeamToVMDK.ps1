$ErrorActionPreference = 'Continue'

$VBKSourceRoot = 'F:\Backup_Repo'
$TempOutPath = 'E:\XXX-Temp'
$DestPath = "E:\DR-VMDK"
$RestoreVM = 'DC1'

function Export-VeeamToVMDK {
    param (
        [Parameter ( Mandatory = $true )][string] $VMName
    )
    $RestorePoint = Get-VBRRestorePoint -Name ${VMName} | Sort-Object creationtime -Descending | Select-Object -first 1
    if ($RestorePoint.Count -eq 0) {
        # 沒有找到還原點
        "$(Get-date -format s) !!!! [WARN] Cannot find VM ${DestPath}" | Out-File -append "${TempOutPath}\Restore_${VMName}.log"
        return $false
    }

    # 檢查匯出暫存目錄是否存在，若存在先刪除
    if (Test-path "${TempOutPath}\${VMName}") {
        Remove-Item -Force -Recurse "${TempOutPath}\${VMName}" -ErrorAction SilentlyContinue | Out-Null
    }

    # 建立空目錄
    New-Item -ItemType Directory "${TempOutPath}\${VMName}" -ErrorAction SilentlyContinue | Out-Null

    # 寫入 LOG
    $starttime = Get-date
    "$($starttime.ToString('s')) .... [INFO] Start Export 【${VMName}@$($RestorePoint.CreationTime.tostring('s'))】" | Out-File -append "${TempOutPath}\Restore_${VMName}.log"

    # 開始匯出
    $proc = Start-VBRRestoreVMFiles -RestorePoint $RestorePoint -server ${restoreServer} -Path "${TempOutPath}\${VMName}" -Reason "Batch export"
    if ($proc.result -eq 'Success') {
        # 匯出成功
        Write-Host "Restore $($proc.Name) Success"
        $endtime = Get-date
        "$($endtime.ToString('s')) .... [INFO] Export OK【${VMName}@$($RestorePoint.CreationTime.tostring('s'))】" | Out-File -append "${TempOutPath}\Restore_${VMName}.log"
        $timespan = $endtime - $starttime
        "                                Spent $(($timespan.Hours).ToString('00')):$(($timespan.Minutes).ToString('00')):$(($timespan.Seconds).ToString('00'))" | Out-File -append "${TempOutPath}\Restore_${VMName}.log"

        if (Test-Path $DestPath) {
            # 如果正式目錄已存在舊檔，先刪除
            if (Test-Path "${DestPath}\${VMName}") {
                "$(Get-date -format s) .... [INFO] Remove old backup ${DestPath}\${VMName}" | Out-File -append "${TempOutPath}\Restore_${VMName}.log"
                Remove-Item -Recurse -Force "${DestPath}\${VMName}" -ErrorAction SilentlyContinue | Out-Null
            }

            # 移動暫存檔到正式目錄
            Move-Item "${TempOutPath}\${VMName}" "${DestPath}" -Force
            "$(Get-date -format s) .... [INFO] Move to ${DestPath}\${VMName} OK" | Out-File -append "${TempOutPath}\Restore_${VMName}.log"

            # 複製 LOG 到目的地
            Copy-Item "${TempOutPath}\Restore_${RestoreVM}.log" "${DestPath}\${VMName}"
            Rename-Item "${DestPath}\${VMName}\Restore_${RestoreVM}.log" "Restore_${RestoreVM}-$($RestorePoint.CreationTime.ToString('yyyyMMddHHmmss')).log"
            
            Rename-Item "${TempOutPath}\Restore_${RestoreVM}.log" "${TempOutPath}\Restore_${RestoreVM}-$($RestorePoint.CreationTime.ToString('yyyyMMddHHmmss')).log"
            return $true
        }
        else {
            # 正式目錄不存在
            "$(Get-date -format s) !!!! [WARN] Cannot find backup dest path ${DestPath}" | Out-File -append "${TempOutPath}\Restore_${VMName}.log"
            return $true
        }
    }
    else {
        # 匯出失敗
        "$(Get-date -format s) !!!! 【${VMName}】 Export FAIL" | Out-File -append "${TempOutPath}\Restore_${VMName}.log"
        Remove-Item -Recurse -Force "${TempOutPath}\${VMName}" -ErrorAction SilentlyContinue | Out-Null
        return $false
    }
}

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
    "$(Get-date -format s) #### [ERRR] Connection error" | Out-File -append "${TempOutPath}\Restore_${RestoreVM}.log"
}
else {
    "$(Get-Date -Format s) ==== [INFO] Start Export 【${RestoreVM}】" | Out-File -append "${TempOutPath}\Restore_${RestoreVM}.log"

    # 檢查要匯出的主機是否在本機備份清單中
    if ((Get-VBRJobObject -Job $(Get-VBRJob)).Name -contains ${RestoreVM}) {
        # 有就直接匯出
        $result = Export-VeeamToVMDK -VMName ${RestoreVM}
    }
    else {
        # 如果沒有，跑匯入備份檔並匯出的流程

        # 找到 VBK 檔
        $VBKSource = Get-ChildItem -Recurse -File "$VBKSourceRoot" | Where-Object { $_.Extension -eq '.vbk' -and $_.Name -like "$($RestoreVM).*" } | Sort-Object LastWriteTime -Descending | Select-Object Name, Fullname, LastWriteTime -First 1

        # 如果有找到再繼續
        if (-not [string]::IsNullOrEmpty($VBKSource.Name)) {
            # 匯入 VBK
            try {
                "$(Get-date -format s) .... [INFO] import $($VBKSource.Name)" | Out-File -append "${TempOutPath}\Restore_${RestoreVM}.log"
                Import-VBRBackup -Server $RestoreServer -FileName "$($VBKSource.Fullname)"
            }
            catch {
                "$(Get-date -format s) #### [ERRR] import ${RestoreVM} fail" | Out-File -append "${TempOutPath}\Restore_${RestoreVM}.log"
            }

            # 開始匯出
            $result = Export-VeeamToVMDK -VMName ${RestoreVM}
            if ($result) {
                # 成功匯出後移除匯入檔
                #$imported = Get-VBRBackup -Name '*imported'
                #if ($imported.count -gt 0) {
                    #if ((Get-VBRRestorePoint -Backup $imported -Name $RestoreVM).Count -gt 0) {
                        #"$(Get-Date -Format s) !!!! [WRAN] 【${RestoreVM}】 imported" | Out-File -append "${TempOutPath}\Restore_${RestoreVM}.log"
                    #}
                    #Remove-VBRBackup -Confirm
                #}
            }
            else {
                
            }
        }
        else {
            "$(Get-date -format s) #### [ERRR] ${RestoreVM} Not found" | Out-File -append "${TempOutPath}\Restore_${RestoreVM}.log"
        }
    }

    # 關閉連線
    Disconnect-VBRServer
}

if (Test-Path "${TempOutPath}\Restore_${RestoreVM}.log") {
    Rename-Item "${TempOutPath}\Restore_${RestoreVM}.log" "${TempOutPath}\FAIL-Restore_${RestoreVM}-$(Get-Date -Format 'yyyyMMddHHmmss').log"
}