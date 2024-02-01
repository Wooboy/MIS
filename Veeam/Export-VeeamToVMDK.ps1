$ErrorActionPreference = 'Continue'

# ======== SCRIPT SETTINGS     ========
$vms = @('VM1', 'VM2')
$VBKSourceRoot = 'F:\Backup_Repo'
$TargetDisk = 'D:'
$DestPath = "${TargetDisk}\02-DR_VMDK"
$LogFolder = "${TargetDisk}\99-Log"
$LogPath = "${LogFolder}\Backup_2nd-$(Get-Date -Format 'yyyyMMdd').log"
# ======== SCRIPT SETTINGS END ========

# ======== CREATE FOLDER     ========
New-Item -ItemType Directory "${DestPath}" -Force -ErrorAction SilentlyContinue | Out-Null
New-Item -ItemType Directory "${LogFolder}" -Force -ErrorAction SilentlyContinue | Out-Null
# ======== CREATE FOLDER END ========

Import-Module "${PSScriptRoot}\Out-LogFile.ps1"

function Export-VeeamToVMDK {
    param (
        [Parameter(Mandatory = $true)][string] $VMName,
        [Parameter(Mandatory = $true)][string] $OutPath
    )
    <#
    .DESCRIPTION
    Only tested on Veeam 12

    .OUTPUTS
    boolean 
    #>

    $RestorePoint = Get-VBRRestorePoint -Name ${VMName} | Sort-Object creationtime -Descending | Select-Object -first 1
    if ($RestorePoint.Count -eq 0) {
        # 沒有找到還原點
        Out-LogFile -Level Error -Content "Cannot find VM 【${VMName}】 restore point" -LogFilePath "${env:Temp}\Restore_${VMName}.log"
        return $false
    }

    # 若輸出目錄不存在，強制建立
    if (-not (Test-Path ${OutPath})) {
        New-Item -ItemType Directory "${OutPath}" -Force -ErrorAction SilentlyContinue | Out-Null
    }

    # 檢查匯出暫存目錄是否存在，若存在先刪除
    if (Test-path "${OutPath}\${VMName}-Temp") {
        Remove-Item -Force -Recurse "${OutPath}\${VMName}-Temp" -ErrorAction SilentlyContinue | Out-Null
    }

    # 建立空目錄
    New-Item -ItemType Directory "${OutPath}\${VMName}-Temp" -ErrorAction SilentlyContinue | Out-Null

    # 寫入 LOG
    $starttime = Get-date
    Out-LogFile -Level Info -Content "Start Export 【${VMName}@$($RestorePoint.CreationTime.tostring('s'))】" -LogFilePath "${env:Temp}\Restore_${VMName}.log"

    # 開始匯出
    $proc = Start-VBRRestoreVMFiles  -RestorePoint $RestorePoint -server ${VeeamBackupServer} -Path "${OutPath}\${VMName}-Temp" -Reason "Batch export"
    if ($proc.result -eq 'Success') {
        # 匯出成功
        $endtime = Get-date
        $timespan = $endtime - $starttime
        $logmsg = "Export OK【${VMName}@$($RestorePoint.CreationTime.tostring('s'))】"
        $logmsg += "`n"
        $logmsg += "                              Spent $(($timespan.Hours).ToString('00')):$(($timespan.Minutes).ToString('00')):$(($timespan.Seconds).ToString('00'))"
        Out-LogFile -Level Info -Content $logmsg -LogFilePath "${env:Temp}\Restore_${VMName}.log"

        # 如果正式目錄已存在舊檔，先刪除
        if (Test-Path "${OutPath}\${VMName}") {
            Out-LogFile -Level Info -Content "Remove old backup ${OutPath}\${VMName}" -LogFilePath "${env:Temp}\Restore_${VMName}.log"
            Remove-Item -Recurse -Force "${OutPath}\${VMName}" -ErrorAction SilentlyContinue | Out-Null
        }

        # 重命名為正式目錄
        Rename-Item "${OutPath}\${VMName}-Temp" "${OutPath}\${VMName}" -Force
        if (Test-Path "${OutPath}\${VMName}") {
            Out-LogFile -Level Info -Content "Move to ${OutPath}\${VMName} OK" -LogFilePath "${env:Temp}\Restore_${VMName}.log"
        }

        # 複製 LOG 到目的地
        Copy-Item "${env:Temp}\Restore_${VMName}.log" "${OutPath}\${VMName}"
        Rename-Item "${OutPath}\${VMName}\Restore_${VMName}.log" "Restore_${VMName}-$($RestorePoint.CreationTime.ToString('yyyyMMdd_HHmm')).log"
        
        return $true
    }
    else {
        # 匯出失敗，刪除暫存檔
        Out-LogFile -Level Error -Content "【${VMName}】 Export FAIL" -LogFilePath "${env:Temp}\Restore_${VMName}.log"
        if (Test-path "${OutPath}\${VMName}-Temp") {
            Remove-Item -Force -Recurse "${OutPath}\${VMName}-Temp" -ErrorAction SilentlyContinue | Out-Null
        }
        return $false
    }
}

Out-LogFile -Level Info -Content "======== Start Export VMDK" -LogFilePath "${LogPath}"

# 開啟連線
$VeeamBackupServer = $null
# 先檢查是否已連線
try {
    $VeeamBackupServer = (Get-VBRServer -Type Local -ErrorAction SilentlyContinue)
}
catch {
    Connect-VBRServer -Server localhost
    $VeeamBackupServer = (Get-VBRServer -Type Local -ErrorAction SilentlyContinue)
}

if ($null -eq $VeeamBackupServer) {
    # 如果沒有建立連線，回覆錯誤並結束
    Out-LogFile -Level Error -Content "Connection error" -LogFilePath "${LogPath}" -ShowOnConsole
}
elseif (([Veeam.Backup.Core.CBaseSession]::GetRunning() | Where-Object JobType -eq 'RestoreVmFiles').count -gt 0) {
    # 當已經有任務在執行時中斷
    Out-LogFile -Level Error -Content "Some jobs running" -LogFilePath "${LogPath}" -ShowOnConsole
}
else {
    # 重新掃描儲存庫
    $temp = Get-VBRBackupRepository
    Rescan-VBREntity -Entity $VeeamBackupServer -Wait

    foreach ($VMName in $vms) {
        Out-LogFile -Level Info -Content "Start Export 【${VMName}】" -LogFilePath "${env:Temp}\Restore_${VMName}.log"
        Out-LogFile -Level Info -Content "Start Export 【${VMName}】" -LogFilePath "${LogPath}" -ShowOnConsole
        
        $result = Export-VeeamToVMDK -VMName ${VMName} -OutPath ${DestPath}
        
        <#
        # 檢查要匯出的主機是否在本機備份清單中
        if ((Get-VBRJobObject -Job $(Get-VBRJob)).Name -contains ${VMName}) {
            # 有就直接匯出
        }
        else {
            # 如果沒有，跑匯入備份檔並匯出的流程

            # 找到 VBM 檔
            $vbm = $backupMetas | Where-Object FullName -like "*${VMName}*"
            
            # 檢查 VBM 檔是否存在
            if ($vbm.count -gt 0) {
                $xml = (Select-Xml -path $vbm -XPath '/').Node
                
                # 找到 VMK 檔
                $VBKSource = ($xml.BackupMeta.BackupMetaInfo.Storages.Storage | Where-Object FilePath -like '*.vbk' | Sort-Object CreationTime -Descending | Select-Object -First 1).FilePath

                # 替換網路路徑為本機路徑
                $VBKSource = $VBKSource.ToLower()
                $VBKSource = $VBKSource.Replace('\\BKSERVERNAME\BKFOLDER$', 'F:\backup_repo')
                Out-LogFile -Level Info -Content "VBK: ${VBKSource}" -LogFilePath "${env:Temp}\Restore_${VMName}.log" -ShowOnConsole

                if (Test-Path $VBKSource) {
                    # 如果有找到再繼續
                    # 匯入 VBK
                    try {
                        Out-LogFile -Level Info -Content "import ${VBKSource}" -LogFilePath "${env:Temp}\Restore_${VMName}.log"
                        Out-LogFile -Level Info -Content "import ${VBKSource}" -LogFilePath "${LogPath}" -ShowOnConsole
                        Import-VBRBackup -Server $VeeamBackupServer -FileName "${VBKSource}"
                    }
                    catch {
                        Out-LogFile -Level Error -Content "import ${VBKSource} fail" -LogFilePath "${env:Temp}\Restore_${VMName}.log"
                        Out-LogFile -Level Error -Content "import ${VBKSource} fail" -LogFilePath "${LogPath}" -ShowOnConsole
                    }

                    # 開始匯出
                    $result = Export-VeeamToVMDK -VMName ${VMName} -OutPath ${DestPath}
                    if ($result) {
                        # 成功匯出後移除匯入檔
                        #Remove-VBRBackup -Backup $imported -Confirm
                    }
                    else {
                
                    }
                }
                else {
                    Out-LogFile -Level Error -Content "${VMName} VBK Not found" -LogFilePath "${env:Temp}\Restore_${VMName}.log"
                    Out-LogFile -Level Error -Content "${VMName} VBK Not found" -LogFilePath "${LogPath}" -ShowOnConsole
                }
            }
            else {
                Out-LogFile -Level Error -Content "${VMName} BackupMeta Not found" -LogFilePath "${env:Temp}\Restore_${VMName}.log"
                Out-LogFile -Level Error -Content "${VMName} BackupMeta Not found" -LogFilePath "${LogPath}" -ShowOnConsole
            }
            
        }
        #>

        if (Test-Path "${env:Temp}\Restore_${VMName}.log") {
            Move-Item "${env:Temp}\Restore_${VMName}.log" "${LogFolder}" -Force
            Rename-Item "${LogFolder}\Restore_${VMName}.log" "Restore_${VMName}-$(Get-Date -Format 'yyyyMMdd_HHmm').log" -Force
        }
    }

    # 關閉連線
    Disconnect-VBRServer
}

Out-LogFile -Level Info -Content "========   End Export VMDK" -LogFilePath "${LogPath}" -ShowOnConsole
