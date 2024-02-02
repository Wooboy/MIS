# ======== SCRIPT SETTINGS     ========
$VBKSourceRoot = 'F:\Backup_Repo'
$TargetDisk = 'D:'
$DestBackupRepo = "${TargetDisk}\01-Backup_Veeam\Backup_Repo"
$LogFolder = "${TargetDisk}\99-Log"
$LogPath = "${LogFolder}\Backup_2nd-$(Get-Date -Format 'yyyyMMdd').log"
# ======== SCRIPT SETTINGS END ========

# ======== CREATE FOLDER     ========
New-Item -ItemType Directory "${DestBackupRepo}" -Force -ErrorAction SilentlyContinue | Out-Null
New-Item -ItemType Directory "${LogFolder}" -Force -ErrorAction SilentlyContinue | Out-Null
# ======== CREATE FOLDER END ========

Import-Module "${PSScriptRoot}\Out-LogFile.ps1"

# 開始複製
$starttime = Get-date
Out-LogFile -Level Info -Content "======== Start Backup Veeam" -LogFilePath "${LogPath}" -ShowOnConsole

# 取得所有 VBM
$backupMetas = Get-ChildItem -File "$VBKSourceRoot" -Include '*.vbm' -Recurse

foreach ($backupMeta in $backupMetas) {
    # 讀取 vbm 內容
    $xml = (Select-Xml -path $backupMeta -XPath '/').Node

    $bkname = $xml.BackupMeta.BackupMetaInfo.Objects.Object.Name
    Out-LogFile -Level Info -Content "Start Backup ${bkname}" -LogFilePath "${LogPath}" -ShowOnConsole

    $storages = $xml.BackupMeta.BackupMetaInfo.Storages.Storage | Sort-Object CreationTime -Descending
    
    # 確保 vbm 內容不為空
    if ($null -ne $storages) {
        # 複製 vbm
        Copy-Item -LiteralPath $backupMeta.FullName $DestBackupRepo

        # 檢查是否要略過備份
        
        # 取得最後一筆 VBK (完整備份)
        $latestvbk = $storages | Where-Object FilePath -like '*.vbk' | Sort-Object CreationTime -Descending | Select-Object -First 1
        $pos = $storages.IndexOf($latestvbk)
        
        # 依序複製 vib, vbk
        for ($i = 0; $i -le $pos; $i++) {
            $source = $storages[$i].FilePath.ToLower().Replace('\\BKSERVERNAME\BKFOLDER$', $VBKSourceRoot)
            if(-not (Test-Path "${source}")){
                Out-LogFile -Level Info -Content "   copying ${source}" -LogFilePath "${LogPath}" -ShowOnConsole
                Copy-Item -LiteralPath $source -Destination $DestBackupRepo -ErrorAction SilentlyContinue
            }
        }
    }
    else {
        Out-LogFile -Level Warn -Content "$($backupMeta.FullName) has no content" -LogFilePath "${LogPath}" -ShowOnConsole
    }
}

# 複製結束
$endtime = Get-date
$timespan = $endtime - $starttime
Out-LogFile -Level Info -Content "========  End Backup Veeam .... Spent $(($timespan.Hours).ToString('00')):$(($timespan.Minutes).ToString('00')):$(($timespan.Seconds).ToString('00'))" -LogFilePath "${LogPath}" -ShowOnConsole

