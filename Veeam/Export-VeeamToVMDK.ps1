
function Out-LogFile {
    param (
        [Parameter(Mandatory = $true, Position = 0)]
        [ValidateSet('None', 'Info', 'Warn', 'Error')]
        [string] $Level,
        [Parameter(Mandatory = $true, Position = 1)]
        [string] $Content,
        
        [Parameter(Mandatory = $false)]
        [string] $LogFilePath
    )
    $LevelString = ""

    Write-Host -NoNewline "$(Get-Date -Format s) "
    switch ($Level) {
        'None' {
            break;
        }
        'Info' {
            $LevelString = "INFO"
            Write-Host -NoNewline -ForegroundColor White -BackgroundColor Green $LevelString
            Write-Host -NoNewline " "
            break;
        }
        'Warn' {
            $LevelString = "WARN"
            Write-Host -NoNewline -ForegroundColor Black -BackgroundColor Yellow $LevelString
            Write-Host -NoNewline " "
            break;
        }
        'Error' {
            $LevelString = "ERRR"
            Write-Host -NoNewline -ForegroundColor White -BackgroundColor Red $LevelString
            Write-Host -NoNewline " "
            break;
        }
        Default {}
    }
    Write-Host $Content

    if (-not [string]::IsNullOrEmpty($LogFilePath)) {
        if ($Level.Equals('None')) {
            "$(Get-Date -Format s)           ${Content}" | Out-File -FilePath "${LogFilePath}" -Append
        }
        elseif ($Level.Equals('Info')) {
            "$(Get-Date -Format s) ${LevelString} .... ${Content}" | Out-File -FilePath "${LogFilePath}" -Append
        }
        elseif ($Level.Equals('Warn')) {
            "$(Get-Date -Format s) ${LevelString} !!!! ${Content}" | Out-File -FilePath "${LogFilePath}" -Append
        }
        elseif ($Level.Equals('Error')) {
            "$(Get-Date -Format s) ${LevelString} #### ${Content}" | Out-File -FilePath "${LogFilePath}" -Append
        }
        else {
        }
    }
}

function Export-VeeamToVMDK {
    param (
        [Parameter(Mandatory = $true)][string] $VMName,
        [Parameter(Mandatory = $true)][string] $OutPath
    )
    <#
    .Description
    Only tested on Veeam 12
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
    $proc = Start-VBRRestoreVMFiles  -RestorePoint $RestorePoint -server ${restoreServer} -Path "${OutPath}\${VMName}-Temp" -Reason "Batch export"
    if ($proc.result -eq 'Success') {
        # 匯出成功
        $endtime = Get-date
        $timespan = $endtime - $starttime
        Out-LogFile -Level Info -Content "Export OK【${VMName}@$($RestorePoint.CreationTime.tostring('s'))】`n                         Spent $(($timespan.Hours).ToString('00')):$(($timespan.Minutes).ToString('00')):$(($timespan.Seconds).ToString('00'))" -LogFilePath "${env:Temp}\Restore_${VMName}.log"

        # 如果正式目錄已存在舊檔，先刪除
        if (Test-Path "${OutPath}\${VMName}") {
            Out-LogFile -Level Info -Content "Remove old backup ${OutPath}\${VMName}" -LogFilePath "${env:Temp}\Restore_${VMName}.log"
            Remove-Item -Recurse -Force "${OutPath}\${VMName}" -ErrorAction SilentlyContinue | Out-Null
        }

        # 重命名為正式目錄
        Rename-Item "${OutPath}\${VMName}-Temp" "${OutPath}\${VMName}" -Force
        if(Test-Path "${OutPath}\${VMName}"){
            Out-LogFile -Level Info -Content "Move to ${OutPath}\${VMName} OK" -LogFilePath "${env:Temp}\Restore_${VMName}.log"
        }

        # 複製 LOG 到目的地
        Copy-Item "${env:Temp}\Restore_${VMName}.log" "${OutPath}\${VMName}"
        Rename-Item "${OutPath}\${VMName}\Restore_${VMName}.log" "Restore_${VMName}-$($RestorePoint.CreationTime.ToString('yyyyMMddHHmmss')).log"
        
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
