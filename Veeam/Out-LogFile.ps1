function Out-LogFile {
    param (
        [Parameter(Mandatory = $true, Position = 0)]
        [ValidateSet('None', 'Info', 'Warn', 'Error')]
        [string] $Level,
        [Parameter(Mandatory = $true, Position = 1)]
        [string] $Content,
        [Parameter(Mandatory = $false)]
        [switch] $ShowOnConsole = $false,
        [Parameter(Mandatory = $false)]
        [string] $LogFilePath
    )
    $LevelString = ""

    $ShowOnConsole = $ShowOnConsole -or (-not [string]::IsNullOrEmpty($LogFilePath))

    if ($ShowOnConsole) {
        Write-Host -NoNewline "$(Get-Date -Format s) "
    }
        
    switch ($Level) {
        'None' {
            break;
        }
        'Info' {
            $LevelString = "INFO"
            if ($ShowOnConsole) {
                Write-Host -NoNewline -ForegroundColor White -BackgroundColor Green $LevelString
                Write-Host -NoNewline " .... "
            }
            break;
        }
        'Warn' {
            $LevelString = "WARN"
            if ($ShowOnConsole) {
                Write-Host -NoNewline -ForegroundColor Black -BackgroundColor Yellow $LevelString
                Write-Host -NoNewline " !!!! "
            }
            break;
        }
        'Error' {
            $LevelString = "ERRR"
            if ($ShowOnConsole) {
                Write-Host -NoNewline -ForegroundColor White -BackgroundColor Red $LevelString
                Write-Host -NoNewline " #### "
            }
            break;
        }
        Default {}
    }
    if ($ShowOnConsole) {
        Write-Host $Content
    }
    
    if (-not [string]::IsNullOrEmpty($LogFilePath)) {
        if ($Level.Equals('None')) {
            "$(Get-Date -Format s)           ${Content}" | Out-File -FilePath "${LogFilePath}" -Append
        }
        elseif ($Level.Equals('Info')) {
            "$(Get-Date -Format s) ${LevelString} .... ${Content}" | Out-File -FilePath "${LogFilePath}"
        }
        elseif ($Level.Equals('Warn')) {
            "$(Get-Date -Format s) ${LevelString} !!!! ${Content}" | Out-File -FilePath "${LogFilePath}"
        }
        elseif ($Level.Equals('Error')) {
            "$(Get-Date -Format s) ${LevelString} #### ${Content}" | Out-File -FilePath "${LogFilePath}"
        }
        else {
        }
    }
}
