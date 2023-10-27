# 指定來源
$CertFilePath = Read-Host "輸入 fullchain.pem 路徑"
#$CertFilePath = 'C:\TEST\RSA-fullchain.pem'


if (-not (Test-Path $CertFilePath)) {
    Write-Host -ForegroundColor Yellow -NoNewline "$CertFilePath"
    Write-Host ' 檔案不存在'
    Pause
    exit
}

# 判斷輸入檔案是否為 PEM
if ((Get-Item $CertFilePath).Extension -ne ".pem") {
    Write-Host -ForegroundColor Yellow -NoNewline "$CertFilePath"
    Write-Host ' 輸入檔案格式不正確'
    Pause
    exit
}
# 取得父目錄做為輸出目錄
$outRoot = Split-Path $CertFilePath

# 如果目錄中已存在 RootCA.cer，先刪掉
if (Test-Path "${outRoot}\RootCA.cer") {
    Remove-Item "${outRoot}\RootCA.cer" -Force -ErrorAction SilentlyContinue | Out-Null
}

$i = 0

# 讀取 pem 轉為 X509 物件
$cert = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($CertFilePath)

if(-not $cert.Subject.Equals("CN=cylee.com")){
    Write-Host "選取的憑證非 CN=cylee.com"
}
# 如果簽署者與申請者相同，報錯
if($cert.Issuer -eq $cert.Subject){
    Write-Host "ERROR Issuer equals Subject"
}

$chain = [System.Security.Cryptography.X509Certificates.X509Chain]::new()
# 不檢查憑證是否已撤銷
$chain.ChainPolicy.RevocationMode = [System.Security.Cryptography.X509Certificates.X509RevocationMode]::NoCheck
# 取得憑證鍊
$chain.Build($cert) | Out-Null
if($chain.ChainElements.Count -gt 1){
    foreach($chainElements in $chain.ChainElements){
        # 略過第一張憑證(自身)
        if($i -eq 0){
            $i++
            continue
        }
        Write-Host "Round ${i}"
        $issuer = $chainElements.Certificate

        #Export-Certificate -Cert $issuer -FilePath "${outRoot}\${i}.der" -Type CERT

        $der = $issuer.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Cert)
        '-----BEGIN CERTIFICATE-----' | Out-File -FilePath "${outRoot}\RootCA.cer" -Encoding ascii -Append
        $pem = [System.Convert]::ToBase64String($der, [System.Base64FormattingOptions]::InsertLineBreaks)
        $pem | Out-File -FilePath "${outRoot}\RootCA.cer" -Encoding ascii -Append
        '-----END CERTIFICATE-----' | Out-File -FilePath "${outRoot}\RootCA.cer" -Encoding ascii -Append
        $i++
    }
}
