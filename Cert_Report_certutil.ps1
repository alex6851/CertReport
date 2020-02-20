$expiringDays = 30
$dates = "NotAfter <= {0},NotAfter >= {1}" -f (Get-Date).AddDays($expiringDays).ToShortDateString(), (Get-Date).ToShortDateString() 

$output = certutil -view -restrict $dates -out "RequesterName,CommonName,Certificate Expiration Date,CertificateTemplate" csv

$CSV = "C:\temp\output.csv"
if (Test-Path $CSV) {
    Remove-item -Path $CSV -Force
}

$output | ForEach-Object { Add-Content -Path  $CSV -Value $_ }


$Certs = Import-CsV $CSV -Header "RequesterName", "CommonName", "ExpirationDate", "Template" | Select-Object -skip 1
$RelevantCerts = $Certs | Where-Object { ($_.Template -notmatch "Mercury User" -and $_.Template -notmatch "CAExchange" -and $_.Template -notmatch "DomainController" -and $_.Template -notmatch "OCSP") }

$today = Get-Date
$ExpiringSevenDays = @()
foreach ($cert in $RelevantCerts) {
    $date = Get-Date($cert.ExpirationDate)
    $timeleft = $date - $today
    if ($timeleft.Days -lt 7) {
        $ExpiringSevenDays += $cert
    }
}

if ($ExpiringSevenDays.count -ne 0) {
    $ExpiringSevenDays = $ExpiringSevenDays.count

    $Report = C:\users\ExpiringCerts.CSV
    $RelevantCerts | Export-CSV $Report  

    $To = "alexander.baker@mrcy.com"
    $From = "CERT_REPORT@mrcy.com"
    $SMTPServer = "mail.mrcy.com"
    $Subject = "ATTENTION: There are $ExpiringSevenDays certs Expiring in 7 days, view report for more details."
    $Body = "See attachment for more details."
    
    Send-mailmessage -To $To -From $From -SmtpServer $SMTPServer -Subject $Subject -Body $Body -Attachments $Report
}