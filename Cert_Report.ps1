Install-Module -Name PSPKI -RequiredVersion 3.2.7.0 

Import-Module PSPKI


#Variables
$TempFile = "C:\Temp\CA_Report.html"
$Today = get-date
$To = "alexander.baker@mrcy.com"
$From = "CERT_REPORT@mrcy.com"
$SMTPServer = "mail.mrcy.com"

#Get the CA Name
$CAName = (Get-CA | select Computername).Computername

#Get Details on Issued Certs
$Output = Get-CA | Get-IssuedRequest -Filter "NotAfter -ge $(Get-Date)", "NotAfter -le $((Get-Date).AddMonths(6))" | select RequestID, CommonName, NotAfter, CertificateTemplate | sort Notafter

#Take the above, and exclude CAExchange Certs, Select the first one, and get an integer value on how many days until the earliest renewal is necessary
$RelevantInfo = ($Output | where-Object {$_.CertificateTemplate -notlike "CAExchange" -and $_.CertificateTemplate -notlike "1.3.6.1.4.1.311.21.8.1409591.2629332.1724303.11985699.9815899.150.16496907.14499982" -and $_.CertificateTemplate -notlike "DomainController"})
$EarliestExpiryInteger = ([math]::abs(($Today - ($RelevantInfo[0].Notafter)).Days)).ToString()

#Write the Relevant Info to a temp file
$RelevantInfo | ConvertTo-HTML | out-file $TempFile

#Get Details on Pending Requests
$Pending = Get-CA | Get-PendingRequest

#Get number of pending requests - If pending requests is null, then PendingCount is left at zero
If ($Pending){$PendingCount = ($Pending | Measure-Object).count}
Else {
$PendingCount = 0
$Pending = "`r`nNone"
} #End Else
$PendingCountStr = $PendingCount.ToString()

#Make the mail body
$Body = "See Attached"

if ($EarliestExpiryInteger -le 7)
{
    $Subject = "WARNING Cert expires $EarliestExpiryInteger days from now, $PendingCountStr Requests Pending"
}
else {
    $Subject = "Next Expiring Cert is $EarliestExpiryInteger from now, $PendingCountStr Requests Pending"
}

Send-mailmessage -To $To -From $From -SmtpServer $SMTPServer -Subject $Subject -Body $Body -Attachments $TempFile

Remove-Item $TempFile -force


