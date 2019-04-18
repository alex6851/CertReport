Import-Module PSPKI

#Variables
$TempFile = "C:\Temp\CA_Report.html"
$Today = get-date
$To = "alexander.baker@mrcy.com"
$From = "alexander.baker@mrcy.com"
$SMTPServer = "10.160.60.57"

#Get the CA Name
$CAName = (Get-CA | select Computername).Computername

#Get Details on Issued Certs
$Output = Get-CA | Get-IssuedRequest | select RequestID, CommonName, NotAfter, CertificateTemplate | sort Notafter

#Take the above, and exclude CAExchange Certs, Select the first one, and get an integer value on how many days until the earliest renewal is necessary
$RelevantInfo = ($Output | where-Object {$_.CertificateTemplate -notlike "CAExchange"})
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

$Subject = "PS Report - Issuing CA Info (Next Expiration is $EarliestExpiryInteger from now, $PendingCountStr Requests Pending)"

Send-mailmessage -To $To -From $From -SmtpServer $SMTPServer -Subject $Subject -Body $Body -Attachments $TempFile

Remove-Item $TempFile -force
