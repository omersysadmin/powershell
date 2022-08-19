$from = "xolanindembele@gmail.com"
$to = "omerandriaminah@gmail.com"
$smtpserver = "smtp.gmail.com"
$smtpport = "587"
$subject =  "This is a test mail from powershell"
$body = "this is the body of the test"

Send-MailMessage -From $from -to $to -Subject $subject -body $body -SmtpServer $smtpserver -Port $smtpport -UseSsl -Credential (Get-Credential) 