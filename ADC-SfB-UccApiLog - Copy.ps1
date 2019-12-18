 Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned -Force
 
 $To=@('ADC_SFB@hclo365.mail.onmicrosoft.com', 'adc_sfb@hcl.com')
If (Test-Path 'hkcu:\Software\Microsoft\Office\15.0\Outlook\Profiles') {



    $from=Get-ItemProperty 'hkcu:\Software\Microsoft\Office\15.0\Outlook\Profiles\*\9375CFF0413111d3B88A00104B2A6676\*' |
	Where{$_.'Account Name' -notmatch 'Outlook Address Book|Outlook Data File' -and $_.'Account Name' -like '*@*'} |
	Select  'account name' -ExpandProperty 'account name'

}elseif(Test-Path 'hkcu:\Software\Microsoft\Office\16.0\Outlook\Profiles'){


    $from=Get-ItemProperty 'hkcu:\Software\Microsoft\Office\16.0\Outlook\Profiles\*\9375CFF0413111d3B88A00104B2A6676\*' |
	Where{$_.'Account Name' -notmatch 'Outlook Address Book|Outlook Data File' -and $_.'Account Name' -like '*@*'} |
	Select  'account name' -ExpandProperty 'account name'

}else{
	Write-Host 'Outlook not found'
	exit
}

$a = $from.split('@')
$from1=$a[0]  

 Compress-Archive -Path E:\Users\$from1\AppData\Local\Microsoft\Office\16.0\Lync\Tracing\Lync-UccApi-0.UccApilog.bak -DestinationPath E:\Users\$from1\AppData\Local\Microsoft\Office\16.0\Lync\Tracing\lynclogs.zip -CompressionLevel Optimal -Force
 
$SMTPServer="mail.hcl.com"
$subject="SIP Log file-UccApiLog for User $from"
$body="Hi, Please find SfB Client Tracing Log file')"
Send-MailMessage -From $from -to $To -Subject $Subject -body $body -SmtpServer $SMTPServer -Attachments "E:\Users\$from1\AppData\Local\Microsoft\Office\16.0\Lync\Tracing\lynclogs.zip" 
 
 
 