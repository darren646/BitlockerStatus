'Send Log File
Set objMessage = CreateObject("CDO.Message") 
objMessage.Subject = "[Compliance] GSC Romaina Bitlocker Checking" 
objMessage.From = "admin@gsc.huawei.com" 
objMessage.To = "adrian.iosub@huawei.com;baluta.cosmin@huawei.com;victor.pulca@huawei.com;dan.toda@huawei.com;liyongling@huawei.com;gscithelpdesk@huawei.com;yanyouhong@huawei.com;sircgsc@huawei.com" 
objMessage.TextBody = "GSC Romaina Bitlocker Checking."
ystr=year(Now)
mstr=Month(Now)
if len(mstr)<2 then mstr="0"&mstr
dstr=day(Now)
if len(dstr)<2 then dstr="0"&dstr
dateml=ystr&mstr&dstr
objMessage.AddAttachment "C:\workstsstatus\Bitlocker gv status GSC Romania "&dateml&".csv"
'msgbox("C:\workstsstatus\GSC Romania Bitlock gv status"&dateml&".csv")
 
'==This section provides the configuration information for the remote SMTP server.
'==Normally you will only change the server name or IP.
objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 
 
'Name or IP of Remote SMTP Server
objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "172.16.88.144"
 
'Server port (typically 25)
objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25 
 
objMessage.Configuration.Fields.Update
 
'==End remote SMTP server configuration section==
 
objMessage.Send

