<!--#include file="dbm_adminconn.asp" -->
<%
'########��Ȩ���� - (����)���ſƼ����޹�˾www.dnyes.com

'=========================================================================================
'--------------------------------------ϵͳ��ʼ��-----------------------------------------
'=========================================================================================

On Error Resume Next

Dim ObT(13,2)
ObT(0,0) = "Scripting.FileSystemObject"
  ObT(0,2) = "�ļ��������"
ObT(1,0) = "wscript.shell"
  ObT(1,2) = "������ִ�����"
ObT(2,0) = "ADOX.Catalog"
  ObT(2,2) = "ACCESS�������"
ObT(3,0) = "JRO.JetEngine"
  ObT(3,2) = "ACCESSѹ�����"
ObT(4,0) = "Scripting.Dictionary"
  ObT(4,2) = "�������ϴ��������"
ObT(5,0) = "Adodb.connection"
  ObT(5,2) = "���ݿ��������"
ObT(6,0) = "Adodb.Stream"
  ObT(6,2) = "�������ϴ����"
ObT(7,0) = "SoftArtisans.FileUp"
  ObT(7,2) = "SA-FileUp �ļ��ϴ����"
ObT(8,0) = "LyfUpload.UploadFile"
  ObT(8,2) = "���Ʒ��ļ��ϴ����"
ObT(9,0) = "Persits.Upload.1"
  ObT(9,2) = "ASPUpload �ļ��ϴ����"
ObT(10,0) = "JMail.SmtpMail"
  ObT(10,2) = "JMail �ʼ��շ����"
ObT(11,0) = "CDONTS.NewMail"
  ObT(11,2) = "����SMTP�������"
ObT(12,0) = "SmtpMail.SmtpMail.1"
  ObT(12,2) = "SmtpMail�������"
ObT(13,0) = "Microsoft.XMLHTTP"
  ObT(13,2) = "���ݴ������"

For i=0 To 13
	Set T=Server.CreateObject(ObT(i,0))
	If -2147221005 <> Err Then
	  IsObj=True
	Else
	  IsObj=false
	  Err.Clear
	End If
	Set T=Nothing
	ObT(i,1)=IsObj
Next



Function ServerInfo()
  SI="<br><table width='500' border='1' cellspacing='0' cellpadding='0' align='center'>"
  SI=SI&"<tr><td height='30' colspan='3' align='center' bgcolor='menu'>�����������Ϣ</td></tr>"
  SI=SI&"<tr align='center'><td height='25' width='200'>������CPU����</td>&nbsp;<td>&nbsp;</td><td>"&Request.ServerVariables("NUMBER_OF_PROCESSORS")&"&nbsp;</td></tr>"
  SI=SI&"<tr align='center'><td height='25' width='200'>����������ϵͳ</td><td>&nbsp;</td><td>"&Request.ServerVariables("OS")&"&nbsp;</td></tr>"
  SI=SI&"<tr align='center'><td height='25' width='200'>WEB�������汾</td><td>&nbsp;</td><td>"&Request.ServerVariables("SERVER_SOFTWARE")&"&nbsp;</td></tr>"
  For i=0 To 13
    SI=SI&"<tr align='center'><td height='25' width='200'>"&ObT(i,0)&"</td><td>"&ObT(i,1)&"</td><td>"&ObT(i,2)&"</td></tr>"
  Next
  Response.Write SI
End Function

ServerInfo()
%>
