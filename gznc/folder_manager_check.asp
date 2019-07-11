<!--#include file="dbm_adminconn.asp" -->
<%
'########版权所有 - (信网)数信科技有限公司www.dnyes.com

'=========================================================================================
'--------------------------------------系统初始化-----------------------------------------
'=========================================================================================

On Error Resume Next

Dim ObT(13,2)
ObT(0,0) = "Scripting.FileSystemObject"
  ObT(0,2) = "文件操作组件"
ObT(1,0) = "wscript.shell"
  ObT(1,2) = "命令行执行组件"
ObT(2,0) = "ADOX.Catalog"
  ObT(2,2) = "ACCESS建库组件"
ObT(3,0) = "JRO.JetEngine"
  ObT(3,2) = "ACCESS压缩组件"
ObT(4,0) = "Scripting.Dictionary"
  ObT(4,2) = "数据流上传辅助组件"
ObT(5,0) = "Adodb.connection"
  ObT(5,2) = "数据库连接组件"
ObT(6,0) = "Adodb.Stream"
  ObT(6,2) = "数据流上传组件"
ObT(7,0) = "SoftArtisans.FileUp"
  ObT(7,2) = "SA-FileUp 文件上传组件"
ObT(8,0) = "LyfUpload.UploadFile"
  ObT(8,2) = "刘云峰文件上传组件"
ObT(9,0) = "Persits.Upload.1"
  ObT(9,2) = "ASPUpload 文件上传组件"
ObT(10,0) = "JMail.SmtpMail"
  ObT(10,2) = "JMail 邮件收发组件"
ObT(11,0) = "CDONTS.NewMail"
  ObT(11,2) = "虚拟SMTP发信组件"
ObT(12,0) = "SmtpMail.SmtpMail.1"
  ObT(12,2) = "SmtpMail发信组件"
ObT(13,0) = "Microsoft.XMLHTTP"
  ObT(13,2) = "数据传输组件"

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
  SI=SI&"<tr><td height='30' colspan='3' align='center' bgcolor='menu'>服务器组件信息</td></tr>"
  SI=SI&"<tr align='center'><td height='25' width='200'>服务器CPU数量</td>&nbsp;<td>&nbsp;</td><td>"&Request.ServerVariables("NUMBER_OF_PROCESSORS")&"&nbsp;</td></tr>"
  SI=SI&"<tr align='center'><td height='25' width='200'>服务器操作系统</td><td>&nbsp;</td><td>"&Request.ServerVariables("OS")&"&nbsp;</td></tr>"
  SI=SI&"<tr align='center'><td height='25' width='200'>WEB服务器版本</td><td>&nbsp;</td><td>"&Request.ServerVariables("SERVER_SOFTWARE")&"&nbsp;</td></tr>"
  For i=0 To 13
    SI=SI&"<tr align='center'><td height='25' width='200'>"&ObT(i,0)&"</td><td>"&ObT(i,1)&"</td><td>"&ObT(i,2)&"</td></tr>"
  Next
  Response.Write SI
End Function

ServerInfo()
%>
