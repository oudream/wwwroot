<%@ Language=VBScript %>
<HTML>
<HEAD>
<TITLE>
IIS日志浏览
</TITLE>
</HEAD>
<BODY>
<% 
dim mLog
dim mBol
set mLog = Server.CreateObject("MSWC.IISLog") 
mlog.OpenLogFile "d:\ex010215.log",1, "",0, ""
Response.Write "<ul>"
for i=1 to 10
  mlog.ReadLogRecord 
  mBol=mlog.AtEndOfLog
  Response.Write "<li>"
  Response.Write "<font color=green>时间:</font>" & mlog.DateTime 
  Response.Write "<br>"
  Response.Write "<font color=green>客户IP:</font>" & mlog.ClientIP
  Response.Write "<br>"
  Response.Write "<font color=green>请求方法:</font>" & mlog.method
  Response.Write "<br>"
  Response.Write "<font color=green>请求目标:</font>" & mlog.URIStem
  Response.Write "<br>"
  Response.Write "<font color=green>协议状态:</font>" & mlog.ProtocolStatus

  'Response.Write "<br>" & mBol
  Response.Write "</li>"
next
'loop until mBol=True
Response.Write "</ul>"
Response.Write i
%>

</BODY>
</HTML>
