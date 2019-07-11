<%@ language="VBScript" %>
<%
  Option Explicit
  Const lngMaxFormBytes = 200
  Dim objASPError, blnErrorWritten, strServername, strServerIP, strRemoteIP
  Dim strMethod, lngPos, datNow, strQueryString, strURL
  If Response.Buffer Then
    Response.Clear
    Response.Status = "500 Internal Server Error"
    Response.ContentType = "text/html"
    Response.Expires = 0
  End If
  Set objASPError = Server.GetLastError
%>
<html>
<title>本页无法显示</title>
<META HTTP-EQUIV="Content-Type" Content="text-html; charset=gb2312">
<body bgcolor="FFFFFF">
<table border=1>
<tr><td>错误类型：</td>
<td>
<%
  Dim bakCodepage
  bakCodepage = Session.Codepage
  Session.Codepage = 936
  Response.Write Server.HTMLEncode(objASPError.Category)
  If objASPError.ASPCode > "" Then Response.Write Server.HTMLEncode(", " & objASPError.ASPCode)
  Response.Write Server.HTMLEncode(" (0x" & Hex(objASPError.Number) & ")" ) 
%>
</td></tr><tr><td>错误描述:</td><td>
<%
  If objASPError.ASPDescription > "" Then Response.Write Server.HTMLEncode(objASPError.ASPDescription) 
%>
</td></tr><tr><td>错误定位:</td><td>
<%
  blnErrorWritten = False
  If objASPError.Source > "" Then
    strServername = LCase(Request.ServerVariables("SERVER_NAME"))
    strServerIP = Request.ServerVariables("LOCAL_ADDR")
    strRemoteIP =  Request.ServerVariables("REMOTE_ADDR")
    If (strServername = "localhost" Or strServerIP = strRemoteIP) And objASPError.File <> "?" Then
      Response.Write Server.HTMLEncode(objASPError.File)
      If objASPError.Line > 0 Then Response.Write ", 第 " & objASPError.Line & " 行"
      If objASPError.Column > 0 Then Response.Write ", 第 " & objASPError.Column & " 列"
      Response.Write Server.HTMLEncode(objASPError.Source) & "<br>"
      If objASPError.Column > 0 Then Response.Write String((objASPError.Column - 1), "-") & "^<br>"
      blnErrorWritten = True
    End If
  End If
  If Not blnErrorWritten And objASPError.File <> "?" Then
    Response.Write "<b>"
    Response.Write Server.HTMLEncode(objASPError.File)
    If objASPError.Line > 0 Then Response.Write Server.HTMLEncode(", 第 " & objASPError.Line & " 行")
    If objASPError.Column > 0 Then Response.Write ", 第 " & objASPError.Column & " 列"
    Response.Write "</b><br>"
  End If
%>
</td></tr>
<tr><td>浏览器类型：</td><td>
<%= Request.ServerVariables("HTTP_USER_AGENT") %>
</td></tr>
<tr><td>时间：</td><td>
<%
  datNow = Now()
  Response.Write Server.HTMLEncode(FormatDateTime(datNow, 1) & ", " & FormatDateTime(datNow, 3))
  Session.Codepage = bakCodepage
%>
</td></tr>
</table>
</body>
</html>
