<!--#include file="mymefaq5411jkjkh/favorible/showme.asp" -->
<!--#include file="css.asp" -->
 <%
'开始向数据库写入数据，并检测是否已有此用户
set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT * FROM user where uid= '" & request("uid") & "'"
rs.open sql,conn,1,1
if not (rs.Bof or rs.eof) then
rs.Close
set rs=nothing
response.Write("<table width=100% height=100% border=0 align=center cellpadding=0 cellspacing=0><tr><td align=center><br>此帐号已被注册<br><br>请选择其他的用户名<br><br>如有任何问题请与<a href="&"emailtome.asp"&"><font color=red>我们联系</font></a><br><br></td></tr></table>")
response.end
else
response.Write("<table width=100% height=100% border=0 align=center cellpadding=0 cellspacing=0><tr><td align=center>此帐号可以注册,请继续<br><br></td></tr></table>")
response.end()
end if
%>