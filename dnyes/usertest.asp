<!--#include file="mymefaq5411jkjkh/favorible/showme.asp" -->
<!--#include file="css.asp" -->
 <%
'��ʼ�����ݿ�д�����ݣ�������Ƿ����д��û�
set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT * FROM user where uid= '" & request("uid") & "'"
rs.open sql,conn,1,1
if not (rs.Bof or rs.eof) then
rs.Close
set rs=nothing
response.Write("<table width=100% height=100% border=0 align=center cellpadding=0 cellspacing=0><tr><td align=center><br>���ʺ��ѱ�ע��<br><br>��ѡ���������û���<br><br>�����κ���������<a href="&"emailtome.asp"&"><font color=red>������ϵ</font></a><br><br></td></tr></table>")
response.end
else
response.Write("<table width=100% height=100% border=0 align=center cellpadding=0 cellspacing=0><tr><td align=center>���ʺſ���ע��,�����<br><br></td></tr></table>")
response.end()
end if
%>