<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!--#include file="inc/usercp.asp"-->
<%
'=========================================================
' File: favlist.asp
' Version:5.0
' Date: 2002-9-11
' Script Written by satan
'=========================================================
' Copyright (C) 2001,2002 AspSky.Net. All rights reserved.
' Web: http://www.aspsky.net,http://www.dvbbs.net
' Email: info@aspsky.net,eway@aspsky.net
'=========================================================

stats="�����ղؼ�"
if not founduser then
	Errmsg=Errmsg+"<br>"+"<li>����û��<a href=login.asp>��½��̳</a>�����ܲ鿴�ղء��������û��<a href=reg.asp>ע��</a>������<a href=reg.asp>ע��</a>��"
	Founderr=true
end if
call nav()
if Founderr then
	call head_var(1)
	call dvbbs_error()
else
	call cp_var()
	if request("action")="delet" then
		call delete()
	else
		call favlist()
	end if
	if founderr then call dvbbs_error()
end if
call activeonline()
call footer()

sub favlist()
set rs=server.createobject("adodb.recordset")
sql="select * from bookmark where username='"&membername&"' order by id desc"
rs.open sql,conn,1,1
if rs.eof and rs.bof then
	Errmsg=Errmsg+"<br>"+"<li>�����ղؼл�û���ղأ��������ղ���ָ̳�����ӣ����ղ��������ݺ󣬱���Ϣ���Զ�ɾ����"
	founderr=true
	exit sub
else
%>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
	                        <tr>
	                        <th valign=top colspan="3">>> �����ղؼ� <<</th></tr>
  <tr>
    <td width="70%" class=tablebody2><B>����</B></td>
    <td width="20%" class=tablebody2><B>ʱ��</B></td>
    <td width="10%" class=tablebody2><B>����</B></td>
  </tr>
<%
	do while not rs.eof
%>
  <tr>
    <td width="70%" class=tablebody1><a href="<%=rs("url")%>"><%=htmlencode(rs("topic"))%></a></td>
    <td width="20%" class=tablebody1><%=rs("addtime")%></td>
    <td width="10%" class=tablebody1><a href="favlist.asp?action=delet&id=<%=rs("id")%>"><img src="pic/a_delete.gif" border=0></a></td>
  </tr>
<%
	rs.movenext
	loop
%>
</table>
			</td></tr>
</table>
<%
end if
rs.close
set rs=nothing
end sub

sub delete()
if isInteger(request("id")) then
	sql="delete from bookmark where username='"&membername&"' and id="&cstr(request("id"))
	conn.execute sql
end if
sucmsg="<li>��ɾ�������ղؼ�����Ӧ��¼��"
call dvbbs_suc()
end sub
%>