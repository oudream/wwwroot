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

stats="您的收藏夹"
if not founduser then
	Errmsg=Errmsg+"<br>"+"<li>您还没有<a href=login.asp>登陆论坛</a>，不能查看收藏。如果您还没有<a href=reg.asp>注册</a>，请先<a href=reg.asp>注册</a>！"
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
	Errmsg=Errmsg+"<br>"+"<li>您的收藏夹还没有收藏，您可以收藏论坛指定贴子，当收藏中有数据后，本信息将自动删除！"
	founderr=true
	exit sub
else
%>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
	                        <tr>
	                        <th valign=top colspan="3">>> 您的收藏夹 <<</th></tr>
  <tr>
    <td width="70%" class=tablebody2><B>标题</B></td>
    <td width="20%" class=tablebody2><B>时间</B></td>
    <td width="10%" class=tablebody2><B>操作</B></td>
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
sucmsg="<li>已删除您的收藏夹中相应纪录。"
call dvbbs_suc()
end sub
%>