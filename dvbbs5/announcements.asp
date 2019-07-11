<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!--#include file="inc/ubbcode.asp"-->
<%
'=========================================================
' File: announcements.asp
' Version:5.0
' Date: 2002-9-28
' Script Written by satan
'=========================================================
' Copyright (C) 2001,2002 AspSky.Net. All rights reserved.
' Web: http://www.aspsky.net,http://www.dvbbs.net
' Email: info@aspsky.net,eway@aspsky.net
'=========================================================

dim abgcolor
dim canpostann,caneditann
dim canmodifyusername
dim username
canpostann=false
caneditann=false
canmodifyusername=false

if (master or superboardmaster or boardmaster) and Cint(GroupSetting(25))=1 then
	canpostann=true
else
	canpostann=false
end if
if FoundUserPer and Cint(GroupSetting(25))=1 then
	canpostann=true
elseif FoundUserPer and Cint(GroupSetting(25))=0 then
	canpostann=false
end if

if (master or superboardmaster or boardmaster) and Cint(GroupSetting(26))=1 then
	caneditann=true
else
	caneditann=false
end if
if FoundUserPer and Cint(GroupSetting(26))=1 then
	caneditann=true
elseif FoundUserPer and Cint(GroupSetting(26))=0 then
	caneditann=false
end if

stats="浏览论坛公告"
call nav()
if BoardID=0 or not isInteger(BoardID) then
	BoardID=0
	call head_var(1)
else
	BoardID=clng(BoardID)
	call head_var(2)
end if
if canpostann or caneditann then
	response.write "<p align=center><a href=?action=AddAnn&boardid="&boardid&">发布公告</a> | <a href=?action=EditAnn&boardid="&boardid&">管理公告</a></p>"
end if
if request("action")="AddAnn" then
	call addann()
elseif request("action")="SaveAnn" then
	call saveann()
elseif request("action")="EditAnn" then
	call editann()
elseif request("action")="EditAnnInfo" then
	call EditAnnInfo()
elseif request("action")="SaveEdit" then
	call SaveEdit()
elseif request("action")="delann" then
	call delann()
else
	call main()
end if
if founderr then call dvbbs_error()
call footer()
sub main()
	response.write "<table cellpadding=3 cellspacing=1 align=center class=tableborder1>"
	set rs=server.createobject("adodb.recordset")
	sql="select * from bbsnews where boardid="&BoardID&" order by id desc"
	set rs=conn.execute(sql)
	if rs.eof and rs.bof then
%>
<tr><th align=center valign=top>>> 当前没有任何公告 <<</td></tr>
<tr><td class=tablebody1 valign=top style='LEFT: 0px; WIDTH: 100%; WORD-WRAP: break-word'><br>
<p><blockquote>请进入论坛管理页面来发布一个公告(必须是管理员或者版主)。 <br>当你发布一次公告后，本公告就会自动消失，无需你手动删除！
</blockquote></p></td></tr><tr>
<td class=tablebody1 valign=middle>
<table width=100% border=0 cellpadding=0 cellspacing=0>
<tr><td align=left class=tablebody2>&nbsp;&nbsp;&nbsp;<b>发布人</b>： 本站的默认公告
</td><td align=right class=tablebody2><b>发布时间</b>： <%=Now()%>&nbsp;&nbsp;&nbsp;
</tr></table></td></tr>
<%
	else
	do while not rs.eof
%>
<tr><th align=center valign=top>>> <%=htmlencode(rs("title"))%> <<</td></tr>
<tr><td class=tablebody1 valign=top style='LEFT: 0px; WIDTH: 100%; WORD-WRAP: break-word'><br>
<p><blockquote><%=dvbcode(rs("content"),1)%>
</blockquote></p></td></tr><tr>
<td class=tablebody2 valign=middle>
<table width=100% border=0 cellpadding=0 cellspacing=0>
<tr><td align=left class=tablebody2>&nbsp;&nbsp;&nbsp;<b>发布人</b>： <%=htmlencode(rs("username"))%>
</td><td align=right class=tablebody2><b>发布时间</b>： <%=rs("addtime")%>&nbsp;&nbsp;&nbsp;
</tr></table></td></tr>
<%
	rs.movenext
	loop
	end if
	rs.close
	set rs=nothing
	response.write "</table>"
	call activeonline()
end sub

sub AddAnn()
if not canpostann then
	Errmsg=Errmsg+"<br><li>您没有执行此操作的权限。"
	founderr=true
	exit sub
end if
%>
<form action="?action=SaveAnn" method="post"> 
<table cellpadding=6 cellspacing=1 align=center class=tableborder1>
    <tr>
    <th valign=middle colspan=2>发布论坛公告</th></tr>
    <tr>
    <td class=tablebody1 valign=middle><b>用户名</b></td>
    <td class=tablebody1 valign=middle>
	<input type=text name=username value="<%=membername%>" disabled>
	<INPUT name=username type=hidden value="<%=membername%>"></td></tr>
    <tr>
    <td class=tablebody1 valign=middle><b>版 面</b></td>
    <td class=tablebody1 valign=middle>
<%
sql="select boardid,boardtype from board"
set rs=conn.execute(sql)

response.write "<select name=boardid size=1>"
response.write "<option value=0>论坛首页</option>"

do while not rs.eof
response.write "<option value='"&rs("BoardID")&"'"
if boardid=rs("boardid") then response.write " selected "
response.write ">"&rs("Boardtype")&"</option>"
rs.movenext
loop
set rs=nothing
%>
</select>&nbsp;&nbsp;<%if boardmaster then%>您只能在您管理的版面发布公告<%end if%>
	</td></tr>
    <tr>
    <td class=tablebody1 valign=middle><b>标 题</b></td>
    <td class=tablebody1 valign=middle><INPUT name="title" type=text size=60></td></tr>
    <tr>
    <td class=tablebody1 valign=top width=30%><b>内 容</b></td>
    <td class=tablebody1 valign=middle>
<textarea class="smallarea" cols="60" name="Content" rows="8" wrap="VIRTUAL"></textarea>
                </td></tr>
    <tr>
    <td class=tablebody2 valign=middle colspan=2 align=center><input type=submit name="submit" value="发 布"></td></tr></table>
</form>
<%
end sub

sub SaveAnn()
if not canpostann then
	Errmsg=Errmsg+"<br><li>您没有执行此操作的权限。"
	founderr=true
	exit sub
end if
dim username,title,content
if request("username")="" then
	Errmsg=Errmsg+"<br>"+"<li>请输入发布者。"
	founderr=true
else
	username=checkstr(request("username"))
end if
if request("title")="" then
	Errmsg=Errmsg+"<br>"+"<li>请输入新闻标题。"
	founderr=true
else
	title=checkstr(request("title"))
end if
if request("content")="" then
	Errmsg=Errmsg+"<br>"+"<li>请输入新闻内容。"
	founderr=true
else
	content=checkstr(request("content"))
end if
if founderr then exit sub
set rs=server.createobject("adodb.recordset")
sql="select * from bbsnews"
rs.open sql,conn,1,3
rs.addnew
rs("username")=username
rs("title")=title
rs("content")=content
rs("addtime")=Now()
rs("boardid")=boardid
rs.update
rs.close
set rs=nothing
sucmsg="<li>您已经成功的发布了公告。"
call dvbbs_suc()
end sub

sub EditAnn()
if not caneditann then
	Errmsg=Errmsg+"<br><li>您没有执行此操作的权限。"
	founderr=true
	exit sub
end if
%>
<FORM METHOD=POST ACTION="?action=delann&boardid=<%=boardid%>">
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
<tr>
<th width="5%">ID</th>
<th width="55%">标题</th>
<th width="15%">作者</th>
<th width="20%">日期</th>
<th width="5%" height=25>操作</th>
</tr>
<%
if boardid=0 then
	set rs=conn.execute("select * from bbsnews order by addtime desc")
else
	set rs=conn.execute("select * from bbsnews where boardid="&boardid&" order by addtime desc")
end if
do while not rs.eof
%>
<tr>
<td width="5%" align=center height=25 class=tablebody2><%=rs("id")%></td>
<td width="55%" class=tablebody1><a href="?action=EditAnnInfo&boardid=<%=rs("boardid")%>&id=<%=rs("id")%>"><%=htmlencode(rs("title"))%></a></td>
<td width="15%" align=center class=tablebody2><a href="dispuser.asp?name=<%=htmlencode(rs("username"))%>"><%=htmlencode(rs("username"))%></a></td>
<td width="20%" align=center class=tablebody1><%=rs("addtime")%></td>
<td width="5%" align=center class=tablebody2><input type="checkbox" name="id" value="<%=rs("id")%>"></td>
</tr>
<%
rs.movenext
loop
set rs=nothing
%>
<tr><td class=tablebody2 colspan=5 align=right>请选择要删除的公告，<input type=checkbox name=chkall value=on onclick="CheckAll(this.form)">全选 <input type=submit name=Submit value=执行  onclick="{if(confirm('您确定执行的操作吗?')){this.document.paper.submit();return true;}return false;}"></td></tr>
</table>
</FORM>
<%
end sub

sub EditAnnInfo()
dim trs
if not caneditann then
	Errmsg=Errmsg+"<br><li>您没有执行此操作的权限。"
	founderr=true
	exit sub
end if
if not isnumeric(request("id")) then
	Errmsg=Errmsg+"<br><li>错误的参数。"
	founderr=true
	exit sub
end if
set rs=conn.execute("select * from bbsnews where id="&request("id"))
%>
<form action="?action=SaveEdit&id=<%=request("id")%>" method="post"> 
<table cellpadding=6 cellspacing=1 align=center class=tableborder1>
    <tr>
    <th valign=middle colspan=2>管理论坛公告</th></tr>
    <tr>
    <td class=tablebody1 valign=middle><b>用户名</b></td>
    <td class=tablebody1 valign=middle>
	<input type=text name=username value="<%=rs("username")%>" <%if not master then response.write "disabled"%>>
	<%if not master then%>
	<INPUT name=username type=hidden value="<%=rs("username")%>">
	<%end if%>
	</td></tr>
    <tr>
    <td class=tablebody1 valign=middle><b>版 面</b></td>
    <td class=tablebody1 valign=middle>
<%
sql="select boardid,boardtype from board"
set trs=conn.execute(sql)

response.write "<select name=""boardid"" size=1>"
response.write "<option value=0>论坛首页</option>"

do while not trs.eof
response.write "<option value='"&trs("BoardID")&"'"
if rs("boardid")=trs("boardid") then response.write " selected "
response.write ">"&trs("Boardtype")&"</option>"
trs.movenext
loop
set trs=nothing
%>
</select>&nbsp;&nbsp;<%if boardmaster then%>您只能在您管理的版面发布公告<%end if%>
	</td></tr>
    <tr>
    <td class=tablebody1 valign=middle><b>标 题</b></td>
    <td class=tablebody1 valign=middle><INPUT name="title" type=text size=60 value="<%=rs("title")%>"></td></tr>
    <tr>
    <td class=tablebody1 valign=top width=30%><b>内 容</b></td>
    <td class=tablebody1 valign=middle>
<textarea class="smallarea" cols="60" name="Content" rows="8" wrap="VIRTUAL"><%=replace(replace(rs("content"),"<br>",chr(13)),"&nbsp;"," ")%></textarea>
                </td></tr>
    <tr>
    <td class=tablebody2 valign=middle colspan=2 align=center><input type=submit name="submit" value="发 布"></td></tr></table>
</form>
<%
end sub

sub SaveEdit()
if not caneditann then
	Errmsg=Errmsg+"<br><li>您没有执行此操作的权限。"
	founderr=true
	exit sub
end if
dim username,title,content
if not isnumeric(request("id")) then
	Errmsg=Errmsg+"<br>"+"<li>错误的公告参数。"
	founderr=true
end if
if request("username")="" then
	Errmsg=Errmsg+"<br>"+"<li>请输入发布者。"
	founderr=true
else
	username=checkstr(request("username"))
end if
if request("title")="" then
	Errmsg=Errmsg+"<br>"+"<li>请输入新闻标题。"
	founderr=true
else
	title=checkstr(request("title"))
end if
if request("content")="" then
	Errmsg=Errmsg+"<br>"+"<li>请输入新闻内容。"
	founderr=true
else
	content=checkstr(request("content"))
end if
if founderr then exit sub
set rs=server.createobject("adodb.recordset")
sql="select * from bbsnews where id="&cstr(request("id"))
rs.open sql,conn,1,3
rs("username")=username
rs("title")=title
rs("content")=content
rs("addtime")=Now()
rs("boardid")=boardid
rs.update
rs.close
sucmsg="<li>修改公告成功！"
call dvbbs_suc()
end sub

sub delann()
if not caneditann then
	Errmsg=Errmsg+"<br><li>您没有执行此操作的权限。"
	founderr=true
	exit sub
end if
dim delid
delid=replace(request.form("id"),"'","")
conn.execute("delete from bbsnews where id in ("&delid&")")
sucmsg="<li>删除公告成功。"
call dvbbs_suc()
end sub
%>