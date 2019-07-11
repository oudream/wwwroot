<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!-- #include file="inc/chkinput.asp" -->
<!--#include file="md5.asp"-->
<%
'=========================================================
' File: smallpaper.asp
' Version:5.0
' Date: 2002-9-10
' Script Written by satan
'=========================================================
' Copyright (C) 2001,2002 AspSky.Net. All rights reserved.
' Web: http://www.aspsky.net,http://www.dvbbs.net
' Email: info@aspsky.net,eway@aspsky.net
'=========================================================

dim msg
dim cansmallpaper
cansmallpaper=false
stats="发布小字报"
if boardid=0 then
	founderr=true
	Errmsg=Errmsg+"<br><li>请选择您要发布小字报的版面！"
end if

if cint(Forum_Setting(56))=0 then
	founderr=true
	Errmsg=Errmsg+"<br><li>该版面没有开放小字报功能！"
end if

if Cint(GroupSetting(17))=0 then
	Errmsg=Errmsg+"<br>"+"<li>您没有发布小字报的权限，请<a href=login.asp>登陆</a>或者同管理员联系。"
	founderr=true
else
	if not founduser then
		membername="客人"
	end if
	cansmallpaper=true
end if

if founderr then
	call nav()
	call head_var(1)
	call dvbbs_error()
else
	call nav()
	call head_var(2)
	if request("action")="savepaper" then
		call savepaper()
	else
		call main()
	end if
	call activeonline()
	if founderr then call dvbbs_error()
end if
call footer()
sub main()
conn.execute("delete from smallpaper where datediff('d',s_addtime,Now())>1")
%>
<form action="smallpaper.asp?action=savepaper" method="post"> 
        <table cellpadding=6 cellspacing=1 align=center class=tableborder1>
    <tr>
    <th valign=middle colspan=2>
    请详细填写以下信息(<%=msg%>)</th></tr>
    <tr>
    <td class=tablebody1 valign=middle><b>用户名</b></td>
    <td class=tablebody1 valign=middle><INPUT name=username type=text value="<%=membername%>"> &nbsp; <a href="reg.asp">没有注册？</a></td></tr>
    <tr>
    <td class=tablebody1 valign=middle><b>密 码</b></font></td>
    <td class=tablebody1 valign=middle><INPUT name=password type=password value="<%=memberword%>"> &nbsp; <a href="lostpass.asp">忘记密码？</a></td></tr>
    <tr>
    <td class=tablebody1 valign=middle><b>标 题</b>(最多80字)</td>
    <td class=tablebody1 valign=middle><INPUT name="title" type=text size=60></td></tr>
    <tr>
    <td class=tablebody1 valign=top width=30%>
<b>内 容</b><BR>
在本版发布小字报将您将付<font color="<%=Forum_body(8)%>"><b><%=GroupSetting(46)%></b></font>元费用<br>
<font color="<%=Forum_body(8)%>"><b>48</b></font>小时内发表的小字报将随机抽取<font color="<%=Forum_body(8)%>"><b>5</b></font>条滚动显示于论坛上<br>
<li>HTML标签： <%if Forum_Setting(35)=0 then%>不可用<%else%>允许<%end if%>
<li>UBB 标签： <%if Forum_Setting(34)=0 then%>不可用<%else%>允许<%end if%>
<li>内容不得超过500字
</td>
    <td class=tablebody1 valign=middle>
<textarea class="smallarea" cols="60" name="Content" rows="8" wrap="VIRTUAL"></textarea>
<INPUT name="boardid" type=hidden value="<%=boardid%>">
                </td></tr>
    <tr>
    <td class=tablebody2 valign=middle colspan=2 align=center><input type=submit name="submit" value="发 布"></td></tr></table>
</form>
<%end sub%>
<%
sub savepaper()
dim username
dim password
dim title
dim content
userName=Checkstr(trim(request.form("username")))
PassWord=Checkstr(trim(request.form("password")))
title=Checkstr(trim(request.form("title")))
Content=Checkstr(request.form("Content"))
if chkpost=false then
	ErrMsg=ErrMsg+"<Br>"+"<li>您提交的数据不合法，请不要从外部提交发言。"
	FoundErr=True
end if
if UserName="" then
	ErrMsg=ErrMsg+"<Br>"+"<li>请输入姓名"
	FoundErr=True
end if
if title="" then
	FoundErr=True
	ErrMsg=ErrMsg+"<Br>"+"<li>主题不应为空。"
elseif strLength(title)>80 then
	FoundErr=True
	ErrMsg=ErrMsg+"<Br>"+"<li>主题长度不能超过80"
end if
if content="" then
	ErrMsg=ErrMsg+"<Br>"+"<li>没有填写内容。"
	FoundErr=true
elseif strLength(content)>500 then
	ErrMsg=ErrMsg+"<Br>"+"<li>发言内容不得大于500"
	FoundErr=true
end if

'客人不允许发，验证用户
if not founderr and cansmallpaper then
	if PassWord<>memberword then
		password=md5(password)
	end if
	set rs=server.createobject("adodb.recordset")
	sql="Select userWealth From [User] Where UserName='"&UserName&"' and UserPassWord='"&PassWord&"'"
	rs.open sql,conn,1,3
	if not(rs.eof and rs.bof) then
		if Clng(rs("UserWealth"))<Clng(GroupSetting(46)) then
   			ErrMsg=ErrMsg+"<Br>"+"<li>您没有足够的金钱来发布小字报，快到论坛浇点水吧。"
   			FoundErr=true
		else
			rs("UserWealth")=rs("UserWealth")-Cint(GroupSetting(46))
			rs.update
		end if
	end if
	rs.close
	set rs=nothing
end if
if founderr then
	exit sub
else
	sql="insert into smallpaper (s_boardid,s_username,s_title,s_content) values "&_
		"("&_
		boardid&",'"&_
		username&"','"&_
		title&"','"&_
		content&"')"
		'response.write sql
	conn.execute(sql)
	sucmsg="<li>您成功的发布了小字报。"
	call dvbbs_suc()
end if
end sub
%>