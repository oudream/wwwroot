<!-- #include file="conn.asp" -->
<!-- #include file="inc/const.asp" -->
<%
'=========================================================
' File: favadd.asp
' Version:5.0
' Date: 2002-9-10
' Script Written by satan
'=========================================================
' Copyright (C) 2001,2002 AspSky.Net. All rights reserved.
' Web: http://www.aspsky.net,http://www.dvbbs.net
' Email: info@aspsky.net,eway@aspsky.net
'=========================================================

dim announceid
dim rootid
dim topic
dim url
stats="论坛收藏夹"

if not founduser then
	Errmsg=Errmsg+"<br>"+"<li>您还没有<a href=reg.asp>登录</a>。"
	Founderr=true
end if
if request("id")="" then
	founderr=true
	Errmsg=Errmsg+"<br>"+"<li>请指定相关贴子。"
elseif not isInteger(request("id")) then
	founderr=true
	Errmsg=Errmsg+"<br>"+"<li>非法的贴子参数。"
else
	AnnounceID=request("id")
end if
if founderr then
	call nav()
	call head_var(1)
	call dvbbs_error()
else
	call nav()
	call head_var(2)
	url="dispbbs.asp?"
	url=url & "boardid="&boardid&"&id="&announceid
	call chkurl()
	if founderr then
		call dvbbs_error()
	else
		call favadd()
		if founderr then
			call dvbbs_Error()
		else
			sucmsg="本帖子已经收入您在论坛的<a href=favlist.asp>收藏夹</a>"
			call dvbbs_suc()
		end if
	end if
	call activeonline()
end if
call footer()

sub chkurl()
	sql="select title from topic where topicid="&announceid
	set rs=conn.execute(sql)
	if rs.eof and rs.bof then
		Errmsg=Errmsg+"<br>"+"<li>没有相关贴子。"
		Founderr=true
	else
		topic=rs(0)
	end if
	rs.close
	set rs=nothing
end sub
sub favadd()
	sql="select * from bookmark where url='"&trim(url)&"'"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,3
	if not rs.eof and not rs.bof then
		Errmsg=Errmsg+"<br>"+"<li>该贴子已经在您的论坛收藏夹中。"
		Founderr=true
	else
		rs.addnew
		rs("username")=membername
		rs("topic")=topic
		rs("url")=url
		rs("addtime")=Now()
		rs.update
	end if
	rs.close
	set rs=nothing
end sub
%>