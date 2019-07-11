<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<%
'=========================================================
' File: report.asp
' Version:5.0
' Date: 2002-9-10
' Script Written by satan
'=========================================================
' Copyright (C) 2001,2002 AspSky.Net. All rights reserved.
' Web: http://www.aspsky.net,http://www.dvbbs.net
' Email: info@aspsky.net,eway@aspsky.net
'=========================================================

dim announceid
dim username
dim rootid
dim topic
dim mailbody
dim email
dim content
dim postname
dim incepts
dim announce
stats="报告有问题的帖子"
if request("id")="" then
	founderr=true
	Errmsg=Errmsg+"<br>"+"<li>请指定相关贴子。"
elseif not isInteger(request("id")) then
	founderr=true
	Errmsg=Errmsg+"<br>"+"<li>非法的贴子参数。"
else
	AnnounceID=request("id")
end if
if not founduser then
	founderr=true
	Errmsg=Errmsg+"<br>"+"<li>请<a href=login.asp>登陆</a>后进行相关操作。"
end if
if founderr then
	call nav()
	call head_var(1)
	call dvbbs_error()
else
	call nav()
	call head_var(2)
	if request("action")="send" then
		call announceinfo()
	else
		call pag()
	end if
	call activeonline()
	if founderr then call dvbbs_error()
end if
call footer()

sub announceinfo()
	dim body
	dim writer
	dim incept
	dim topic,topic_1
	body=checkStr(request("content"))
	writer=membername
	incept=checkStr(request("boardmaster"))
	sql="select title from topic where TopicID="&AnnounceID
	set rs=conn.execute(sql)
	if not(rs.bof and rs.eof) then
		topic_1=rs(0)
		topic="报告有问题的帖子"
		body=body & "您可以到http://"&request.servervariables("server_name")&replace(request.servervariables("script_name"),"report.asp","")&"dispbbs.asp?boardid="&boardid&"&id="&Announceid&"这里浏览这个贴子"
	else
		foundErr = true
		ErrMsg=ErrMsg+"<br>"+"<li>您指定的贴子不存在</li>"
		exit sub
	end if
	rs.close
	sql="insert into message (incept,sender,title,content,sendtime,flag,issend) values ('"&incept&"','"&membername&"','"&topic&"','"&body&"',Now(),0,1)"
	conn.execute(sql)
	sucmsg="<li>报告发送成功！"
	call dvbbs_suc()
	set rs=nothing
end sub

sub pag()
  	sql="select boardmaster from board where "&guestlist&" boardID="&cstr(boardid)
	set rs=server.createobject("adodb.recordset")
 	rs.open sql,conn,1,1
 	if not(rs.bof and rs.eof) then
		boardmasterlist=rs(0)
		master_2=""
		if trim(boardmasterlist)<>"" then
			master_1=split(boardmasterlist, "|")
			for i = 0 to ubound(master_1)
			master_2=""&master_2&"<option value="""&master_1(i)&""">"&master_1(i)&"</option>&nbsp;"
			next
		else
			master_2="无版主"
		end if
		'response.write master_2
		'response.end
	else
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>您选择的版面不存在或者您没有权限察看该版面。"
		exit sub
	end if
	rs.close
	set rs=nothing
%>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
        
    <form action="report.asp?action=send&boardid=<%=boardid%>&id=<%=announceid%>" method=post>
    <tr>
    <th valign=middle colspan=2 height=25>发送报告给版主</th></tr>
    <tr>
    <td class=tablebody1><b>报告发送给哪个版主：</b></td>
    <td class=tablebody1><select name=boardmaster size=1><%=master_2%></select></td>
    </tr>
	<tr>
    <td class=tablebody1><b>消息内容：</b><br>垃圾贴、广告贴、非法贴等。。。
非必要情况下不要使用这项功能！</td>
    <td  class=tablebody1><textarea name="content" cols="55" rows="6">管理员，您好，由于如下原因，我向你报告这有问题的贴子：</textarea></td>
    </tr><tr>
    <td colspan=2  class=tablebody2 align=center><input type=submit value="发 送" name="Submit"></table></td></form></tr></table>
<%end sub%>
