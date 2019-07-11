<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/chkinput.asp" -->
<!-- #include file="inc/email.asp" -->
<%
'=========================================================
' File: sendpage.asp
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
dim sendmail
stats="发送帖子"
if Cint(GroupSetting(15))=0 then
	Errmsg=Errmsg+"<br>"+"<li>您没有将本页面发送的权限，请<a href=login.asp>登陆</a>或者同管理员联系。"
	founderr=true
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
	if request("action")="sendmail" then
		if IsValidEmail(trim(Request.Form("mail")))=false then
  			errmsg=errmsg+"<br>"+"<li>您的Email有错误。</li>"
  			founderr=true
		else
			email=trim(Request.Form("mail"))
		end if
		if request("postname")="" then
   			errmsg=errmsg+"<br>"+"<li>请输入您的姓名。</li>"
   			founderr=true
		else
			postname=request("postname")
		end if
		if request("incept")="" then
   			errmsg=errmsg+"<br>"+"<li>请输入收件人姓名。</li>"
   			founderr=true
		else
			incepts=request("incept")
		end if
		if request("content")="" then
   			errmsg=errmsg+"<br>"+"<li>邮件内容不能为空。</li>"
   			founderr=true
		else
			content=request("content")
		end if
		if founderr then
			dvbbs_error()
		else
			call announceinfo()
			if founderr then
				dvbbs_error()
			else
			if Forum_Setting(2)=0 then
				errmsg=errmsg+"<br>"+"<li>本论坛不支持发送邮件。</li>"
   				founderr=true
				call dvbbs_error()
			elseif Forum_Setting(2)=1 then
				call jmail(email,topic,mailbody)
				sucmsg="恭喜您，您的页面发送成功！"
				call dvbbs_suc()
			elseif Forum_Setting(2)=2 then
				call Cdonts(email,topic,mailbody)
				sucmsg="恭喜您，您的页面发送成功！"
				call dvbbs_suc()
			elseif Forum_Setting(2)=3 then
				call aspemail(email,topic,mailbody)
				sucmsg="恭喜您，您的页面发送成功！"
				call dvbbs_suc()
			end if
			end if
		end if
	else
		call pag()
	end if
	call activeonline()
end if
call footer()

sub announceinfo()
   	set rs=server.createobject("adodb.recordset")
    set rs=conn.execute("select title from topic where topicID="&AnnounceID)
	if not(rs.bof and rs.eof) then
		topic="您的朋友"&postname&"给您发来了一个"&Forum_info(0)&"上的贴子"
	else
		foundErr = true
		ErrMsg=ErrMsg+"<br>"+"<li>您指定的贴子不存在</li>"
		exit sub
	end if
	rs.close
	set rs=nothing
	mailbody=mailbody &"<style>A:visited {	TEXT-DECORATION: none	}"
	mailbody=mailbody &"A:active  {	TEXT-DECORATION: none	}"
	mailbody=mailbody &"A:hover   {	TEXT-DECORATION: underline overline	}"
	mailbody=mailbody &"A:link 	  {	text-decoration: none;}"
	mailbody=mailbody &"A:visited {	text-decoration: none;}"
	mailbody=mailbody &"A:active  {	TEXT-DECORATION: none;}"
	mailbody=mailbody &"A:hover   {	TEXT-DECORATION: underline overline}"
	mailbody=mailbody &"BODY   {	FONT-FAMILY: 宋体; FONT-SIZE: 9pt;}"
	mailbody=mailbody &"TD	   {	FONT-FAMILY: 宋体; FONT-SIZE: 9pt	}</style>"
	mailbody=mailbody &"<TABLE border=0 width='95%' align=center><TBODY><TR><TD>"
	mailbody=mailbody &""&incepts&"，您好：<br><br>"
	mailbody=mailbody &"您的朋友"&postname&"给您发来了一个"&Forum_info(0)&"--"&boardtype&"上的贴子<BR><br>"
	mailbody=mailbody &"标题是："&htmlencode(topic)&"<br><br>"
	mailbody=mailbody &""&htmlencode(content)&"<br><br>"
	mailbody=mailbody &"您可以到<a href=http://"&request.servervariables("server_name")&replace(request.servervariables("script_name"),"sendpage.asp","")&"dispbbs.asp?boardid="&boardid&"&id="&Announceid&">"&topic&"</a>这里浏览这个贴子<br>"
	mailbody=mailbody &""&Copyright&"&nbsp;&nbsp;"&Version&""
	mailbody=mailbody &"</TD></TR></TBODY></TABLE>"
'	response.write mailbody
'	mailbody=""
	end sub

sub pag()
%>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
    <form action="sendpage.asp?action=sendmail&boardid=<%=boardid%>&id=<%=announceid%>" method=post>
    <tr>
    <th valign=middle colspan=2>发送邮件给朋友</th></tr>
    <tr>
    <td  class=tablebody1 valign=middle colspan=2>
    <b>通过邮件发送本帖子给您的朋友。</b>　下列所有项必填，并请输入正确的邮件地址！
    <br>你可以添加一些自己的信息在下面的内容框内。至于这个帖子的主题和 URL 你可以不必写，因为本程序会在发送的 Email 中自动添加的！
        </td></tr>
    <tr>
    <td class=tablebody1><b>您的姓名：</b></td>
    <td class=tablebody1><input type=text size=40 name="postname"></td>
    </tr><tr>
    <td class=tablebody1><b>您朋友的名字：</b></td>
    <td class=tablebody1><input type=text size=40 name="incept"></td>
    </tr><tr>
    <td class=tablebody1><b>您朋友的 Email：</b></td>
    <td class=tablebody1><input type=text size=40 name="mail"></td>
    </tr><tr>
    <td class=tablebody1><b>消息内容：</b></td>
    <td class=tablebody1><textarea name="content" cols="55" rows="6">我想你对 '<%=Forum_info(0)%>' 的这个帖子内容会感兴趣的！请去看看！</textarea></td>
    </tr><tr>
    <td colspan=2 class=tablebody2 align=center><input type=submit value="发 送" name="Submit"></td></tr></table></form>
<%end sub%>
