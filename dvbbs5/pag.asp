<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/chkinput.asp" -->
<!-- #include file="inc/email.asp" -->
<!-- #include file="inc/ubbcode.asp" -->
<%
'=========================================================
' File: pag.asp
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
dim useremail
dim announce
dim abgcolor
dim usersign
dim sendmail
dim userclass
dim TotalUseTable
usersign=false
stats="打包邮递"
if Cint(GroupSetting(15))=0 then
	Errmsg=Errmsg+"<br>"+"<li>您没有将本帖打包的权限，请<a href=login.asp>登陆</a>或者同管理员联系。"
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
			errmsg=errmsg+"<br>"+"<li>您的Email有错误!</li>"
			founderr=true
		else
			useremail=trim(Request.Form("mail"))
		end if
		call announceinfo()
		if founderr then
			call dvbbs_Error()
		else
			if Forum_Setting(2)=0 then
				errmsg=errmsg+"<br>"+"<li>本论坛不支持发送邮件。</li>"
				founderr=true
				call dvbbs_error()
			elseif Forum_Setting(2)=1 then
				call jmail(useremail,topic,mailbody)
				sucmsg="恭喜您，您的打包邮递发送成功！"
				call dvbbs_suc()
			elseif Forum_Setting(2)=2 then
				call Cdonts(useremail,topic,mailbody)
				sucmsg="恭喜您，您的打包邮递发送成功！"
				call dvbbs_suc()
			elseif Forum_Setting(2)=3 then
				call aspemail(useremail,topic,mailbody)
				sucmsg="恭喜您，您的打包邮递发送成功！"
				call dvbbs_suc()
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
    set rs=conn.execute("select title,PostTable from topic where topicID="&AnnounceID)
	if not(rs.bof and rs.eof) then
		topic=rs(0)
		TotalUseTable=rs(1)
	else
		foundErr = true
		ErrMsg=ErrMsg+"<br>"+"<li>您指定的贴子不存在</li>"
		exit sub
	end if
	rs.close
	mailbody=mailbody &"<style>A:visited {	TEXT-DECORATION: none	}"
	mailbody=mailbody &"A:active  {	TEXT-DECORATION: none	}"
	mailbody=mailbody &"A:hover   {	TEXT-DECORATION: underline overline	}"
	mailbody=mailbody &"A:link 	  {	text-decoration: none;}"
	mailbody=mailbody &"A:visited {	text-decoration: none;}"
	mailbody=mailbody &"A:active  {	TEXT-DECORATION: none;}"
	mailbody=mailbody &"A:hover   {	TEXT-DECORATION: underline overline}"
	mailbody=mailbody &"BODY   {	FONT-FAMILY: 宋体; FONT-SIZE: 9pt;}"
	mailbody=mailbody &"TD	   {	FONT-FAMILY: 宋体; FONT-SIZE: 9pt	}</style>"
	mailbody=mailbody &"<TABLE border=0 width='95%' align=center><TBODY><TR>"
	mailbody=mailbody &"<TD valign=middle align=top>"
	mailbody=mailbody &"-&nbsp;&nbsp;<b>"&Forum_info(0)&"</b>&nbsp;&nbsp;(http://"&request.servervariables("server_name")&replace(request.servervariables("script_name"),"pag.asp","")&"index.asp)<br>"
	mailbody=mailbody &"--&nbsp;&nbsp;<b>"&boardtype&"</b>&nbsp;&nbsp;(http://"&request.servervariables("server_name")&replace(request.servervariables("script_name"),"pag.asp","")&"list.asp?boardid="&boardid&")<br>"
	mailbody=mailbody &"----&nbsp;&nbsp;<b>"&htmlencode(topic)&"</b>&nbsp;&nbsp;(http://"&request.servervariables("server_name")&replace(request.servervariables("script_name"),"pag.asp","")&"dispbbs.asp?boardid="&boardid&"&id="&announceid&")"
	mailbody=mailbody &"</TD></TR></TBODY></TABLE><br><hr>"

	Rs.open "Select b.UserName,b.Topic,b.dateandtime,b.body,u.UserGroupID from "&TotalUseTable&" b inner join [user] u on b.PostUserID=u.userid where b.boardid="&boardid&" and b.rootid="&Announceid&" order by b.announceid",conn,1,1
	do while not rs.eof

	announce=announce &"<TABLE border=0 width='95%' align=center><TBODY><TR>"
	announce=announce &"<TD valign=middle align=top>"
	announce=announce &"--&nbsp;&nbsp;作者："&rs("username")&"<br>"
	announce=announce &"--&nbsp;&nbsp;发布时间："&rs("dateandtime")&"<br><br>"
	announce=announce &"--&nbsp;&nbsp;"&htmlencode(rs("topic"))&"<br>"
	announce=announce &""&dvbcode(rs("body"),rs("UserGroupID"))&""
	announce=announce &"<hr></TD></TR></TBODY></TABLE>"

	rs.movenext
	loop	
	rs.close
	set rs=nothing
	mailbody=mailbody+announce
	mailbody=mailbody &""&Copyright&"&nbsp;&nbsp;"&Version&""
'	response.write mailbody
'	mailbody=""
end sub

sub pag()
%>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1 style="width:60%">
    <form action="pag.asp?action=sendmail&boardid=<%=boardid%>&id=<%=announceid%>" method=post>
    <tr>
    <th align=left valign=middle colspan=2 height=25>
    打包邮递</th></tr>
    <tr>
    <td class=tablebody1 valign=middle colspan=2>
    <b>把本帖打包邮递。</b><br>请正确输入你要邮递的邮件地址！
        </td></tr><tr>
    
    <td class=tablebody1><b>邮递的 Email 地址：</b></td>
    <td class=tablebody1><input type=text size=40 name="mail"></td>
    </tr><tr>
    <td colspan=2 class=tablebody2 align=center><input type=submit value="发 送" name="Submit"></td></tr></table></form>
<%end sub%>