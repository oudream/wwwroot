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
stats="��������"
if Cint(GroupSetting(15))=0 then
	Errmsg=Errmsg+"<br>"+"<li>��û�н���ҳ�淢�͵�Ȩ�ޣ���<a href=login.asp>��½</a>����ͬ����Ա��ϵ��"
	founderr=true
end if

if request("id")="" then
	founderr=true
	Errmsg=Errmsg+"<br>"+"<li>��ָ��������ӡ�"
elseif not isInteger(request("id")) then
	founderr=true
	Errmsg=Errmsg+"<br>"+"<li>�Ƿ������Ӳ�����"
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
  			errmsg=errmsg+"<br>"+"<li>����Email�д���</li>"
  			founderr=true
		else
			email=trim(Request.Form("mail"))
		end if
		if request("postname")="" then
   			errmsg=errmsg+"<br>"+"<li>����������������</li>"
   			founderr=true
		else
			postname=request("postname")
		end if
		if request("incept")="" then
   			errmsg=errmsg+"<br>"+"<li>�������ռ���������</li>"
   			founderr=true
		else
			incepts=request("incept")
		end if
		if request("content")="" then
   			errmsg=errmsg+"<br>"+"<li>�ʼ����ݲ���Ϊ�ա�</li>"
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
				errmsg=errmsg+"<br>"+"<li>����̳��֧�ַ����ʼ���</li>"
   				founderr=true
				call dvbbs_error()
			elseif Forum_Setting(2)=1 then
				call jmail(email,topic,mailbody)
				sucmsg="��ϲ��������ҳ�淢�ͳɹ���"
				call dvbbs_suc()
			elseif Forum_Setting(2)=2 then
				call Cdonts(email,topic,mailbody)
				sucmsg="��ϲ��������ҳ�淢�ͳɹ���"
				call dvbbs_suc()
			elseif Forum_Setting(2)=3 then
				call aspemail(email,topic,mailbody)
				sucmsg="��ϲ��������ҳ�淢�ͳɹ���"
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
		topic="��������"&postname&"����������һ��"&Forum_info(0)&"�ϵ�����"
	else
		foundErr = true
		ErrMsg=ErrMsg+"<br>"+"<li>��ָ�������Ӳ�����</li>"
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
	mailbody=mailbody &"BODY   {	FONT-FAMILY: ����; FONT-SIZE: 9pt;}"
	mailbody=mailbody &"TD	   {	FONT-FAMILY: ����; FONT-SIZE: 9pt	}</style>"
	mailbody=mailbody &"<TABLE border=0 width='95%' align=center><TBODY><TR><TD>"
	mailbody=mailbody &""&incepts&"�����ã�<br><br>"
	mailbody=mailbody &"��������"&postname&"����������һ��"&Forum_info(0)&"--"&boardtype&"�ϵ�����<BR><br>"
	mailbody=mailbody &"�����ǣ�"&htmlencode(topic)&"<br><br>"
	mailbody=mailbody &""&htmlencode(content)&"<br><br>"
	mailbody=mailbody &"�����Ե�<a href=http://"&request.servervariables("server_name")&replace(request.servervariables("script_name"),"sendpage.asp","")&"dispbbs.asp?boardid="&boardid&"&id="&Announceid&">"&topic&"</a>��������������<br>"
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
    <th valign=middle colspan=2>�����ʼ�������</th></tr>
    <tr>
    <td  class=tablebody1 valign=middle colspan=2>
    <b>ͨ���ʼ����ͱ����Ӹ��������ѡ�</b>������������������������ȷ���ʼ���ַ��
    <br>��������һЩ�Լ�����Ϣ����������ݿ��ڡ�����������ӵ������ URL ����Բ���д����Ϊ��������ڷ��͵� Email ���Զ���ӵģ�
        </td></tr>
    <tr>
    <td class=tablebody1><b>����������</b></td>
    <td class=tablebody1><input type=text size=40 name="postname"></td>
    </tr><tr>
    <td class=tablebody1><b>�����ѵ����֣�</b></td>
    <td class=tablebody1><input type=text size=40 name="incept"></td>
    </tr><tr>
    <td class=tablebody1><b>�����ѵ� Email��</b></td>
    <td class=tablebody1><input type=text size=40 name="mail"></td>
    </tr><tr>
    <td class=tablebody1><b>��Ϣ���ݣ�</b></td>
    <td class=tablebody1><textarea name="content" cols="55" rows="6">������� '<%=Forum_info(0)%>' ������������ݻ����Ȥ�ģ���ȥ������</textarea></td>
    </tr><tr>
    <td colspan=2 class=tablebody2 align=center><input type=submit value="�� ��" name="Submit"></td></tr></table></form>
<%end sub%>
