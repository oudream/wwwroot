<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!--#include file="inc/ubbcode.asp"-->
<!--#include file="inc/chkinput.asp"-->
<!--#include file="inc/usercp.asp"-->
<%
'=========================================================
' File: messanger.asp
' Version:5.0
' Date: 2002-9-11
' Script Written by satan
'=========================================================
' Copyright (C) 2001,2002 AspSky.Net. All rights reserved.
' Web: http://www.aspsky.net,http://www.dvbbs.net
' Email: info@aspsky.net,eway@aspsky.net
'=========================================================
%>
<html>
<head>
<title><%=Forum_info(0)%>--����Ϣ</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--#include file="inc/Forum_css.asp"-->
</head>
<body <%=Forum_body(11)%>" onkeydown="if(event.keyCode==13 && event.ctrlKey)messager.submit()">
<%
dim msg
dim abgcolor
dim username
if not founduser then
  	errmsg=errmsg+"<br>"+"<li>��û��<a href=login.asp target=_blank>��¼</a>��"
	founderr=true
end if
if Cint(GroupSetting(32))=0 then
	Errmsg=Errmsg+"<br>"+"<li>��û��������Ͷ��ŵ�Ȩ�ޣ���<a href=login.asp>��½</a>����ͬ����Ա��ϵ��"
	founderr=true
end if
stats="���Ź���"
if founderr then
	call dvbbs_error()
else
	select case request("action")
	case "new"
		call sendmsg()
	case "read"
		call read()
	case "outread"
		call read()
	case "delet"
		call delete()
	case "newmsg"
		call newmsg()
	case "send"
		call savemsg()
	case "fw"
		call fw()
	case "edit"
		call edit()
	case "savedit"
		call savedit()
	case "ɾ���ռ�"
		call nav()
		call cp_var()
		call delinbox()
	case "����ռ���"
		call nav()
		call cp_var()
		call AllDelinbox()
	case "ɾ���ݸ�"
		call nav()
		call cp_var()
		call deloutbox()
	case "��ղݸ���"
		call nav()
		call cp_var()
		call AllDeloutbox()
	case "ɾ���ѷ��͵���Ϣ"
		call nav()
		call cp_var()
		call delissend()
	case "����ѷ��͵���Ϣ"
		call nav()
		call cp_var()
		call AllDelissend()
	case "ɾ������"
		call nav()
		call cp_var()
		call delrecycle()
	case "���������"
		call nav()
		call cp_var()
		call AllDelrecycle()
	case else
	  	errmsg=errmsg+"<br>"+"<li>��ָ����ȷ�Ĳ�����"
		founderr=true
	end select
	if founderr then call dvbbs_error()
end if
call activeonline()
call footer()

'������Ϣ
sub sendmsg()
stats="���Ͷ���"
dim sendtime,title,content
if request("id")<>"" and isNumeric(request("id")) then
set rs=server.createobject("adodb.recordset")
sql="select sendtime,title,content from message where incept='"&membername&"' and id="&request("id")
rs.open sql,conn,1,1
if not(rs.eof and rs.bof) then
sendtime=rs("sendtime")
title="RE " & rs("title")
content=rs("content")
end if
rs.close
set rs=nothing
end if
%>
<form action="messanger.asp" method=post name=messager>
<input type=hidden name="action" value="send">
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
          <tr> 
            <th colspan=3>���Ͷ���Ϣ��������������Ϣ��</td>
          </tr>
          <tr> 
            <td class=tablebody1 valign=middle><b>�ռ��ˣ�</b></td>
            <td class=tablebody1 valign=middle>
              <input type=text name="touser" value="<%=request("touser")%>" size=50>
              <SELECT name=font onchange=DoTitle(this.options[this.selectedIndex].value)>
              <OPTION selected value="">ѡ��</OPTION>
<%
set rs=server.createobject("adodb.recordset")
sql="select F_friend from Friend where F_username='"&membername&"' order by F_addtime desc"
rs.open sql,conn,1,1
do while not rs.eof
%>
			  <OPTION value="<%=rs(0)%>"><%=rs(0)%></OPTION> 
<%
rs.movenext
loop
rs.close
set rs=nothing
%>
			  </SELECT>
            </td>
          </tr>
          <tr> 
            <td class=tablebody1 valign=top width=15%><b>���⣺</b></td>
            <td class=tablebody1 valign=middle>
              <input type=text name="title" size=70 maxlength=80 value="<%=title%>">
            </td>
          </tr>
          <tr> 
            <td class=tablebody1 valign=top width=15%><b>���ݣ�</b></td>
            <td  class=tablebody1 valign=middle>
              <textarea cols=70 rows=6 name="message" title="Ctrl+Enter����"><%if request("id")<>"" then%>
====== �� <%=sendtime%> ��������д���� ======
<%=Server.htmlencode(content)%>
=====================================
<%end if%></textarea>
            </td>
          </tr>
          <tr> 
            <td  class=tablebody1 colspan=2>
<b>˵��</b>��<br>
�� ������ʹ��<b>Ctrl+Enter</b>����ݷ��Ͷ���<br>
�� ������Ӣ��״̬�µĶ��Ž��û�������ʵ��Ⱥ�������<b><%=GroupSetting(33)%></b>���û�<br>
�� �������<b>50</b>���ַ����������<b><%=GroupSetting(34)%></b>���ַ�<br>
            </td>
          </tr>
          <tr> 
            <td  class=tablebody2 valign=middle colspan=2 align=center> 
              <input type=Submit value="����" name=Submit>
              &nbsp; 
              <input type=Submit value="����" name=Submit>
              &nbsp; 
              <input type="reset" name="Clear" value="���">
              &nbsp; 
<%if request("reaction")="chatlog" then%>
              <input type=button value="�ر������¼" name="chatlog" onclick="location.href='?action=new&id=<%=request("id")%>&touser=<%=request("touser")%>'">
<%else%>
              <input type=button value="�鿴�����¼" name="chatlog" onclick="location.href='?action=new&id=<%=request("id")%>&touser=<%=request("touser")%>&reaction=chatlog'">
<%end if%>
              &nbsp; 
              <input type="button" name="close" value="�ر�" onclick="window.close()">
            </td>
          </tr>
<%if request("reaction")="chatlog" then%>
          <tr> 
            <th colspan=3>����<%=request("touser")%>�������¼</th>
          </tr>
<%if membername=request("touser") then%>
          <tr> 
            <td class=tablebody1 colspan=3>�Լ����Լ��������¼ûʲô�ÿ��ģ���</td>
          </tr>
<%else%>
<%
	set rs=server.createobject("adodb.recordset")
	sql="select * from message where ((sender='"&trim(membername)&"' and incept='"&replace(request("touser"),"'","")&"') or (sender='"&replace(request("touser"),"'","")&"' and incept='"&membername&"')) and delS=0 order by sendtime desc"
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
%>
          <tr> 
            <td class=tablebody1 colspan=3>��û���κ������¼��</td>
          </tr>
<%
	else
	do while not rs.eof
%>
                <tr>
                    <td class=tablebody2 height=25 colspan=3>
<%if rs("sender")=membername then%>
                    ��<b><%=rs("sendtime")%></b>�������ʹ���Ϣ��<b><%=htmlencode(rs("incept"))%></b>��
<%else%>
		    ��<b><%=rs("sendtime")%></b>��<b><%=htmlencode(rs("sender"))%></b>�������͵���Ϣ��
<%end if%></td>
                </tr>
                <tr>
                    <td  class=tablebody1 valign=top align=left colspan=3>
                    <b>��Ϣ���⣺<%=htmlencode(rs("title"))%></b><hr size=1>
                    <%=dvbcode(rs("content"),4)%>
		    </td>
                </tr>
<%
	rs.movenext
	loop
	end if
	rs.close
	set rs=nothing
%>
<%end if%>
<%end if%>
        </table>
</form>
<%
end sub
'��ȡ��Ϣ
sub read()
stats="�Ķ�����"
if request("id")="" or not isNumeric(request("id")) then
Errmsg=Errmsg+"<br>"+"<li>��ָ����ز�����"
Founderr=true
exit sub
end if
	set rs=server.createobject("adodb.recordset")
	if request("action")="read" then
   	sql="update message set flag=1 where ID="&cstr(request("id"))
	conn.execute(sql)
	end if
	sql="select * from message where (incept='"&membername&"' or sender='"&membername&"') and id="&cstr(request("id"))
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
		errmsg=errmsg+"<br>"+"<li>���ǲ����ܵ����˵������������߸���Ϣ�Ѿ��ռ���ɾ����"
		founderr=true
	end if
	if not founderr then
%>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
            <tr>
                <th colspan=3>��ӭʹ�ö���Ϣ���գ�<%=membername%></th>
            </tr>
            <tr>
                <td class=tablebody1 valign=middle align=center colspan=3><a href="messanger.asp?action=delet&id=<%=rs("id")%>"><img src="<%=Forum_info(7)%>m_delete.gif" border=0 alt="ɾ����Ϣ"></a> &nbsp; <a href="messanger.asp?action=new"><img src="<%=Forum_info(7)%>m_write.gif" border=0 alt="������Ϣ"></a> &nbsp;<a href="messanger.asp?action=new&touser=<%=htmlencode(rs("sender"))%>&id=<%=request("id")%>"><img src="<%=Forum_info(7)%>m_reply.gif" border=0 alt="�ظ���Ϣ"></a>&nbsp;<a href="messanger.asp?action=fw&id=<%=request("id")%>"><img src=<%=Forum_info(7)%>m_fw.gif border=0 alt=ת����Ϣ></a></td>
            </tr>
                <tr>
                    <td class=tablebody2 height=25>
<%if request("action")="outread" then%>
                    ��<b><%=rs("sendtime")%></b>�������ʹ���Ϣ��<b><%=htmlencode(rs("incept"))%></b>��
<%else%>
		    ��<b><%=rs("sendtime")%></b>��<b><%=htmlencode(rs("sender"))%></b>�������͵���Ϣ��
<%end if%></td>
                </tr>
                <tr>
                    <td  class=tablebody1 valign=top align=left>
                    <b>��Ϣ���⣺<%=htmlencode(rs("title"))%></b><hr size=1>
                    <%=dvbcode(rs("content"),4)%>
		    </td>
                </tr>
<%
rs.close
set rs=nothing
	sql="select id,sender from message where incept='"&membername&"' and flag=0 and issend=1 and id>"&cstr(request("id")&" order by sendtime")
	set rs=conn.execute(sql)
	if not (rs.eof and rs.bof) then
%>
                <tr>
                    <td  class=tablebody2 valign=top align=right><a href=messanger.asp?action=read&id=<%=rs(0)%>&sender=<%=rs(1)%>>[��ȡ��һ����Ϣ]</a>
		    </td>
                </tr>
<%
end if
rs.close
set rs=nothing
%>
                </table>
<%end if%>
<%
end sub
'ת����Ϣ
sub fw()
stats="���Ͷ���"
dim title,content,sender
if request("id")<>"" and isNumeric(request("id")) then
set rs=server.createobject("adodb.recordset")
sql="select title,content,sender from message where (incept='"&membername&"' or sender='"&membername&"') and id="&request("id")
rs.open sql,conn,1,1
if rs.eof and rs.bof then
Errmsg=Errmsg+"<br>"+"<li>��ѡ����ز�����"
Founderr=true
exit sub
else
title=rs("title")
content=rs("content")
sender=rs("sender")
end if
rs.close
set rs=nothing
end if
%>
<form action="messanger.asp" method=post name=messager>
        <table cellpadding=3 cellspacing=1 align=center class=tableborder1>
          <tr> 
            <th colspan=2 height=25>
              <input type=hidden name="action" value="send">
              ���Ͷ���Ϣ--����������������Ϣ</th>
          </tr>
          <tr> 
            <td class=tablebody1 valign=middle width=15%><b>�ռ��ˣ�</b></td>
            <td  class=tablebody1 valign=middle>
              <input type=text name="touser" value="<%=request("touser")%>" size=50>
              <SELECT name=font onchange=DoTitle(this.options[this.selectedIndex].value)>
              <OPTION selected value="">ѡ��</OPTION>
<%
set rs=server.createobject("adodb.recordset")
sql="select F_friend from Friend where F_username='"&membername&"' order by F_addtime desc"
rs.open sql,conn,1,1
do while not rs.eof
%>
			  <OPTION value="<%=rs(0)%>"><%=rs(0)%></OPTION> 
<%
rs.movenext
loop
rs.close
set rs=nothing
%>
			  </SELECT>
            </td>
          </tr>
          <tr> 
            <td class=tablebody1 valign=top><b>���⣺</b></td>
            <td class=tablebody1 valign=middle>
              <input type=text name="title" size=70 maxlength=80 value="Fw��<%=title%>">&nbsp;
            </td>
          </tr>
          <tr> 
            <td class=tablebody1 valign=top><b>���ݣ�</b></td>
            <td class=tablebody1 valign=middle>
              <textarea cols=70 rows=6 name="message" title="Ctrl+Enter����">


========== ������ת����Ϣ =========
ԭ�����ˣ�<%=sender%><%=chr(13)&chr(13)%>
<%=Server.htmlencode(content)%>
===================================</textarea>
            </td>
          </tr>
          <tr> 
            <td  class=tablebody1 colspan=2>
<b>˵��</b>��<br>
�� ������ʹ��<b>Ctrl+Enter</b>����ݷ��Ͷ���<br>
�� ������Ӣ��״̬�µĶ��Ž��û�������ʵ��Ⱥ�������<b><%=GroupSetting(33)%></b>���û�<br>
�� �������<b>50</b>���ַ����������<b><%=GroupSetting(34)%></b>���ַ�<br>
            </td>
          </tr>
          <tr> 
            <td class=tablebody2 valign=middle colspan=2 align=center> 
              <input type=Submit value="����" name=Submit>
              &nbsp; 
              <input type=Submit value="����" name=Submit>
              &nbsp; 
              <input type="reset" name="Clear" value="���">
              &nbsp; 
              <input type="button" name="close" value="�ر�" onclick="window.close()">
            </td>
          </tr>
        </table>
</form>
<%
end sub

sub savemsg()
stats="���Ͷ��ųɹ�"
dim incept,title,message,subtype
if request("touser")="" then
	errmsg=errmsg+"<br>"+"<li>��������д���Ͷ����˰ɡ�"
	founderr=true
	exit sub
else
	incept=CheckStr(request("touser"))
	incept=split(incept,",")
end if
if request("title")="" then
	errmsg=errmsg+"<br>"+"<li>����û����д����ѽ��"
	founderr=true
	exit sub
elseif strlength(request("title"))>50 then
	errmsg=errmsg+"<br>"+"<li>�����޶����50���ַ���"
	founderr=true
	exit sub
else
	title=CheckStr(request("title"))
end if
if request("message")="" then
	errmsg=errmsg+"<br>"+"<li>�����Ǳ���Ҫ��д���ޡ�"
	founderr=true
	exit sub
elseif strlength(request("message"))>Cint(GroupSetting(34)) then
	errmsg=errmsg+"<br>"+"<li>�����޶����"&GroupSetting(34)&"���ַ���"
	founderr=true
	exit sub
else
	message=CheckStr(request("message"))
end if

for i=0 to ubound(incept)
sql="select username from [user] where username='"&replace(incept(i),"'","")&"'"
set rs=conn.execute(sql)
if rs.eof and rs.bof then
	errmsg=errmsg+"<br>"+"<li>��̳û������û���������ķ��Ͷ���д�����"
	founderr=true
	exit sub
end if
set rs=nothing

if request("Submit")="����" then
	sql="insert into message (incept,sender,title,content,sendtime,flag,issend) values ('"&incept(i)&"','"&membername&"','"&title&"','"&message&"',Now(),0,1)"
	subtype="�ѷ�����Ϣ"
elseif request("Submit")="����" then
	sql="insert into message (incept,sender,title,content,sendtime,flag,issend) values ('"&incept(i)&"','"&membername&"','"&title&"','"&message&"',Now(),0,0)"
	subtype="������"
else
	sql="insert into message (incept,sender,title,content,sendtime,flag,issend) values ('"&incept(i)&"','"&membername&"','"&title&"','"&message&"',Now(),0,1)"
	subtype="�ѷ�����Ϣ"
end if
conn.execute(sql)
if i>Cint(GroupSetting(33))-1 then
	errmsg=errmsg+"<br>"+"<li>���ֻ�ܷ��͸�"&GroupSetting(33)&"���û�����������"&GroupSetting(33)&"λ�Ժ�������·���"
	founderr=true
	exit sub
	exit for
end if
next
sucmsg=sucmsg+"<br>"+"<li><b>��ϲ�������Ͷ���Ϣ�ɹ���</b><br>���͵���Ϣͬʱ����������"&subtype&"�С�"
call dvbbs_suc()
end sub

'������Ϣ
sub edit()
stats="�޸Ķ���"
dim incept,title,content,id
if request("id")<>"" and isNumeric(request("id")) then
set rs=server.createobject("adodb.recordset")
sql="select id,incept,title,content from message where sender='"&membername&"' and issend=0 and id="&request("id")
rs.open sql,conn,1,1
if not(rs.eof and rs.bof) then
incept=rs("incept")
title=rs("title")
content=rs("content")
id=rs("id")
else
Errmsg=Errmsg+"<br>"+"<li>û���ҵ���Ҫ�༭����Ϣ��"
Founderr=true
exit sub
end if
rs.close
set rs=nothing
else
Errmsg=Errmsg+"<br>"+"<li>��ָ����ز�����"
Founderr=true
exit sub
end if
%>
<form action="messanger.asp" method=post name=messager>
        <table cellpadding=3 cellspacing=1 align=center class=tableborder1>
          <tr> 
            <th colspan=2 height=25> 
              <input type=hidden name="action" value="savedit"> 
              <input type=hidden name="id" value="<%=id%>">
              ���Ͷ���Ϣ--����������������Ϣ</th>
          </tr>
          <tr> 
            <td  class=tablebody1 valign=middle width=70><b>�ռ��ˣ�</b></td>
            <td  class=tablebody1 valign=middle>
              <input type=text name="touser" value="<%=incept%>" size=70>
            </td>
          </tr>
          <tr> 
            <td class=tablebody1 valign=top><b>���⣺</b></td>
            <td  class=tablebody1 valign=middle>
              <input type=text name="title" size=70 maxlength=80 value="<%=title%>">
            </td>
          </tr>
          <tr> 
            <td  class=tablebody1 valign=top><b>���ݣ�</b></td>
            <td  class=tablebody1 valign=middle>
              <textarea cols=70 rows=8 name="message" title=""><%=Server.htmlencode(content)%></textarea>
            </td>
          </tr>
          <tr> 
            <td  class=tablebody1 colspan=2>
<b>˵��</b>��<br>
�� ������ʹ��<b>Ctrl+Enter</b>����ݷ��Ͷ���<br>
�� �������<b>50</b>���ַ����������<b><%=GroupSetting(34)%></b>���ַ�<br>
            </td>
          </tr>
          <tr> 
            <td  class=tablebody2 valign=middle colspan=2 align=center> 
              <input type=Submit value="����" name=Submit>
              &nbsp; 
              <input type=Submit value="����" name=Submit>
              &nbsp; 
              <input type="reset" name="Clear" value="���">
              &nbsp; 
              <input type="button" name="close" value="�ر�" onclick="window.close()">
            </td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
</form>
<%
end sub

sub savedit()
dim incept,title,message,subtype
if request("id")="" or not isNumeric(request("id")) then
	Errmsg=Errmsg+"<br>"+"<li>��ָ����ز�����"
	Founderr=true
	exit sub
end if
if request("touser")="" then
	errmsg=errmsg+"<br>"+"<li>��������д���Ͷ����˰ɡ�"
	founderr=true
	exit sub
else
	incept=checkStr(request("touser"))
end if
if request("title")="" then
	errmsg=errmsg+"<br>"+"<li>����û����д����ѽ��"
	founderr=true
	exit sub
else
	title=checkStr(request("title"))
end if
if request("message")="" then
	errmsg=errmsg+"<br>"+"<li>�����Ǳ���Ҫ��д���ޡ�"
	founderr=true
	exit sub
else
	message=checkStr(request("message"))
end if

sql="select username from [user] where username='"&incept&"'"
set rs=conn.execute(sql)
if rs.eof and rs.bof then
	errmsg=errmsg+"<br>"+"<li>��̳û������û���������ķ��Ͷ���д�����"
	founderr=true
	exit sub
end if
set rs=nothing

if request("Submit")="����" then
	sql="update message set incept='"&incept&"',sender='"&membername&"',title='"&title&"',content='"&message&"',sendtime=Now(),flag=0,issend=1 where id="&request("id")
	subtype="�ѷ�����Ϣ"
	else
	sql="update message set incept='"&incept&"',sender='"&membername&"',title='"&title&"',content='"&message&"',sendtime=Now(),flag=0,issend=0 where id="&request("id")
	subtype="������"
end if
conn.execute(sql)

sucmsg=sucmsg+"<br>"+"<li><b>��ϲ�������Ͷ���Ϣ�ɹ���</b><br>���͵���Ϣͬʱ����������"&subtype&"�С�"
call dvbbs_suc()
end sub

'�ռ��߼�ɾ�������ڻ���վ������ֶ�delR������������������ɾ��
sub delinbox()
dim delid
delid=replace(request("id"),"'","")
if delid="" or isnull(delid) then
Errmsg=Errmsg+"<li>"+"��ѡ����ز�����"
Founderr=true
else
	conn.execute("update message set delR=1 where incept='"&trim(membername)&"' and id in ("&delid&")")
	sucmsg=sucmsg+"<br>"+"<li>ɾ������Ϣ�ɹ���ɾ������Ϣ��ת�Ƶ����Ļ���վ��"
	call dvbbs_suc()
end if
end sub
sub AllDelinbox()
	conn.execute("update message set delR=1 where incept='"&trim(membername)&"' and delR=0")
	sucmsg=sucmsg+"<br>"+"<li>ɾ������Ϣ�ɹ���ɾ������Ϣ��ת�Ƶ����Ļ���վ��"
	call dvbbs_suc()
end sub

'�����߼�ɾ�������ڻ���վ������ֶ�delS������������������ɾ��
sub deloutbox()
dim delid
delid=replace(request("id"),"'","")
if delid="" or isnull(delid) then
Errmsg=Errmsg+"<li>"+"��ѡ����ز�����"
Founderr=true
else
	conn.execute("update message set delS=1 where sender='"&trim(membername)&"' and issend=0 and id in ("&delid&")")
	sucmsg=sucmsg+"<br>"+"<li>ɾ������Ϣ�ɹ���ɾ������Ϣ��ת�Ƶ����Ļ���վ��"
	call dvbbs_suc()
end if
end sub
sub AllDeloutbox()
	conn.execute("update message set delS=1 where sender='"&trim(membername)&"' and delS=0 and issend=0")
	sucmsg=sucmsg+"<br>"+"<li>ɾ������Ϣ�ɹ���ɾ������Ϣ��ת�Ƶ����Ļ���վ��"
	call dvbbs_suc()
end sub

'�ѷ����߼�ɾ�������ڻ���վ������ֶ�delS������������������ɾ��
'delS��0δ������1������ɾ����2�����ߴӻ���վɾ��
sub delissend()
dim delid
delid=replace(request("id"),"'","")
if delid="" or isnull(delid) then
Errmsg=Errmsg+"<li>"+"��ѡ����ز�����"
Founderr=true
else
	conn.execute("update message set delS=1 where sender='"&trim(membername)&"' and issend=1 and id in ("&delid&")")
	sucmsg=sucmsg+"<br>"+"<li>ɾ������Ϣ�ɹ���ɾ������Ϣ��ת�Ƶ����Ļ���վ��"
	call dvbbs_suc()
end if
end sub
sub AllDelissend()
	conn.execute("update message set delS=1 where sender='"&trim(membername)&"' and delS=0 and issend=1")
	sucmsg=sucmsg+"<br>"+"<li>ɾ������Ϣ�ɹ���ɾ������Ϣ��ת�Ƶ����Ļ���վ��"
	call dvbbs_suc()
end sub

'�û�����ȫɾ���յ���Ϣ���߼�ɾ����������Ϣ���߼�ɾ����������Ϣ��������ֶ�delS����Ϊ2
sub delrecycle()
dim delid
delid=replace(request("id"),"'","")
'response.write delid
if delid="" or isnull(delid) then
Errmsg=Errmsg+"<li>"+"��ѡ����ز�����"
Founderr=true
exit sub
else
	conn.execute("delete from message where incept='"&membername&"' and delR=1 and id in ("&delid&")")
	conn.execute("update message set delS=2 where sender='"&trim(membername)&"' and delS=1 and id in ("&delid&")")
	sucmsg=sucmsg+"<br>"+"<li>ɾ������Ϣ�ɹ���ɾ������Ϣ�����ɻָ���"
	call dvbbs_suc()
end if
end sub
sub AllDelrecycle()
	conn.execute("delete from message where incept='"&membername&"' and delR=1")	
	conn.execute("update message set delS=2 where sender='"&trim(membername)&"' and delS=1")
	sucmsg=sucmsg+"<br>"+"<li>ɾ������Ϣ�ɹ���ɾ������Ϣ�����ɻָ���"
	call dvbbs_suc()
end sub

sub delete()
dim delid
delid=checkstr(request("id"))
if not isNumeric(request("id")) or delid="" or isnull(delid) then
Errmsg=Errmsg+"<li>"+"��ѡ����ز�����"
Founderr=true
else
	conn.execute("update message set delR=1 where incept='"&trim(membername)&"' and id="&delid)
	conn.execute("update message set delS=1 where sender='"&trim(membername)&"' and id="&delid)
	sucmsg=sucmsg+"<br>"+"<li>ɾ������Ϣ�ɹ���ɾ������Ϣ���������Ļ���վ�ڡ�"
	call dvbbs_suc()
end if
end sub
%>
<script language="javascript"> 
function DoTitle(addTitle) {  
 var revisedTitle;  
 var currentTitle = document.messager.touser.value; 

 if(currentTitle=="") revisedTitle = addTitle; 
 else { 
  var arr = currentTitle.split(","); 
  for (var i=0; i < arr.length; i++) { 
   if( addTitle.indexOf(arr[i]) >=0 && arr[i].length==addTitle.length ) return; 
  } 
  revisedTitle = currentTitle+","+addTitle; 
 } 

 document.messager.touser.value=revisedTitle;  
 document.messager.touser.focus(); 
 return; 
} 
</script>