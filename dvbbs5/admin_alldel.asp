<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<title><%=Forum_info(0)%>--����ҳ��</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta NAME=GENERATOR Content="Microsoft FrontPage 3.0" CHARSET=GB2312>
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0" marginheight="0" marginwidth="0">
<%
	Server.ScriptTimeout=9999999
	if not master or session("flag")="" then
		Errmsg=Errmsg+"<br>"+"<li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��½</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
		call dvbbs_error()
	else
		dim body
		call main()
		conn.close
		set conn=nothing
	end if

	sub main()
	select case request("action")
	case "alldel"
		call alldel()
	case "userdel"
		call del()
	case "alldelTopic"
		call alldelTopic()
	case "deluser"
		call deluser()
	case "moveinfo"
		call moveinfo()
	case "MoveUserTopic"
		call moveusertopic()
	case "MoveDateTopic"
		call movedatetopic()
	case else
%>
<table cellpadding=3 cellspacing=1 border=0 width=95% align=center>
	<tr>
    <td width="100%" valign=top>
<B>ע��</B>��������뻹ԭ���ӣ��뵽��̳����վ��
            <br>���������������ɾ����̳���ӡ������ȷ��������������ϸ������������Ϣ��
</td>
</tr>
</table>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<form action="admin_alldel.asp?action=alldel" method="post">
            <tr>
            <th valign=middle colspan=2 height=23 align=left>ɾ��ָ������������</b>(�����ܲ��۳��û��������ͻ���)</th></tr>
            <tr>
            <td valign=middle width=40% class=forumrow>ɾ��������ǰ������(��д����)</td><td class=forumrow><input name="TimeLimited" value=0 size=30>&nbsp;<input type=submit name="submit" value="�� ��"></td></tr>
            <tr>
            <td valign=middle width=40%  class=forumrow>��̳����</td><td class=forumrow>
<select name="delboardid" size=1>
<%
set rs=conn.execute("select boardid,boardtype from board order by boardid")
if rs.eof and rs.bof then
response.write "<option value=0>û����̳</option>"
else
response.write "<option value=all>ȫ����̳</option>"
do while not rs.eof
response.write "<option value="&rs(0)&">"&rs(1)&"</option>"
rs.movenext
loop
end if
rs.close
set rs=nothing
%>
</select>
			</td></tr>
</form>
<form action="admin_alldel.asp?action=alldelTopic" method="post">
            <tr>
            <th valign=middle colspan=2 height=23 align=left>ɾ��ָ��������û�лظ�������(�����ܲ��۳��û��������ͻ���)</th></tr>
            <tr>
            <td valign=middle width=40%  class=forumrow>ɾ��������ǰ������(��д����)</td><td class=forumrow><input name="TimeLimited" value=0 size=30>&nbsp;<input type=submit name="submit" value="�� ��"></td></tr>
            <tr>
            <td valign=middle width=40%  class=forumrow>��̳����</td><td class=forumrow>
<select name="delboardid" size=1>
<%
set rs=conn.execute("select boardid,boardtype from board order by boardid")
if rs.eof and rs.bof then
response.write "<option value=0>û����̳</option>"
else
response.write "<option value=all>ȫ����̳</option>"
do while not rs.eof
response.write "<option value="&rs(0)&">"&rs(1)&"</option>"
rs.movenext
loop
end if
rs.close
set rs=nothing
%>
</select>
			</td></tr>
</form>
<form action="admin_alldel.asp?action=userdel" method="post">
            <tr>
            <th valign=middle colspan=2 height=23 align=left>ɾ��ĳ�û�����������</td></tr>
            <tr>
            <td valign=middle width=40%  class=forumrow>�������û���</td><td class=forumrow><input type=text name="username" size=30>&nbsp;<input type=submit name="submit" value="�� ��"></td></tr>
            <tr>
            <td valign=middle width=40%  class=forumrow>��̳����</td><td class=forumrow>
<select name="delboardid" size=1>
<%
set rs=conn.execute("select boardid,boardtype from board order by boardid")
if rs.eof and rs.bof then
response.write "<option value=0>û����̳</option>"
else
response.write "<option value=all>ȫ����̳</option>"
do while not rs.eof
response.write "<option value="&rs(0)&">"&rs(1)&"</option>"
rs.movenext
loop
end if
rs.close
set rs=nothing
%>
</select>
			</td></tr>
</form>
<!--
<form action="admin_alldel.asp?action=delUser" method="post">
            <tr>
            <td bgcolor="<%=Forum_body(4)%>" valign=middle><font color="<%=Forum_body(7)%>">ɾ��ָ��������û�е�½���û�</font></td>
            <td bgcolor="<%=Forum_body(4)%>" valign=middle>
<select name=TimeLimited size=1> 
<option value=all>ɾ�����е�
<option value=1>ɾ��һ��ǰ��
<option value=2>ɾ������ǰ��
<option value=7>ɾ��һ����ǰ��
<option value=15>ɾ�������ǰ��
<option value=30>ɾ��һ����ǰ��
<option value=60>ɾ��������ǰ��
<option value=180>ɾ������ǰ��
</select>
</select><input type=submit name="submit" value="�� ��"></td></tr></form>
-->
</table>
<%end select%>
<%if founderr then call dvbbs_error()%>
<%
	end sub

	sub moveinfo()
%>
<table cellpadding=3 cellspacing=1 border=0 width=95% align=center>
	<tr>
    <td width="100%" valign=top>
<B>ע��</B>������ֻ���ƶ����ӣ������ǿ�������ɾ����
            <br>���������ɾ��ԭ��̳���ӣ����ƶ�����ָ������̳�С������ȷ��������������ϸ������������Ϣ��
</td>
</tr>
</table>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<form action="admin_alldel.asp?action=MoveDateTopic" method="post">
            <tr>
            <th valign=middle colspan=2 height=23 align=left>�������ƶ�</th></tr>
            <tr>
            <td valign=middle width=40% class=forumrow>�ƶ�������ǰ������(��д����)</td><td class=forumrow><input name="TimeLimited" value=0 size=30>&nbsp;<input type=submit name="submit" value="�� ��"></td></tr>
            <tr>
            <td valign=middle width=40%  class=forumrow>ԭ��̳</td><td class=forumrow>
<select name="outboardid" size=1>
<%
set rs=conn.execute("select boardid,boardtype from board order by boardid")
if rs.eof and rs.bof then
response.write "<option value=0>û����̳</option>"
else
do while not rs.eof
response.write "<option value="&rs(0)&">"&rs(1)&"</option>"
rs.movenext
loop
end if
rs.close
set rs=nothing
%>
</select>
			</td></tr>
            <tr>
            <td valign=middle width=40%  class=forumrow>Ŀ����̳</td><td class=forumrow>
<select name="inboardid" size=1>
<%
set rs=conn.execute("select boardid,boardtype from board order by boardid")
if rs.eof and rs.bof then
response.write "<option value=0>û����̳</option>"
else
do while not rs.eof
response.write "<option value="&rs(0)&">"&rs(1)&"</option>"
rs.movenext
loop
end if
rs.close
set rs=nothing
%>
</select>
			</td></tr>
</form>
<form action="admin_alldel.asp?action=MoveUserTopic" method="post">
            <tr>
            <th valign=middle colspan=2 height=23 align=left>���û��ƶ�</th></tr>
            <tr>
            <td valign=middle width=40%  class=forumrow>����д�û���</td><td class=forumrow><input name="username" size=30>&nbsp;<input type=submit name="submit" value="�� ��"></td></tr>
            <tr>
            <td valign=middle width=40%  class=forumrow>ԭ��̳</td><td class=forumrow>
<select name="outboardid" size=1>
<%
set rs=conn.execute("select boardid,boardtype from board order by boardid")
if rs.eof and rs.bof then
response.write "<option value=0>û����̳</option>"
else
do while not rs.eof
response.write "<option value="&rs(0)&">"&rs(1)&"</option>"
rs.movenext
loop
end if
rs.close
set rs=nothing
%>
</select>
			</td></tr>
            <tr>
            <td valign=middle width=40%  class=forumrow>Ŀ����̳</td><td class=forumrow>
<select name="inboardid" size=1>
<%
set rs=conn.execute("select boardid,boardtype from board order by boardid")
if rs.eof and rs.bof then
response.write "<option value=0>û����̳</option>"
else
do while not rs.eof
response.write "<option value="&rs(0)&">"&rs(1)&"</option>"
rs.movenext
loop
end if
rs.close
set rs=nothing
%>
</select>
			</td></tr>
</form>
</table>
<%
	end sub
	sub del()
		dim titlenum,delboardid
		if request("delboardid")="0" then
			founderr=true
			errmsg=errmsg+"<br>"+"<li>�Ƿ��İ��������"
			exit sub
		elseif request("delboardid")="all" then
			delboardid=""
		else
			delboardid=" boardid="&request("delboardid")&" and "
		end if
		if request("username")="" then
			founderr=true
			errmsg=errmsg+"<br>"+"<li>�����뱻����ɾ���û�����"
			exit sub
		end if
		titlenum=0
		for i=0 to ubound(allposttable)
		set rs=conn.execute("Select Count(announceID) from "&allposttable(i)&" where "&delboardid&"   username='"&replace(request("username"),"'","")&"'") 
   		titlenum=titlenum+rs(0)

		sql="update "&allposttable(i)&" set locktopic=2 where "&delboardid&"  username='"&replace(request("username"),"'","")&"'"
		conn.Execute(sql)
		next
		set rs=conn.execute("select topicid,posttable from topic where "&delboardid&"  postusername='"&replace(request("username"),"'","")&"'")
		do while not rs.eof
		conn.execute("update "&rs(1)&" set locktopic=2 where rootid="&rs(0))
		rs.movenext
		loop
		set rs=nothing
		conn.execute("update topic set locktopic=2 where "&delboardid&"  Postusername='"&replace(request("username"),"'","")&"'")
		if isnull(titlenum) then titlenum=0
		sql="update [user] set article=article-"&titlenum&",userWealth=userWealth-"&titlenum*Forum_user(3)&",userEP=userEP-"&titlenum*Forum_user(8)&",userCP=userCP-"&titlenum*Forum_user(13)&" where username='"&replace(request("username"),"'","")&"'"
		conn.Execute(sql)
	response.write "ɾ���ɹ������Ҫ��ȫɾ�������뵽��̳����վ<BR>��������������̳�����и���һ����̳���ݣ�����<a href=admin_alldel.asp>����</a>"
	end sub

	sub alldel()
	Dim TimeLimited,delboardid
	if request("delboardid")="0" then
		founderr=true
		errmsg=errmsg+"<br>"+"<li>�Ƿ��İ��������"
		exit sub
	elseif request("delboardid")="all" then
		delboardid=""
	else
		delboardid=" and boardid="&request("delboardid")&" "
	end if
	TimeLimited=request.form("TimeLimited")
	if not isnumeric(TimeLimited) then
		founderr=true
		errmsg=errmsg+"<br>"+"<li>�Ƿ��Ĳ�����"
		exit sub
	else
	for i=0 to ubound(allposttable)
	conn.execute("update "&allposttable(i)&" set LockTopic=2 where datediff('d',DateAndTime,Now())>"&TimeLimited&" "&delboardid&"")
	next
	conn.execute("update topic set LockTopic=2 where datediff('d',DateAndTime,Now())>"&TimeLimited&" "&delboardid&"")
	end if
	response.write "ɾ���ɹ������Ҫ��ȫɾ�������뵽��̳����վ<BR>��������������̳�����и���һ����̳���ݣ�����<a href=admin_alldel.asp>����</a>"
	end sub

	sub alldelTopic()
	Dim TimeLimited,delboardid
	if request("delboardid")="0" then
		founderr=true
		errmsg=errmsg+"<br>"+"<li>�Ƿ��İ��������"
		exit sub
	elseif request("delboardid")="all" then
		delboardid=""
	else
		delboardid=" boardid="&request("delboardid")&" and "
	end if
	TimeLimited=request.form("TimeLimited")
	if not isnumeric(TimeLimited) then
		founderr=true
		errmsg=errmsg+"<br>"+"<li>�Ƿ��Ĳ�����"
		exit sub
	else
	conn.execute("update topic set LockTopic=2 where "&delboardid&"   datediff('d',DateAndTime,Now())>"&TimeLimited&" and Child=0")
	set rs=conn.execute("select Topicid,PostTable from topic where "&delboardid&"   datediff('d',DateAndTime,Now())>"&TimeLimited&" and Child=0")
	do while not rs.eof
		conn.execute("update "&rs(1)&" set locktopic=2 where rootid="&rs(0))
	rs.movenext
	loop
	set rs=nothing
	end if
	response.write "ɾ���ɹ������Ҫ��ȫɾ�������뵽��̳����վ<BR>��������������̳�����и���һ����̳���ݣ�����<a href=admin_alldel.asp>����</a>"
	end sub

	sub delUser()
	Dim TimeLimited
	TimeLimited=request.form("TimeLimited")
	if TimeLimited="all" then
	conn.execute("delete from [user]")
	else
	conn.execute("delete from [user] where datediff('d',LastLogin,Now())>"&TimeLimited&"")
	end if
	response.write "ɾ���ɹ������Ҫ��ȫɾ�������뵽��̳����վ<BR>��������������̳�����и���һ����̳���ݣ�����<a href=admin_alldel.asp>����</a>"
	end sub

	sub MoveUserTopic()
	if not isnumeric(request("inboardid")) then
	response.write "����İ��������"
	exit sub
	end if
	if not isnumeric(request("outboardid")) then
	response.write "����İ��������"
	exit sub
	end if
	if request("username")="" then
	response.write "����д�û�����"
	exit sub
	end if
	if Cint(request("outboardid"))=Cint(request("inboardid")) then
	response.write "��������ͬ��������ƶ�������"
	exit sub
	end if
	for i=0 to ubound(allposttable)
	conn.execute("update "&allposttable(i)&" set boardid="&request("inboardid")&" where Boardid="&request("outboardid")&" and username='"&request("username")&"'")
	next
	set rs=conn.execute("select topicid,posttable from topic where Boardid="&request("outboardid")&" and Postusername='"&request("username")&"'")
	do while not rs.eof
		conn.execute("update "&rs(1)&" set boardid="&request("inboardid")&" where rootid="&rs(0))
	rs.movenext
	loop
	set rs=nothing
	conn.execute("update topic set boardid="&request("inboardid")&" where Boardid="&request("outboardid")&" and Postusername='"&request("username")&"'")
	conn.execute("update besttopic set boardid="&request("inboardid")&" where Boardid="&request("outboardid")&" and Postusername='"&request("username")&"'")
	response.write "�ƶ��ɹ���"
	end sub

	sub MoveDateTopic()
	if not isnumeric(request("TimeLimited")) then
	response.write "��������ڲ�����"
	exit sub
	end if
	if not isnumeric(request("inboardid")) then
	response.write "����İ��������"
	exit sub
	end if
	if not isnumeric(request("outboardid")) then
	response.write "����İ��������"
	exit sub
	end if
	if Cint(request("outboardid"))=Cint(request("inboardid")) then
	response.write "��������ͬ��������ƶ�������"
	exit sub
	end if
	for i=0 to ubound(allposttable)
	conn.execute("update "&allposttable(i)&" set boardid="&request("inboardid")&" where Boardid="&request("outboardid")&" and datediff('d',DateAndTime,Now())>"&request.Form("TimeLimited")&"")
	next
	conn.execute("update topic set boardid="&request("inboardid")&" where Boardid="&request("outboardid")&" and datediff('d',DateAndTime,Now())>"&request.Form("TimeLimited")&"")
	conn.execute("update besttopic set boardid="&request("inboardid")&" where Boardid="&request("outboardid")&" and datediff('d',DateAndTime,Now())>"&request.Form("TimeLimited")&"")
	response.write "�ƶ��ɹ���"
	end sub
%>