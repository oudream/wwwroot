<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!--#include file="md5.asp"-->
<title><%=Forum_info(0)%>--����ҳ��</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta NAME=GENERATOR Content="Microsoft FrontPage 3.0" CHARSET=GB2312>
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0" marginheight="0" marginwidth="0" bgcolor="#DDEEFF">
<%
	if not master or session("flag")="" then
		Errmsg=Errmsg+"<br>"+"<li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��½</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
		call dvbbs_error()
		response.end
	end if
dim trs
dim userinfo
dim usertitle
%>
<table cellpadding="2" cellspacing="1" border="0" width="95%" class="tableBorder" align=center>
<tr>
<th align=left colspan=6 height=23>�û�����</th>
</tr>
<tr>
<td width=20% class=forumrow>ע������</td>
<td width=80% class=forumrow colspan=5>�ٵ�ɾ����ť��ɾ����ѡ�����û����˲����ǲ�����ģ��������������ƶ��û�����Ӧ���飻�۵��û���������Ӧ�����ϲ������ܵ��û�����½IP�ɽ�������IP�������ݵ��û�Email�������û�����Email</td>
</tr>
<form action="?action=userSearch" method=post>
<tr>
<td width=20% class=forumrow>��������</td>
<td width=80% class=forumrow colspan=5>
<select size=1 name="userSearch" onchange="javascript:submit()">
<option value="0">��ѡ���ѯ����</option>
<option value="1" <%if request("userSearch")=1 then%>selected<%end if%>>�г������û�</option>
<option value="2" <%if request("userSearch")=2 then%>selected<%end if%>>�������TOP100</option>
<option value="3" <%if request("userSearch")=3 then%>selected<%end if%>>�������ٵ�100���û�</option>
<option value="4" <%if request("userSearch")=4 then%>selected<%end if%>>���24Сʱ�ڵ�½���û�</option>
<option value="5" <%if request("userSearch")=5 then%>selected<%end if%>>���24Сʱ��ע����û�</option>
<option value="6" <%if request("userSearch")=6 then%>selected<%end if%>>�ȴ�����Ա��֤���û�</option>
<option value="7" <%if request("userSearch")=7 then%>selected<%end if%>>�ȴ��ʼ���֤�Ļ�Ա</option>
<option value="8" <%if request("userSearch")=8 then%>selected<%end if%>>���а����������û�</option>
</select>
</td>
</tr>
</form>
<%if request("action")="" or request("userSearch")="0" then%>
<form action="?action=userSearch" method=post>
<tr>
<th align=left colspan=6 height=23>�߼���ѯ</th>
</tr>
<tr>
<td width=20% class=forumrow>ע������</td>
<td width=80% class=forumrow colspan=5>�ڼ�¼�ܶ���������������Խ���ѯԽ�����뾡�����ٲ�ѯ�����������ʾ��¼��Ҳ����ѡ�����</td>
</tr>
<tr>
<td width=20% class=forumrow>�����ʾ��¼��</td>
<td width=80% class=forumrow colspan=5><input size=45 name="searchMax" type=text value=100></td>
</tr>
<tr>
<td width=20% class=forumrow>�û���</td>
<td width=80% class=forumrow colspan=5><input size=45 name="username" type=text>&nbsp;<input type=checkbox name="usernamechk" value="yes" checked>�û�������ƥ��</td>
</tr>
<tr>
<td width=20% class=forumrow>�û���</td>
<td width=80% class=forumrow colspan=5>
<select size=1 name="usergroups">
<option value=0>����</option>
<%
set rs=conn.execute("select usergroupid,title from usergroups order by usergroupid")
do while not rs.eof
response.write "<option value="&rs(0)&">"&rs(1)&"</option>"
rs.movenext
loop
rs.close
set rs=nothing
%>
</select>
</td>
</tr>
<tr>
<td width=20% class=forumrow>�û��ȼ�</td>
<td width=80% class=forumrow colspan=5>
<select size=1 name="userclass">
<option value=0>����</option>
<%
set rs=conn.execute("select usertitle from usertitle order by usertitleid")
do while not rs.eof
response.write "<option value="&rs(0)&">"&rs(0)&"</option>"
rs.movenext
loop
rs.close
set rs=nothing
%>
</select>
</td>
</tr>
<tr>
<td width=20% class=forumrow>Email����</td>
<td width=80% class=forumrow colspan=5><input size=45 name="userEmail" type=text></td>
</tr>
<tr>
<td width=20% class=forumrow>��ҳ����</td>
<td width=80% class=forumrow colspan=5><input size=45 name="homepage" type=text></td>
</tr>
<tr>
<td width=20% class=forumrow>QQ����</td>
<td width=80% class=forumrow colspan=5><input size=45 name="oicq" type=text></td>
</tr>
<tr>
<td width=20% class=forumrow>ICQ����</td>
<td width=80% class=forumrow colspan=5><input size=45 name="icq" type=text></td>
</tr>
<tr>
<td width=20% class=forumrow>MSN����</td>
<td width=80% class=forumrow colspan=5><input size=45 name="msn" type=text></td>
</tr>
<tr>
<td width=20% class=forumrow>ͷ�ΰ���</td>
<td width=80% class=forumrow colspan=5><input size=45 name="usertitle" type=text></td>
</tr>
<tr>
<td width=20% class=forumrow>ǩ������</td>
<td width=80% class=forumrow colspan=5><input size=45 name="sign" type=text></td>
</tr>
<tr>
<td width=100% class=forumrow align=center colspan=6><input name="submit" type=submit value="   ��  ��   "></td>
</tr>
<input type=hidden value="9" name="userSearch">
</form>
<%
elseif request("action")="userSearch" then
%>
<tr>
<th colspan=6 align=left height=23>�������</th>
</tr>
<%
	dim currentpage,page_count,Pcount
	dim totalrec,endpage
	currentPage=request("page")
	if currentpage="" or not isInteger(currentpage) then
		currentpage=1
	else
		currentpage=clng(currentpage)
		if err then
			currentpage=1
			err.clear
		end if
	end if
	Set rs= Server.CreateObject("ADODB.Recordset")
	select case request("userSearch")
	case 1
		sql="select u.userid,u.username,u.useremail,u.LastLogin,u.UserLastIP,u.article,u.UserGroupID from [user] u inner join UserGroups G on u.usergroupid=g.usergroupid order by u.addDate desc"
	case 2
		sql="select top 100  u.userid,u.username,u.useremail,u.LastLogin,u.UserLastIP,u.article,u.UserGroupID from [user] u inner join UserGroups G on u.usergroupid=g.usergroupid order by u.article desc"
	case 3
		sql="select top 100 u.userid,u.username,u.useremail,u.LastLogin,u.UserLastIP,u.article,u.UserGroupID from [user] u inner join UserGroups G on u.usergroupid=g.usergroupid order by u.article"
	case 4
		sql="select u.userid,u.username,u.useremail,u.LastLogin,u.UserLastIP,u.article,u.UserGroupID from [user] u inner join UserGroups G on u.usergroupid=g.usergroupid where datediff('h',u.LastLogin,Now())<25 order by u.lastlogin desc"
	case 5
		sql="select u.userid,u.username,u.useremail,u.LastLogin,u.UserLastIP,u.article,u.UserGroupID from [user] u inner join UserGroups G on u.usergroupid=g.usergroupid where datediff('h',u.AddDate,Now())<25 order by u.addDate desc"
	case 6
		sql="select u.userid,u.username,u.useremail,u.LastLogin,u.UserLastIP,u.article,u.UserGroupID from [user] u inner join UserGroups G on u.usergroupid=g.usergroupid where u.usergroupid=5 order by u.addDate desc"
	case 7
		sql="select u.userid,u.username,u.useremail,u.LastLogin,u.UserLastIP,u.article,u.UserGroupID from [user] u inner join UserGroups G on u.usergroupid=g.usergroupid where u.usergroupid=6 order by u.addDate desc"
	case 8
		sql="select u.userid,u.username,u.useremail,u.LastLogin,u.UserLastIP,u.article,u.UserGroupID from [user] u inner join UserGroups G on u.usergroupid=g.usergroupid where u.usergroupid<4 order by u.usergroupid"
	case 9
		sqlstr=""
		if request("username")<>"" then
			if request("usernamechk")="yes" then
			sqlstr=" u.username='"&request("username")&"'"
			else
			sqlstr=" u.username like '%"&request("username")&"%'"
			end if
		end if
		if cint(request("usergroups"))>0 then
			if sqlstr="" then
			sqlstr=" u.usergroupid="&request("usergroups")&""
			else
			sqlstr=sqlstr & " and u.usergroupid="&request("usergroupid")&""
			end if
		end if
		if request("userclass")<>"0" then
			if sqlstr="" then
			sqlstr=" u.userclass='"&request("userclass")&"'"
			else
			sqlstr=sqlstr & " and u.userclass='"&request("userclass")&"'"
			end if
		end if
		if request("useremail")<>"" then
			if sqlstr="" then
			sqlstr=" u.useremail like '%"&request("useremail")&"%'"
			else
			sqlstr=sqlstr & " and u.useremail like '%"&request("useremail")&"%'"
			end if
		end if
		if request("homepage")<>"" then
			if sqlstr="" then
			sqlstr=" u.homepage like '%"&request("homepage")&"%'"
			else
			sqlstr=sqlstr & " and u.homepage like '%"&request("homepage")&"%'"
			end if
		end if
		if request("oicq")<>"" then
			if sqlstr="" then
			sqlstr=" u.oicq like '%"&request("oicq")&"%'"
			else
			sqlstr=sqlstr & " and u.oicq like '%"&request("oicq")&"%'"
			end if
		end if
		if request("icq")<>"" then
			if sqlstr="" then
			sqlstr=" u.icq like '%"&request("icq")&"%'"
			else
			sqlstr=sqlstr & " and u.icq like '%"&request("icq")&"%'"
			end if
		end if
		if request("msn")<>"" then
			if sqlstr="" then
			sqlstr=" u.msn like '%"&request("msn")&"%'"
			else
			sqlstr=sqlstr & " and u.msn like '%"&request("msn")&"%'"
			end if
		end if
		if request("title")<>"" then
			if sqlstr="" then
			sqlstr=" u.title like '%"&request("title")&"%'"
			else
			sqlstr=sqlstr & " and u.title like '%"&request("title")&"%'"
			end if
		end if
		if request("sign")<>"" then
			if sqlstr="" then
			sqlstr=" u.sign like '%"&request("sign")&"%'"
			else
			sqlstr=sqlstr & " and u.sign like '%"&request("sign")&"%'"
			end if
		end if
		if sqlstr="" then
		response.write "<tr><td colspan=6 class=forumrow>��ָ������������</td></tr>"
		response.end
		end if
		sql="select top "&request("searchmax")&" u.userid,u.username,u.useremail,u.LastLogin,u.UserLastIP,u.article,u.UserGroupID from [user] u inner join UserGroups G on u.usergroupid=g.usergroupid where "&sqlstr&" order by u.addDate desc"
	case 10
		sql="select u.userid,u.username,u.useremail,u.LastLogin,u.UserLastIP,u.article,u.UserGroupID from [user] u inner join UserGroups G on u.usergroupid=g.usergroupid where u.usergroupid="&request("usergroupid")&" order by u.addDate desc"
	case else
		response.write "<tr><td colspan=6 class=forumrow>����Ĳ�����</td></tr>"
		response.end
	end select
	'response.write sql
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
		response.write "<tr><td colspan=6 class=forumrow>û���ҵ���ؼ�¼��</td></tr>"
	else
%>
<FORM METHOD=POST ACTION="?action=touser">
<tr align=center>
<td class=forumHeaderBackgroundAlternate><B>�û���</B></td>
<td class=forumHeaderBackgroundAlternate><B>Email</B></td>
<td class=forumHeaderBackgroundAlternate><B>Ȩ��</B></td>
<td class=forumHeaderBackgroundAlternate><B>���IP</B></td>
<td class=forumHeaderBackgroundAlternate><B>����½</B></td>
<td class=forumHeaderBackgroundAlternate><B>����</B></td>
</tr>
<%
		rs.PageSize = Cint(Forum_Setting(11))
		rs.AbsolutePage=currentpage
		page_count=0
		totalrec=rs.recordcount
		while (not rs.eof) and (not page_count = Cint(Forum_Setting(11)))
%>
<tr>
<td class=forumrow><a href="?action=modify&userid=<%=rs("userid")%>"><%=rs("username")%></a></td>
<td class=forumrow width=30% ><a href="mailto:<%=rs("useremail")%>"><%=rs("useremail")%></a></td>
<td class=forumrow width=8% align=center><a href="?action=UserPermission&userid=<%=rs("userid")%>&username=<%=rs("username")%>">�༭</a></td>
<td class=forumrow width=20% ><a href="admin_lockIP.asp?userip=<%=rs("UserLastIP")%>" title="����������û�IP"><%=rs("userlastip")%></a></td>
<td class=forumrow width=15% ><%if rs("lastlogin")<>"" and isdate(rs("lastlogin")) then%><%=Formatdatetime(rs("lastlogin"),2)%><%end if%></td>
<td class=forumrow align=center><input type="checkbox" name="userid" value="<%=rs("userid")%>" <%if rs("userGroupid")=1 then response.write "disabled"%>></td>
</tr>
<%
		page_count = page_count + 1
		rs.movenext
		wend
Pcount=rs.PageCount
%>
<tr><td colspan=6 class=forumrow align=center>��ҳ��
<%
	if currentpage > 4 then
	response.write "<a href=""?page=1&userSearch="&request("userSearch")&"&username="&request("username")&"&useremail="&request("useremail")&"&homepage="&request("homepage")&"&oicq="&request("oicq")&"&icq="&request("icq")&"&msn="&request("msn")&"&title="&request("title")&"&sign="&request("sign")&"&userclass="&request("userclass")&"&usergroups="&request("usergroups")&"&action="&request("action")&"&usergroupid="&request("usergroupid")&"&searchmax="&request("searchmax")&""">[1]</a> ..."
	end if
	if Pcount>currentpage+3 then
	endpage=currentpage+3
	else
	endpage=Pcount
	end if
	for i=currentpage-3 to endpage
	if not i<1 then
		if i = clng(currentpage) then
        response.write " <font color=red>["&i&"]</font>"
		else
        response.write " <a href=""?page="&i&"&userSearch="&request("userSearch")&"&username="&request("username")&"&useremail="&request("useremail")&"&homepage="&request("homepage")&"&oicq="&request("oicq")&"&icq="&request("icq")&"&msn="&request("msn")&"&title="&request("title")&"&sign="&request("sign")&"&userclass="&request("userclass")&"&usergroups="&request("usergroups")&"&action="&request("action")&"&usergroupid="&request("usergroupid")&"&searchmax="&request("searchmax")&""">["&i&"]</a>"
		end if
	end if
	next
	if currentpage+3 < Pcount then 
	response.write "... <a href=""?page="&Pcount&"&userSearch="&request("userSearch")&"&username="&request("username")&"&useremail="&request("useremail")&"&homepage="&request("homepage")&"&oicq="&request("oicq")&"&icq="&request("icq")&"&msn="&request("msn")&"&title="&request("title")&"&sign="&request("sign")&"&userclass="&request("userclass")&"&usergroups="&request("usergroups")&"&action="&request("action")&"&usergroupid="&request("usergroupid")&"&searchmax="&request("searchmax")&""">["&Pcount&"]</a>"
	end if
%>
</td></tr>
<tr><td colspan=5 class=forumrow align=center><B>��ѡ������Ҫ���еĲ���</B>��ɾ��<input type="radio" name="useraction" value=1>&nbsp;&nbsp;�ƶ����û���<input type="radio" name="useraction" value=2 checked>
<select size=1 name="selusergroup">
<%
set trs=conn.execute("select usergroupid,title from usergroups where not (usergroupid=1 or usergroupid=7) order by usergroupid")
do while not trs.eof
response.write "<option value="&trs(0)&">"&trs(1)&"</option>"
trs.movenext
loop
trs.close
set trs=nothing

%>
</select>
</td>
<td class=forumrow align=center><input type=checkbox value="on" name="chkall" onclick="CheckAll(this.form)">
</td>
</tr>
<tr><td colspan=6 class=forumrow align=center>
<input type=submit name=submit value="ִ��ѡ���Ĳ���"  onclick="{if(confirm('ȷ��ִ��ѡ��Ĳ�����?')){this.document.recycle.submit();return true;}return false;}">
</td></tr>
</FORM>
<%
	end if
	rs.close
	set rs=nothing	
elseif request("action")="touser" then
	response.write "<tr><th colspan=6 height=23 align=left>ִ�н��</th></tr>"
	if request("useraction")="" then
		response.write "<tr><td colspan=6 class=forumrow>��ָ����ز�����</td></tr>"
		founderr=true
	end if
	if request("userid")="" then
		response.write "<tr><td colspan=6 class=forumrow>��ѡ������û���</td></tr>"
		founderr=true
	end if
	if not founderr then
		if request("useraction")=1 then
			conn.execute("delete from [user] where userid in ("&replace(request("userid"),"'","")&")")
			response.write "<tr><td colspan=6 class=forumrow>�����ɹ���</td></tr>"
		elseif request("useraction")=2 then
			conn.execute("update [user] set UserGroupID="&replace(request("selusergroup"),"'","")&" where userid in ("&replace(request("userid"),"'","")&")")
			response.write "<tr><td colspan=6 class=forumrow>�����ɹ���</td></tr>"
		else
		response.write "<tr><td colspan=6 class=forumrow>����Ĳ�����</td></tr>"
		end if
	end if
elseif request("action")="modify" then
dim realname,character,personal,country,province,city,shengxiao,blood,belief,occupation,marital, education,college,userphone,iaddress
	response.write "<tr><th colspan=6 height=23 align=left>�û����ϲ���</th></tr>"
	if not isnumeric(request("userid")) then
		response.write "<tr><td colspan=6 class=forumrow>������û�������</td></tr>"
		founderr=true
	end if
	if not founderr then
		Set rs= Server.CreateObject("ADODB.Recordset")
		sql="select * from [user] where userid="&request("userid")
		rs.open sql,conn,1,1
		if rs.eof and rs.bof then
		response.write "<tr><td colspan=6 class=forumrow>û���ҵ�����û���</td></tr>"
		founderr=true
		else
if rs("userinfo")<>"" then
	userinfo=split(rs("userinfo"),"|||")
	if ubound(userinfo)=14 then
		realname=userinfo(0)
		character=userinfo(1)
		personal=userinfo(2)
		country=userinfo(3)
		province=userinfo(4)
		city=userinfo(5)
		shengxiao=userinfo(6)
		blood=userinfo(7)
		belief=userinfo(8)
		occupation=userinfo(9)
		marital=userinfo(10)
		education=userinfo(11)
		college=userinfo(12)
		userphone=userinfo(13)
		iaddress=userinfo(14)
	else
		realname=""
		character=""
		personal=""
		country=""
		province=""
		city=""
		shengxiao=""
		blood=""
		belief=""
		occupation=""
		marital=""
		education=""
		college=""
		userphone=""
		iaddress=""
	end if
else
	realname=""
	character=""
	personal=""
	country=""
	province=""
	city=""
	shengxiao=""
	blood=""
	belief=""
	occupation=""
	marital=""
	education=""
	college=""
	userphone=""
	iaddress=""
end if
%>
<FORM METHOD=POST ACTION="?action=saveuserinfo">
<tr>
<td width=20% class=forumrow valign=top>��ݲ���ѡ��</td>
<td width=40% class=forumrow colspan=2 valign=top>
��<a href="mailto:<%=rs("useremail")%>">�� <%=rs("username")%> ���͵����ʼ�</a><BR>
��<a href="messanger.asp?action=new&touser=<%=rs("username")%>" target=_blank>�� <%=rs("username")%> ����һ������</a><BR>
�۲��� <%=rs("username")%> ���������<BR>
��<a href="dispuser.asp?id=<%=rs("userid")%>" target=_blank>Ԥ�� <%=rs("username")%> ����ʾ����</a>
</td>
<td width=40% class=forumrow colspan=3 valign=top>
��<a href="?action=UserPermission&userid=<%=rs("userid")%>&username=<%=rs("username")%>">�༭ <%=rs("username")%> ����̳Ȩ��</a><BR>
�޲鿴 <%=rs("username")%> �������Դ<BR>
��<a href="?action=touser&useraction=1&userid=<%=rs("userid")%>">���û��б�ɾ�� <%=rs("username")%></a><BR>
</td>
</tr>
<tr><th colspan=6 height=23 align=left>�û����������޸ģ���<%=rs("username")%></th></tr>
<tr>
<td width=20% class=forumrow>�û���</td>
<td width=80% class=forumrow colspan=5>
<select size=1 name="usergroups">
<%
set trs=conn.execute("select usergroupid,title from usergroups order by usergroupid")
do while not trs.eof
response.write "<option value="&trs(0)
if rs("usergroupid")=trs(0) then response.write " selected "
response.write ">"&trs(1)&"</option>"
trs.movenext
loop
trs.close
set trs=nothing
%>
</select>
</td>
</tr>
<input name="userid" type=hidden value="<%=rs("userid")%>">
<tr>
<td width=20% class=forumrow>�û���</td>
<td width=80% class=forumrow colspan=5><input size=45 name="username" type=text value="<%=rs("username")%>"></td>
</tr>
<tr>
<td width=20% class=forumrow>��  ��</td>
<td width=80% class=forumrow colspan=5><input size=45 name="password" type=text>&nbsp;������޸�������</td>
</tr>
<tr>
<td width=20% class=forumrow>��������</td>
<td width=80% class=forumrow colspan=5><input size=45 name="quesion" type=text value="<%=rs("quesion")%>"></td>
</tr>
<tr>
<td width=20% class=forumrow>�����</td>
<td width=80% class=forumrow colspan=5><input size=45 name="answer" type=text>&nbsp;������޸�������</td>
</tr>
<tr>
<td width=20% class=forumrow>�û��ȼ�</td>
<td width=80% class=forumrow colspan=5>
<select size=1 name="userclass">
<%
set trs=conn.execute("select usertitle from usertitle order by usertitleid")
do while not trs.eof
response.write "<option value="&trs(0)
if rs("userclass")=trs(0) then response.write " selected "
response.write ">"&trs(0)&"</option>"
trs.movenext
loop
trs.close
set trs=nothing
%>
</select>
</td>
</tr>
<tr>
<td width=20% class=forumrow>Email</td>
<td width=80% class=forumrow colspan=5><input size=45 name="userEmail" type=text value="<%=rs("useremail")%>"></td>
</tr>
<tr>
<td width=20% class=forumrow>������ҳ</td>
<td width=80% class=forumrow colspan=5><input size=45 name="homepage" type=text value="<%=rs("homepage")%>"></td>
</tr>
<tr>
<td width=20% class=forumrow>ͷ��</td>
<td width=80% class=forumrow colspan=5><input size=45 name="face" type=text value="<%=rs("face")%>">&nbsp;��ȣ�<input size=3 name="width" type=text value="<%=rs("width")%>">&nbsp;�߶ȣ�<input size=3 name="height" type=text value="<%=rs("height")%>"></td>
</tr>
<tr>
<td width=20% class=forumrow>OICQ</td>
<td width=80% class=forumrow colspan=5><input size=45 name="oicq" type=text value="<%=rs("oicq")%>"></td>
</tr>
<tr>
<td width=20% class=forumrow>ICQ</td>
<td width=80% class=forumrow colspan=5><input size=45 name="icq" type=text value="<%=rs("icq")%>"></td>
</tr>
<tr>
<td width=20% class=forumrow>MSN</td>
<td width=80% class=forumrow colspan=5><input size=45 name="msn" type=text value="<%=rs("msn")%>"></td>
</tr>
<tr>
<td width=20% class=forumrow>ͷ��</td>
<td width=80% class=forumrow colspan=5><input size=45 name="usertitle" type=text value="<%=rs("title")%>"></td>
</tr>
<tr>
<td width=20% class=forumrow>�ȼ�ͼƬ</td>
<td width=80% class=forumrow colspan=5><input size=45 name="titlepic" type=text value="<%=rs("titlepic")%>"></td>
</tr>
<tr><th colspan=6 height=23 align=left>�û���ֵ�����޸�</th></tr>
<tr>
<td width=20% class=forumrow>��������</td>
<td width=80% class=forumrow colspan=5><input size=45 name="article" type=text value="<%=rs("article")%>"></td>
</tr>
<tr>
<td width=20% class=forumrow>��ɾ����</td>
<td width=80% class=forumrow colspan=5><input size=45 name="Userdel" type=text value="<%=rs("userdel")%>"></td>
</tr>
<tr>
<td width=20% class=forumrow>��������</td>
<td width=80% class=forumrow colspan=5><input size=45 name="userisbest" type=text value="<%=rs("userisbest")%>"></td>
</tr>
<tr>
<td width=20% class=forumrow>��Ǯ</td>
<td width=80% class=forumrow colspan=5><input size=45 name="userwealth" type=text value="<%=rs("userwealth")%>"></td>
</tr>
<tr>
<td width=20% class=forumrow>����</td>
<td width=80% class=forumrow colspan=5><input size=45 name="userep" type=text value="<%=rs("userep")%>"></td>
</tr>
<tr>
<td width=20% class=forumrow>����</td>
<td width=80% class=forumrow colspan=5><input size=45 name="usercp" type=text value="<%=rs("usercp")%>"></td>
</tr>
<tr>
<td width=20% class=forumrow>����</td>
<td width=80% class=forumrow colspan=5><input size=45 name="userpower" type=text value="<%=rs("userpower")%>"></td>
</tr>
<tr><th colspan=6 height=23 align=left>�������</th></tr>
<tr>
<td width=20% class=forumrow>����</td>
<td width=80% class=forumrow colspan=5><input size=45 name="birthday" type=text value="<%=rs("birthday")%>">&nbsp;��ʽ��2001-2-2</td>
</tr>
<tr>
<td width=20% class=forumrow>ע��ʱ��</td>
<td width=80% class=forumrow colspan=5><input size=45 name="adddate" type=text value="<%=rs("adddate")%>"></td>
</tr>
<tr>
<td width=20% class=forumrow>����½</td>
<td width=80% class=forumrow colspan=5><input size=45 name="lastlogin" type=text value="<%=rs("lastlogin")%>"></td>
</tr>
<tr><th colspan=6 height=23 align=left>�û���ϸ����</th></tr>
<tr>
<td width=20% class=forumrow>��ʵ����</td>
<td width=80% class=forumrow colspan=5><input size=45 name="realname" type=text value="<%=realname%>"></td>
</tr>
<tr>
<td width=20% class=forumrow>��������</td>
<td width=80% class=forumrow colspan=5><input size=45 name="country" type=text value="<%=country%>"></td>
</tr>
<tr>
<td width=20% class=forumrow>��ϵ�绰</td>
<td width=80% class=forumrow colspan=5><input size=45 name="userphone" type=text value="<%=userphone%>"></td>
</tr><tr>
<td width=20% class=forumrow>ͨ�ŵ�ַ</td>
<td width=80% class=forumrow colspan=5><input size=45 name="address" type=text value="<%=iaddress%>"></td>
</tr>
<tr>
<td width=20% class=forumrow>ʡ������</td>
<td width=80% class=forumrow colspan=5><input size=45 name="province" type=text value="<%=province%>"></td>
</tr>
<tr>
<td width=20% class=forumrow>�ǡ�����</td>
<td width=80% class=forumrow colspan=5><input size=45 name="city" type=text value="<%=city%>"></td>
</tr><tr>
<td width=20% class=forumrow>������Ф</td>
<td width=80% class=forumrow colspan=5>
<select size=1 name=shengxiao>
<option <%if shengxiao="" then%>selected<%end if%>></option>
<option value=�� <%if shengxiao="��" then%>selected<%end if%>>��</option>
<option value=ţ <%if shengxiao="ţ" then%>selected<%end if%>>ţ</option>
<option value=�� <%if shengxiao="��" then%>selected<%end if%>>��</option>
<option value=�� <%if shengxiao="��" then%>selected<%end if%>>��</option>
<option value=�� <%if shengxiao="��" then%>selected<%end if%>>��</option>
<option value=�� <%if shengxiao="��" then%>selected<%end if%>>��</option>
<option value=�� <%if shengxiao="��" then%>selected<%end if%>>��</option>
<option value=�� <%if shengxiao="��" then%>selected<%end if%>>��</option>
<option value=�� <%if shengxiao="��" then%>selected<%end if%>>��</option>
<option value=�� <%if shengxiao="��" then%>selected<%end if%>>��</option>
<option value=�� <%if shengxiao="��" then%>selected<%end if%>>��</option>
<option value=�� <%if shengxiao="��" then%>selected<%end if%>>��</option>
</select>
</td>
</tr>
<tr>
<td width=20% class=forumrow>Ѫ������</td>
<td width=80% class=forumrow colspan=5>
<select size=1 name=blood>
<option <%if blood="" then%>selected<%end if%>></option>
<option value=A <%if blood="A" then%>selected<%end if%>>A</option>
<option value=B <%if blood="B" then%>selected<%end if%>>B</option>
<option value=AB <%if blood="AB" then%>selected<%end if%>>AB</option>
<option value=O <%if blood="O" then%>selected<%end if%>>O</option>
<option value=���� <%if blood="����" then%>selected<%end if%>>����</option>
</select>
</td>
</tr>
<tr>
<td width=20% class=forumrow>�š�����</td>
<td width=80% class=forumrow colspan=5>
<select size=1 name=belief>
<option <%if belief="" then%>selected<%end if%>></option>
<option value=��� <%if belief="���" then%>selected<%end if%>>���</option>
<option value=���� <%if belief="����" then%>selected<%end if%>>����</option>
<option value=������ <%if belief="������" then%>selected<%end if%>>������</option>
<option value=������ <%if belief="������" then%>selected<%end if%>>������</option>
<option value=�ؽ� <%if belief="�ؽ�" then%>selected<%end if%>>�ؽ�</option>
<option value=�������� <%if belief="��������" then%>selected<%end if%>>��������</option>
<option value=���������� <%if belief="����������" then%>selected<%end if%>>����������</option>
<option value=���� <%if belief="����" then%>selected<%end if%>>����</option>
</select>
</td>
</tr><tr>
<td width=20% class=forumrow>ְ����ҵ</td>
<td width=80% class=forumrow colspan=5>
<select name=occupation>
<option <%if occupation="" then%>selected<%end if%>> </option>
<option value="�ƻ�/����" <%if occupation="�ƻ�/����" then%>selected<%end if%>>�ƻ�/����</option>
<option value=����ʦ <%if occupation="����ʦ" then%>selected<%end if%>>����ʦ</option>
<option value=���� <%if occupation="����" then%>selected<%end if%>>����</option>
<option value=����������ҵ <%if occupation="����������ҵ" then%>selected<%end if%>>����������ҵ</option>
<option value=��ͥ���� <%if occupation="��ͥ����" then%>selected<%end if%>>��ͥ����</option>
<option value="����/��ѵ" <%if occupation="����/��ѵ" then%>selected<%end if%>>����/��ѵ</option>
<option value="�ͻ�����/֧��" <%if occupation="�ͻ�����/֧��" then%>selected<%end if%>>�ͻ�����/֧��</option>
<option value="������/�ֹ�����" <%if occupation="������/�ֹ�����" then%>selected<%end if%>>������/�ֹ�����</option>
<option value=���� <%if occupation="����" then%>selected<%end if%>>����</option>
<option value=��ְҵ <%if occupation="��ְҵ" then%>selected<%end if%>>��ְҵ</option>
<option value="����/�г�/���" <%if occupation="����/�г�/���" then%>selected<%end if%>>����/�г�/���</option>
<option value=ѧ�� <%if occupation="ѧ��" then%>selected<%end if%>>ѧ��</option>
<option value=�о��Ϳ��� <%if occupation="�о��Ϳ���" then%>selected<%end if%>>�о��Ϳ���</option>
<option value="һ�����/�ල" <%if occupation="һ�����/�ල" then%>selected<%end if%>>һ�����/�ල</option>
<option value="����/����" <%if occupation="����/����" then%>selected<%end if%>>����/����</option>
<option value="ִ�й�/�߼�����" <%if occupation="ִ�й�/�߼�����" then%>selected<%end if%>>ִ�й�/�߼�����</option>
<option value="����/����/����" <%if occupation="����/����/����" then%>selected<%end if%>>����/����/����</option>
<option value=רҵ��Ա <%if occupation="רҵ��Ա" then%>selected<%end if%>>רҵ��Ա</option>
<option value="�Թ�/ҵ��" <%if occupation="�Թ�/ҵ��" then%>selected<%end if%>>�Թ�/ҵ��</option>
<option value=���� <%if occupation="����" then%>selected<%end if%>>����</option>
</select>
</td>
</tr>
<tr>
<td width=20% class=forumrow>����״��</td>
<td width=80% class=forumrow colspan=5>
<select size=1 name=marital>
<option <%if marital="" then%>selected<%end if%>></option>
<option value=δ�� <%if marital="δ��" then%>selected<%end if%>>δ��</option>
<option value=�ѻ� <%if marital="�ѻ�" then%>selected<%end if%>>�ѻ�</option>
<option value=���� <%if marital="����" then%>selected<%end if%>>����</option>
<option value=ɥż <%if marital="ɥż" then%>selected<%end if%>>ɥż</option>
</select>
</td>
</tr>
<tr>
<td width=20% class=forumrow>���ѧ��</td>
<td width=80% class=forumrow colspan=5>
<select size=1 name=education>
<option <%if education="" then%>selected<%end if%>></option>
<option value=Сѧ <%if education="Сѧ" then%>selected<%end if%>>Сѧ</option>
<option value=���� <%if education="����" then%>selected<%end if%>>����</option>
<option value=���� <%if education="����" then%>selected<%end if%>>����</option>
<option value=��ѧ <%if education="��ѧ" then%>selected<%end if%>>��ѧ</option>
<option value=˶ʿ <%if education="˶ʿ" then%>selected<%end if%>>˶ʿ</option>
<option value=��ʿ <%if education="��ʿ" then%>selected<%end if%>>��ʿ</option>
</select>
</td>
</tr>
<tr>
<td width=20% class=forumrow>��ҵԺУ</td>
<td width=80% class=forumrow colspan=5><input size=45 name="college" type=text value="<%=college%>"></td>
</tr>
<tr>
<td width=20% class=forumrow>�ԡ���</td>
<td width=80% class=forumrow colspan=5>
<textarea name=character rows=4 cols=80><%=character%></textarea>
</td>
</tr><tr>
<td width=20% class=forumrow>���˼��</td>
<td width=80% class=forumrow colspan=5>
<textarea name=personal rows=4 cols=80><%=personal%></textarea>
</td>
</tr><tr>
<td width=20% class=forumrow>�û�ǩ��</td>
<td width=80% class=forumrow colspan=5>
<textarea name="sign" rows=4 cols=80><%=rs("sign")%></textarea>
</td>
</tr>
<tr><th colspan=6 height=23 align=left>�û�����</th></tr>
<tr>
<td width=20% class=forumrow>�û�״̬</td>
<td width=80% class=forumrow colspan=5>
���� <input type="radio" value="0" <%if rs("lockuser")=0 then%>checked<%end if%> name="lockuser">&nbsp;
���� <input type="radio" value="2" <%if rs("lockuser")=1 then%>checked<%end if%> name="lockuser">&nbsp;
���� <input type="radio" value="3" <%if rs("lockuser")=2 then%>checked<%end if%> name="lockuser">
</td>
</tr>
<tr>
<td width=100% class=forumrow align=center colspan=6><input name="submit" type=submit value="   ��  ��   "></td>
</tr>
</FORM>
<%
		end if
		rs.close
		set rs=nothing
	end if
elseif request("action")="saveuserinfo" then
	response.write "<tr><th colspan=6 height=23 align=left>�����û�����</th></tr>"
	userinfo=checkreal(request.Form("realname")) & "|||" & checkreal(request.Form("character")) & "|||" & checkreal(request.Form("personal")) & "|||" & checkreal(request.Form("country")) & "|||" & checkreal(request.Form("province")) & "|||" & checkreal(request.Form("city")) & "|||" & request.Form("shengxiao") & "|||" & request.Form("blood") & "|||" & request.Form("belief") & "|||" & request.Form("occupation") & "|||" & request.Form("marital") & "|||" & request.Form("education") & "|||" & checkreal(request.Form("college")) & "|||" & checkreal(request.Form("userphone")) & "|||" & checkreal(request.Form("address"))
	if not isnumeric(request("userid")) then
		response.write "<tr><td colspan=6 class=forumrow>������û�������</td></tr>"
		founderr=true
	end if
	if not founderr then
	Set rs= Server.CreateObject("ADODB.Recordset")
	sql="select * from [user] where userid="&request("userid")
	rs.open sql,conn,1,3
	if rs.eof and rs.bof then
		response.write "<tr><td colspan=6 class=forumrow>û���ҵ�����û���</td></tr>"
		founderr=true
	else
		rs("username")=request.form("username")
		if request.form("password")<>"" then
		rs("userpassword")=md5(request.form("password"))
		end if
		rs("usergroupid")=request.form("usergroups")
		rs("quesion")=request.form("quesion")
		if request.form("answer")<>"" then
		rs("answer")=md5(request.form("answer"))
		end if
		rs("userclass")=request.form("userclass")
		rs("useremail")=request.form("useremail")
		rs("homepage")=request.form("homepage")
		rs("face")=request.form("face")
		rs("width")=request.form("width")
		rs("height")=request.form("height")
		rs("oicq")=request.form("oicq")
		rs("icq")=request.form("icq")
		rs("msn")=request.form("msn")
		rs("title")=request.form("usertitle")
		rs("titlepic")=request.form("titlepic")
		rs("article")=request.form("article")
		rs("userdel")=request.form("userdel")
		rs("userisbest")=request.form("userisbest")
		rs("userpower")=request.form("userpower")
		rs("userwealth")=request.form("userwealth")
		rs("userep")=request.form("userep")
		rs("usercp")=request.form("usercp")
		rs("birthday")=request.form("birthday")
		rs("addDate")=request.form("adddate")
		rs("lastlogin")=request.form("lastlogin")
		rs("lockuser")=request.form("lockuser")
		rs("sign")=request.form("sign")
		rs("userinfo")=userinfo
		rs.update
	end if
	rs.close
	set rs=nothing
	end if
	if founderr then
		response.write "<tr><td colspan=6 class=forumrow>����ʧ�ܡ�</td></tr>"
	else
		response.write "<tr><td colspan=6 class=forumrow>�����û����ݳɹ���</td></tr>"
	end if
elseif request("action")="UserPermission" then
	response.write "<tr><th colspan=6 height=23 align=left>�༭"&request("username")&"��̳Ȩ�ޣ���ɫ��ʾ���û��ڸð������Զ���Ȩ�ޣ�</th></tr>"
	if not isnumeric(request("userid")) then
		response.write "<tr><td colspan=6 class=forumrow>������û�������</td></tr>"
		founderr=true
	end if
	if not founderr then
	response.write "<tr><td colspan=6 class=forumrow height=25><a href=?action=userBoardPermission&boardid=0&userid="&request("userid")&">�༭���û�������ҳ���Ȩ��</a>����Ҫ��Զ��Ų������ã�</td></tr>"
'----------------------boardinfo--------------------
response.write "<tr><td colspan=6 class=forumrow>"
dim rs_1,rs_2
dim sql_1,sql_2
	sql_2 = "select * from class order by orders"
	set rs_2=conn.execute(sql_2)
	do while not rs_2.Eof
%>
            <table width="95%" border="0" cellspacing="3" cellpadding="0" align=center>
              <tr>

                <td height="21"><B><%=rs_2("class")%></b></td>
              </tr>
            </table>
<%
		sql_1 = "select boardid,boardtype from board where class="&rs_2("id")&" order by orders"
		set rs_1=conn.execute(sql_1)
		do while not rs_1.EOF 
%>
            <table width="95%" border="0" cellspacing="3" cellpadding="0" align=center>
              <tr> 
                <td height="18" width=50% ><li><%=rs_1("boardtype")%></td>
				<td width=50% >
<a href="?action=userBoardPermission&boardid=<%=rs_1("boardid")%>&userid=<%=request("userid")%>">
<%
set trs=conn.execute("select uc_userid from UserAccess where uc_Boardid="&rs_1("boardid")&" and uc_userid="&request("userid"))
if not (trs.eof and trs.bof) then
	response.write "<font color=red>"
end if
%>
�༭�û��ڸ���̳Ȩ��</a></td>
              </tr>
            </table>
<%
		rs_1.MoveNext
		loop
%>
            <table width="95%" border="0" cellspacing="3" cellpadding="0" align=center>
              <tr>
                <td height="21" align=left><hr color=black height=1 width=70% size=1></td>
              </tr>
            </table>
<%
	rs_2.MoveNext 
	Loop
	set trs=nothing
set rs_1=nothing
set rs_2=nothing
response.write "</td></tr>"
'-------------------end-------------------
	end if
elseif request("action")="userBoardPermission" then
	if not isnumeric(request("userid")) then
		response.write "<tr><td colspan=6 class=forumrow>������û�������</td></tr>"
		founderr=true
	end if
	if not isnumeric(request("boardid")) then
		response.write "<tr><td colspan=6 class=forumrow>����İ��������</td></tr>"
		founderr=true
	end if
	if not founderr then
	set rs=conn.execute("select u.UserGroupID,ug.title,u.username from [user] u inner join UserGroups UG on u.userGroupID=ug.userGroupID where u.userid="&request("userid"))
	UserGroupID=rs(0)
	usertitle=rs(1)
	membername=rs(2)
	set rs=conn.execute("select boardtype from board where boardid="&request("boardid"))
	if rs.eof and rs.bof then
	boardtype="��̳����ҳ��"
	else
	boardtype=rs(0)
	end if
	response.write "<tr><th colspan=6 height=23 align=left>�༭ "&membername&" �� "&boardtype&" Ȩ��</th></tr>"
	response.write "<tr><td colspan=6 height=25 class=forumrow>ע�⣺���û����� <B>"&usertitle&"</B> �û����У�����������������Զ���Ȩ�ޣ�����û�Ȩ�޽����Զ���Ȩ��Ϊ��</td></tr>"
%>
<tr><td colspan=6 class=forumrow>
<%
Dim reGroupSetting
Dim FoundGroup,FoundUserPermission,FoundGroupPermission
FoundGroup=false
FoundUserPermission=false
FoundGroupPermission=false

set rs=conn.execute("select * from UserAccess where uc_boardid="&request("boardid")&" and uc_userid="&request("userid"))
if not (rs.eof and rs.bof) then
	reGroupSetting=split(rs("uc_Setting"),",")
	FoundGroup=true
	FoundUserPermission=true
end if

if not foundgroup then
set rs=conn.execute("select * from BoardPermission where boardid="&request("boardid")&" and groupid="&usergroupid)
if not(rs.eof and rs.bof) then
	reGroupSetting=split(rs("PSetting"),",")
	FoundGroup=true
	FoundGroupPermission=true
end if
end if

if not foundgroup then
set rs=conn.execute("select * from usergroups where usergroupid="&usergroupid)
if rs.eof and rs.bof then
	response.write "δ�ҵ����û��飡"
	response.end
else
	FoundGroup=true
	FoundGroupPermission=true
	reGroupSetting=split(rs("GroupSetting"),",")
end if
end if
%>
<table width="100%" border="0" cellspacing="1" cellpadding="0"  align=center>
<FORM METHOD=POST ACTION="?action=saveuserpermission">
<input type=hidden name="userid" value="<%=request("userid")%>">
<input type=hidden name="BoardID" value="<%=request("boardid")%>">
<input type=hidden name="username" value="<%=membername%>">
<tr> 
<td height="23" colspan="2" class=forumrow><input type=radio name="isdefault" value="1" <%if FoundGroupPermission then%>checked<%end if%>><B>ʹ���û���Ĭ��ֵ</B> (ע��: �⽫ɾ���κ�֮ǰ�������Զ�������)</td>
</tr>
<tr> 
<td height="23" colspan="2"  class=forumrow><input type=radio name="isdefault" value="0" <%if FoundUserPermission then%>checked<%end if%>><B>ʹ���Զ�������</B> </td>
</tr>
<tr> 
<th height="23" colspan="2" align=left>�����鿴Ȩ��</th>
</tr>
<tr>
<td height="23" width="60%" class=forumrow>���������̳</td>
<td height="23" width="40%" class=forumrow>��<input name="canview" type=radio value="1" <%if reGroupSetting(0)=1 then%>checked<%end if%>>&nbsp;��<input name="canview" type=radio value="0" <%if reGroupSetting(0)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=forumrow>���Բ鿴��Ա��Ϣ(����������Ա�����Ϻͻ�Ա�б�)
</td>
<td height="23" width="40%" class=forumrow>��<input name="canviewuserinfo" type=radio value="1" <%if reGroupSetting(1)=1 then%>checked<%end if%>>&nbsp;��<input name="canviewuserinfo" type=radio value="0" <%if reGroupSetting(1)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=forumrow>���Բ鿴�����˷���������
</td>
<td height="23" width="40%" class=forumrow>��<input name="canviewpost" type=radio value="1" <%if reGroupSetting(2)=1 then%>checked<%end if%>>&nbsp;��<input name="canviewpost" type=radio value="0" <%if reGroupSetting(2)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>���������������
</td>
<td height="23" width="40%" class=Forumrow>��<input name="canviewbest" type=radio value="1" <%if reGroupSetting(41)=1 then%>checked<%end if%>>&nbsp;��<input name="canviewbest" type=radio value="0" <%if reGroupSetting(41)=0 then%>checked<%end if%>></td>
</tr>
<tr> 
<th height="23" colspan="2" align=left>����<b>����Ȩ��</b></th>
</tr>
<tr>
<td height="23" width="60%" class=forumrow>���Է���������</td>
<td height="23" width="40%" class=forumrow>��<input name="cannewpost" type=radio value="1" <%if reGroupSetting(3)=1 then%>checked<%end if%>>&nbsp;��<input name="cannewpost" type=radio value="0" <%if reGroupSetting(3)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=forumrow>���Իظ��Լ�������
</td>
<td height="23" width="40%" class=forumrow>��<input name="canreplymytopic" type=radio value="1" <%if reGroupSetting(4)=1 then%>checked<%end if%>>&nbsp;��<input name="canreplymytopic" type=radio value="0" <%if reGroupSetting(4)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=forumrow>���Իظ������˵�����
</td>
<td height="23" width="40%" class=forumrow>��<input name="canreplytopic" type=radio value="1" <%if reGroupSetting(5)=1 then%>checked<%end if%>>&nbsp;��<input name="canreplytopic" type=radio value="0" <%if reGroupSetting(5)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=forumrow>��������̳�������ֵ�ʱ���������(�ʻ��ͼ���)?
</td>
<td height="23" width="40%" class=forumrow>��<input name="canpostagree" type=radio value="1" <%if reGroupSetting(6)=1 then%>checked<%end if%>>&nbsp;��<input name="canpostagree" type=radio value="0" <%if reGroupSetting(6)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>�������������Ǯ
</td>
<td height="23" width="40%" class=Forumrow><input name="postagreemoney" type=text size=4 value="<%=reGroupSetting(47)%>"></td>
</tr>
<tr>
<td height="23" width="60%" class=forumrow>�����ϴ�����
</td>
<td height="23" width="40%" class=forumrow>��<input name="canupload" type=radio value="1" <%if reGroupSetting(7)=1 then%>checked<%end if%>>&nbsp;��<input name="canupload" type=radio value="0" <%if reGroupSetting(7)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>����ϴ��ļ�����
</td>
<td height="23" width="40%" class=Forumrow><input name="canuploadnum" type=text size=4 value="<%=reGroupSetting(40)%>"></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>�ϴ��ļ���С����
</td>
<td height="23" width="40%" class=Forumrow><input name="MaxUploadSize" type=text size=4 value="<%=reGroupSetting(44)%>"> KB</td>
</tr>
<tr>
<td height="23" width="60%" class=forumrow>���Է�����ͶƱ</td>
<td height="23" width="40%" class=forumrow>��<input name="canpostvote" type=radio value="1" <%if reGroupSetting(8)=1 then%>checked<%end if%>>&nbsp;��<input name="canpostvote" type=radio value="0" <%if reGroupSetting(8)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=forumrow>���Բ���ͶƱ</td>
<td height="23" width="40%" class=forumrow>��<input name="canvote" type=radio value="1" <%if reGroupSetting(9)=1 then%>checked<%end if%>>&nbsp;��<input name="canvote" type=radio value="0" <%if reGroupSetting(9)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>���Է���С�ֱ�</td>
<td height="23" width="40%" class=Forumrow>��<input name="cansmallpaper" type=radio value="1"  <%if reGroupSetting(17)=1 then%>checked<%end if%>>&nbsp;��<input name="cansmallpaper" type=radio value="0" <%if reGroupSetting(17)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>����С�ֱ������Ǯ</td>
<td height="23" width="40%" class=Forumrow><input name="smallpapermoney" type=text value="<%=reGroupSetting(46)%>" size=4></td>
</tr>
<tr> 
<th height="23" colspan="2" align=left>����<b>����/����༭Ȩ��</b></th>
</tr>
<tr>
<td height="23" width="60%" class=forumrow>���Ա༭�Լ�������
</td>
<td height="23" width="40%" class=forumrow>��<input name="caneditmytopic" type=radio value="1" <%if reGroupSetting(10)=1 then%>checked<%end if%>>&nbsp;��<input name="caneditmytopic" type=radio value="0" <%if reGroupSetting(10)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=forumrow>����ɾ���Լ�������
</td>
<td height="23" width="40%" class=forumrow>��<input name="candelmytopic" type=radio value="1" <%if reGroupSetting(11)=1 then%>checked<%end if%>>&nbsp;��<input name="candelmytopic" type=radio value="0" <%if reGroupSetting(11)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=forumrow>�����ƶ��Լ������ӵ�������̳
</td>
<td height="23" width="40%" class=forumrow>��<input name="canmovemytopic" type=radio value="1" <%if reGroupSetting(12)=1 then%>checked<%end if%>>&nbsp;��<input name="canmovemytopic" type=radio value="0" <%if reGroupSetting(12)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=forumrow>���Դ�/�ر��Լ�����������
</td>
<td height="23" width="40%" class=forumrow>��<input name="canclosemytopic" type=radio value="1" <%if reGroupSetting(13)=1 then%>checked<%end if%>>&nbsp;��<input name="canclosemytopic" type=radio value="0" <%if reGroupSetting(13)=0 then%>checked<%end if%>></td>
</tr>
<tr> 
<th height="23" colspan="2" align=left>����<b>����Ȩ��</b></th>
</tr>
<tr>
<td height="23" width="60%" class=forumrow>����������̳
</td>
<td height="23" width="40%" class=forumrow>��<input name="cansearch" type=radio value="1" <%if reGroupSetting(14)=1 then%>checked<%end if%>>&nbsp;��<input name="cansearch" type=radio value="0" <%if reGroupSetting(14)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=forumrow>����ʹ��'���ͱ�ҳ������'����
</td>
<td height="23" width="40%" class=forumrow>��<input name="canmailtopic" type=radio value="1" <%if reGroupSetting(15)=1 then%>checked<%end if%>>&nbsp;��<input name="canmailtopic" type=radio value="0" <%if reGroupSetting(15)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=forumrow>�����޸ĸ�������
</td>
<td height="23" width="40%" class=forumrow>��<input name="canmodify" type=radio value="1" <%if reGroupSetting(16)=1 then%>checked<%end if%>>&nbsp;��<input name="canmodify" type=radio value="0" <%if reGroupSetting(16)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>�����п��Բ鿴������ͷ��
</td>
<td height="23" width="40%" class=Forumrow>��<input name="canusetitle" type=radio value="1" <%if reGroupSetting(36)=1 then%>checked<%end if%>>&nbsp;��<input name="canusetitle" type=radio value="0" <%if reGroupSetting(36)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>�����п��Բ鿴������ͷ��
</td>
<td height="23" width="40%" class=Forumrow>��<input name="canuseface" type=radio value="1"  <%if reGroupSetting(37)=1 then%>checked<%end if%>>&nbsp;��<input name="canuseface" type=radio value="0" <%if reGroupSetting(37)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>�����п��Կ��Բ鿴������ǩ��
</td>
<td height="23" width="40%" class=Forumrow>��<input name="canusesign" type=radio value="1"  <%if reGroupSetting(38)=1 then%>checked<%end if%>>&nbsp;��<input name="canusesign" type=radio value="0" <%if reGroupSetting(38)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>���������̳�¼�
</td>
<td height="23" width="40%" class=Forumrow>��<input name="canvieweven" type=radio value="1"  <%if reGroupSetting(39)=1 then%>checked<%end if%>>&nbsp;��<input name="canvieweven" type=radio value="0" <%if reGroupSetting(39)=0 then%>checked<%end if%>></td>
</tr>
<tr> 
<th height="23" colspan="2" align=left>��������Ȩ��</th>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>����ɾ������������
</td>
<td height="23" width="40%" class=Forumrow>��<input name="candeltopic" type=radio value="1" <%if reGroupSetting(18)=1 then%>checked<%end if%>>&nbsp;��<input name="candeltopic" type=radio value="0"  <%if reGroupSetting(18)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>�����ƶ�����������
</td>
<td height="23" width="40%" class=Forumrow>��<input name="canmovetopic" type=radio value="1" <%if reGroupSetting(19)=1 then%>checked<%end if%>>&nbsp;��<input name="canmovetopic" type=radio value="0"  <%if reGroupSetting(19)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>���Դ�/�ر�����������
</td>
<td height="23" width="40%" class=Forumrow>��<input name="canclosetopic" type=radio value="1" <%if reGroupSetting(20)=1 then%>checked<%end if%>>&nbsp;��<input name="canclosetopic" type=radio value="0"  <%if reGroupSetting(20)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>���Թ̶�/����̶�����
</td>
<td height="23" width="40%" class=Forumrow>��<input name="cantoptopic" type=radio value="1" <%if reGroupSetting(21)=1 then%>checked<%end if%>>&nbsp;��<input name="cantoptopic" type=radio value="0"  <%if reGroupSetting(21)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>���Խ���/�ͷ������û�
</td>
<td height="23" width="40%" class=Forumrow>��<input name="canawardtopic" type=radio value="1" <%if reGroupSetting(22)=1 then%>checked<%end if%>>&nbsp;��<input name="canawardtopic" type=radio value="0"  <%if reGroupSetting(22)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>���Խ���/�ͷ��û�
</td>
<td height="23" width="40%" class=Forumrow>��<input name="canaward" type=radio value="1" <%if reGroupSetting(43)=1 then%>checked<%end if%>>&nbsp;��<input name="canaward" type=radio value="0"  <%if reGroupSetting(43)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>���Ա༭����������
</td>
<td height="23" width="40%" class=Forumrow>��<input name="canmodifytopic" type=radio value="1" <%if reGroupSetting(23)=1 then%>checked<%end if%>>&nbsp;��<input name="canmodifytopic" type=radio value="0" <%if reGroupSetting(23)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>���Լ���/�����������
</td>
<td height="23" width="40%" class=Forumrow>��<input name="canbesttopic" type=radio value="1" <%if reGroupSetting(24)=1 then%>checked<%end if%>>&nbsp;��<input name="canbesttopic" type=radio value="0"  <%if reGroupSetting(24)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>���Է�������
</td>
<td height="23" width="40%" class=Forumrow>��<input name="canAnnounce" type=radio value="1" <%if reGroupSetting(25)=1 then%>checked<%end if%>>&nbsp;��<input name="canAnnounce" type=radio value="0"  <%if reGroupSetting(25)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>���Թ�����
</td>
<td height="23" width="40%" class=Forumrow>��<input name="canAdminAnnounce" type=radio value="1" <%if reGroupSetting(26)=1 then%>checked<%end if%>>&nbsp;��<input name="canAdminAnnounce" type=radio value="0"  <%if reGroupSetting(26)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>���Թ���С�ֱ�
</td>
<td height="23" width="40%" class=Forumrow>��<input name="canAdminPaper" type=radio value="1" <%if reGroupSetting(27)=1 then%>checked<%end if%>>&nbsp;��<input name="canAdminPaper" type=radio value="0"  <%if reGroupSetting(27)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>��������/����/��������û�
</td>
<td height="23" width="40%" class=Forumrow>��<input name="canAdminUser" type=radio value="1" <%if reGroupSetting(28)=1 then%>checked<%end if%>>&nbsp;��<input name="canAdminUser" type=radio value="0"  <%if reGroupSetting(28)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>����ɾ���û�1��10������������
</td>
<td height="23" width="40%" class=Forumrow>��<input name="canDelUserTopic" type=radio value="1" <%if reGroupSetting(29)=1 then%>checked<%end if%>>&nbsp;��<input name="canDelUserTopic" type=radio value="0"  <%if reGroupSetting(29)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>���Բ鿴����IP����Դ
</td>
<td height="23" width="40%" class=Forumrow>��<input name="canviewip" type=radio value="1" <%if reGroupSetting(30)=1 then%>checked<%end if%>>&nbsp;��<input name="canviewip" type=radio value="0"  <%if reGroupSetting(30)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>�����޶�IP����
</td>
<td height="23" width="40%" class=Forumrow>��<input name="canadminip" type=radio value="1" <%if reGroupSetting(31)=1 then%>checked<%end if%>>&nbsp;��<input name="canadminip" type=radio value="0"  <%if reGroupSetting(31)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>���Թ����û�Ȩ��
</td>
<td height="23" width="40%" class=Forumrow>��<input name="adminpermission" type=radio value="1" <%if reGroupSetting(42)=1 then%>checked<%end if%>>&nbsp;��<input name="adminpermission" type=radio value="0"  <%if reGroupSetting(42)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>��������ɾ�����ӣ�ǰ̨��
</td>
<td height="23" width="40%" class=Forumrow>��<input name="canbatchtopic" type=radio value="1" <%if reGroupSetting(45)=1 then%>checked<%end if%>>&nbsp;��<input name="canbatchtopic" type=radio value="0"  <%if reGroupSetting(45)=0 then%>checked<%end if%>></td>
</tr>
<tr> 
<th height="23" colspan="2" align=left>��������Ȩ��</th>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>���Է��Ͷ���
</td>
<td height="23" width="40%" class=Forumrow>��<input name="cansendsms" type=radio value="1"  <%if reGroupSetting(32)=1 then%>checked<%end if%>>&nbsp;��<input name="cansendsms" type=radio value="0" <%if reGroupSetting(32)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>��෢���û�
</td>
<td height="23" width="40%" class=Forumrow><input name="Maxsendsms" size=5 type=text value="<%=reGroupSetting(33)%>"></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>�������ݴ�С����
</td>
<td height="23" width="40%" class=Forumrow><input name="Maxsmsbody" size=5 type=text value="<%=reGroupSetting(34)%>"> byte</td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>�����С����
</td>
<td height="23" width="40%" class=Forumrow><input name="Maxsmsbox" size=5 type=text value="<%=reGroupSetting(35)%>"> KB</td>
</tr>
<tr> 
<td height="23" width="60%" class=forumrow>
</td>
<td height="23" width="40%" class=forumrow><input type="submit" name="submit" value="�� ��"></td>
</tr>
<input type=hidden value="yes" name="groupaction">
</FORM>
</table>
</td></tr>
<%
	end if
elseif request("action")="saveuserpermission" then
	response.write "<tr><th colspan=6 height=23 align=left>�༭�û� "&request("username")&" Ȩ��</th></tr>"
	if not isnumeric(request("userid")) then
		response.write "<tr><td colspan=6 class=forumrow>������û�������</td></tr>"
		founderr=true
	end if
	if not isnumeric(request("boardid")) then
		response.write "<tr><td colspan=6 class=forumrow>����İ��������</td></tr>"
		founderr=true
	end if
	if not founderr then
	dim myGroupSetting
	myGroupSetting=Request.Form("canview") & "," & Request.Form("canviewuserinfo") & "," & Request.Form("canviewpost") & "," & Request.Form("cannewpost") & "," & Request.Form("canreplymytopic") & "," & Request.Form("canreplytopic") & "," & Request.Form("canpostagree") & "," & Request.Form("canupload") & "," & Request.Form("canpostvote") & "," & Request.Form("canvote") & "," & Request.Form("caneditmytopic") & "," & Request.Form("candelmytopic") & "," & Request.Form("canmovemytopic") & "," & Request.Form("canclosemytopic") & "," & Request.Form("cansearch") & "," & Request.Form("canmailtopic") & "," & Request.Form("canmodify") & "," & Request.Form("cansmallpaper") & "," & Request.Form("candeltopic") & "," & Request.Form("canmovetopic") & "," & Request.Form("canclosetopic") & "," & Request.Form("cantoptopic") & "," & Request.Form("canawardtopic") & "," & Request.Form("canmodifytopic") & "," & Request.Form("canbesttopic") & "," & Request.Form("canAnnounce") & "," & Request.Form("canAdminAnnounce") & "," & Request.Form("canAdminPaper") & "," & Request.Form("canAdminUser") & "," & Request.Form("canDelUserTopic") & "," & Request.Form("canviewip") & "," & Request.Form("canadminip") & "," & Request.Form("cansendsms") & "," & Request.Form("Maxsendsms") & "," & Request.Form("Maxsmsbody") & "," & Request.Form("Maxsmsbox") & "," & Request.Form("canusetitle") & "," & Request.Form("canuseface") & "," & Request.Form("canusesign") & "," & Request.Form("canvieweven") & "," & Request.Form("canuploadnum") & "," & Request.Form("canviewbest") & "," & Request.Form("adminpermission") & "," & request.form("canaward") & "," & request.form("MaxUploadSize") & "," & request.form("canbatchtopic") & "," & request.form("smallpapermoney") & "," & request.form("postagreemoney")
	if request("isdefault")=1 then
		conn.execute("delete from UserAccess where uc_boardid="&request("boardid")&" and uc_userid="&request("userid"))
	else
		set rs=conn.execute("select * from UserAccess where uc_boardid="&request("boardid")&" and uc_userid="&request("userid"))
		if rs.eof and rs.bof then
			conn.execute("insert into UserAccess (uc_userid,uc_boardid,uc_setting) values ("&request("userid")&","&request("boardid")&",'"&myGroupSetting&"')")
		else
			conn.execute("update UserAccess set uc_setting='"&myGroupSetting&"' where uc_boardid="&request("boardid")&" and uc_userid="&request("userid"))
		end if
	end if
	if founderr then
		response.write "<tr><td colspan=6 class=forumrow>����ʧ�ܡ�</td></tr>"
	else
		response.write "<tr><td colspan=6 class=forumrow>�����û�Ȩ�޳ɹ���</td></tr>"
	end if
	end if
end if
function checkreal(v)
	dim w
	if not isnull(v) then
	w=replace(v,"|||","����")
	checkreal=w
	end if
end function
conn.close
set conn=nothing
%>
</table>
<p></p>

<script language="JavaScript">
<!--
function CheckAll(form)  {
  for (var i=0;i<form.elements.length;i++)    {
    var e = form.elements[i];
    if (e.name != 'chkall')       e.checked = form.chkall.checked; 
   }
  }
//-->
</script>