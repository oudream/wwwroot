<!--#include file=conn.asp-->
<!-- #include file="inc/const.asp" -->
<title><%=Forum_info(0)%>--����ҳ��</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta NAME=GENERATOR Content="Microsoft FrontPage 3.0" CHARSET=GB2312>
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0" marginheight="0" marginwidth="0">
<%
	if not master or session("flag")="" then
		Errmsg=Errmsg+"<br>"+"<li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��½</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
		call dvbbs_error()
	else
		if request("action")="save" then
		call savegrade()
		elseif request("action")="add" then
		call add()
		elseif request("action")="savenew" then
		call savenew()
		elseif request("action")="del" then
		call del()
		else
		call gradeinfo()
		end if
		conn.close
		set conn=nothing
	end if


sub gradeinfo()
%>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center>
<tr>
<td valign=top>
<B>������ID��Ϣ</B>��<BR>
<%
	set rs=conn.execute("select * from usergroups order by usergroupid")
	do while not rs.eof
		response.write "�û��飺"&rs("title")&"��ID��<B>" & rs("usergroupid") & "</B><br>"
	rs.movenext
	loop
	rs.close
	set rs=nothing
%><BR>
����û�������޶�Ӧ�ȼ����ƣ���ע���û��Զ�������������<BR>
����û���ĵȼ����ƿ��Ժ��û�������һ��
</td>
<td width="50%" valign=top>
<B>�ڵȼ����趨�û�����ʲô�ã�</B><BR>
һ����˵��ֻ��ע���û�ӵ�еȼ��������ڵȼ���һ�㶼�趨�û���IDΪ4��Ӧע���û��飬������óɱ����ID����ô���û�������������ȼ���ͬʱҲ���Զ����������õ���<BR>
�������������һ���û��飬���Ҹ���������û���ĳһЩȨ�ޣ���ô��������ôﵽһ���ȼ������ӣ����û��Զ����µ�����û�����ʹ������û����Ȩ�ޡ�<BR>�������ĳ���ȼ����û��������������������ȼ�����ô�Ͱ����ٷ�������Ϊ<B>-1</B>��һ��Ϊ�����û�����Ҫ���������ã�����ĳ���������ٷ���Ϊ<B>-1</B>�󣬸ü�����û������ܸ����������Ӷ�����������û�Ҳ�����Զ��������ü���ֻ�����û������з��ܸ����伶��
</td>
</tr>
</table>
<form method="POST" action=admin_grade.asp?action=save>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th height="23" colspan="6" >�û��ȼ��趨</th>
</tr>
<tr> 
<td width="10%" class=forumHeaderBackgroundAlternate><b>�ȼ�<B></td>
<td width="25%" class=forumHeaderBackgroundAlternate><B>����</B></td>
<td width="15%" class=forumHeaderBackgroundAlternate><B>���ٷ���</B></td>
<td width="25%" class=forumHeaderBackgroundAlternate><B>ͼƬ</B></td>
<td width="15%" class=forumHeaderBackgroundAlternate><B>�����ID</B></td>
<td width="10%" class=forumHeaderBackgroundAlternate><B>����</B></td>
</tr>
<%
set rs=conn.execute("select * from usertitle order by minarticle,userclass desc")
do while not rs.eof
%>
<tr> 
<td width="10%" class=Forumrow><input type=hidden value="<%=rs("usertitleid")%>" name="usertitleid"><input size=3 value="<%=rs("userclass")%>" name="userclass" type=text></td>
<td width="25%" class=Forumrow><input size=15 value="<%=rs("usertitle")%>" name="usertitle" type=text></td>
<td width="15%" class=Forumrow><input size=5 value="<%=rs("MinArticle")%>" name="minarticle" type=text></td>
<td width="25%" class=Forumrow><input size=15 value="<%=rs("titlepic")%>" name="titlepic" type=text></td>
<td width="15%" class=Forumrow><input size=5 value="<%=rs("usergroupid")%>" name="groupid" type=text></td>
<td width="10%" class=Forumrow><a href="?action=del&id=<%=rs("usertitleid")%>">ɾ��</a></td>
</tr>
<%
rs.movenext
loop
rs.close
set rs=nothing
%>
<tr> 
<td width="100%" colspan=6 class=Forumrow> 
<input type="submit" name="Submit" value="�� ��">
</td>
</tr>
</table>
</form>
<%
end sub

sub savegrade()
	Server.ScriptTimeout=99999999
	dim usertitleid,iuserclass,usertitle,Minarticle,titlepic,groupid
	for i=1 to request.form("usertitleid").count
	usertitleid=replace(request.form("usertitleid")(i),"'","")
	iuserclass=replace(request.form("userclass")(i),"'","")
	usertitle=replace(request.form("usertitle")(i),"'","")
	minarticle=replace(request.form("minarticle")(i),"'","")
	titlepic=replace(request.form("titlepic")(i),"'","")
	groupid=replace(request.form("groupid")(i),"'","")
	if isnumeric(usertitleid) and isnumeric(iuserclass) and usertitle<>"" and isnumeric(minarticle) and titlepic<>"" and isnumeric(groupID) then
	set rs=conn.execute("select * from usertitle where usertitleid="&usertitleID)
	if rs("usertitle")<>trim(usertitle) or rs("titlepic")<>trim(titlepic) or rs("usergroupid")<>cint(groupid) then
	conn.execute("update [user] set userclass='"&usertitle&"',titlepic='"&titlepic&"',usergroupid="&groupid&" where userclass='"&rs("usertitle")&"'")
	end if
	conn.execute("update usertitle set userclass="&iuserclass&",usertitle='"&usertitle&"',minarticle="&minarticle&",titlepic='"&titlepic&"',usergroupid="&groupid&" where usertitleid="&usertitleID)
	end if
	next
	response.write "���óɹ����뷵�ء�"
	set rs=nothing
end sub

sub add()
%>
<form method="POST" action=admin_grade.asp?action=savenew>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th colspan="2">����µ��û��ȼ�</th>
</tr>
<tr> 
<td width="40%" class=forumrow><b>�ȼ�<B></td>
<td width="60%" class=forumrow><input size=30 name="userclass" type=text></td>
</tr>
<tr>
<td width="40%" class=forumrow><B>�����û���</B></td>
<td width="60%" class=forumrow>
<select size=1 name="usergroupid">
<%
set rs=conn.execute("select * from usergroups order by usergroupid")
do while not rs.eof
%>
<option value="<%=rs("usergroupid")%>" <%if rs("usergroupid")=4 then%>selected<%end if%>><%=rs("title")%></option>
<%
rs.movenext
loop
rs.close
set rs=nothing
%>
</select>
</td>
</tr>
<tr>
<td width="40%" class=forumrow><B>����</B></td>
<td width="60%" class=forumrow><input size=30 name="usertitle" type=text></td>
</tr>
<tr>
<td width="40%" class=forumrow><B>���ٷ���</B><BR>����õȼ��������ƺŻ��߹�����ݣ����������д-1����ʾ��������������������</td>
<td width="60%" class=forumrow><input size=30 name="minarticle" type=text></td>
</tr>
<tr>
<td width="40%" class=forumrow><B>ͼƬ</B></td>
<td width="60%" class=forumrow><input size=30 name="titlepic" type=text></td>
</tr>
<tr> 
<td width="100%" colspan=2 class=forumrow> 
<input type="submit" name="Submit" value="�� ��">
</td>
</tr>
</table>
</form>
<%
end sub
sub savenew()
if request.form("userclass")="" then
	Errmsg="<br><li>�������µĵȼ���š�"
	call dvbbs_error()
	exit sub
elseif not isnumeric(request.form("userclass")) then
	Errmsg="<br><li>�µĵȼ����ֻ�������֡�"
	call dvbbs_error()
	exit sub
end if
if request.form("minarticle")="" then
	Errmsg="<br><li>�������µĵȼ���Ҫ��������"
	call dvbbs_error()
	exit sub
elseif not isnumeric(request.form("minarticle")) then
	Errmsg="<br><li>�µĵȼ�������ֻ�������֡�"
	call dvbbs_error()
	exit sub
end if
if request.form("titlepic")="" then
	Errmsg="<br><li>�������µĵȼ�ͼƬ��"
	call dvbbs_error()
	exit sub
end if
if request.form("usertitle")="" then
	Errmsg="<br><li>�������µĵȼ����ơ�"
	call dvbbs_error()
	exit sub
end if
set rs = server.CreateObject ("Adodb.recordset")
sql="select * from usertitle where usertitle='"&request.form("usertitle")&"'"
rs.open sql,conn,1,3
if rs.eof and rs.bof then
rs.addnew
rs("userclass")=request.form("userclass")
rs("usertitle")=request.form("usertitle")
rs("minarticle")=request.form("minarticle")
rs("titlepic")=request.form("titlepic")
rs("usergroupid")=request.form("usergroupid")
rs.update
else
	Errmsg="<br><li>�õȼ������Ѿ����ڡ�"
	call dvbbs_error()
	exit sub
end if
rs.close
set rs=nothing
response.write "��ӳɹ����������������û������н��и��²�����"
end sub
sub del()
Server.ScriptTimeout=99999999
dim minarticle,minuserclass
if isnumeric(request("id")) then
set rs=conn.execute("select * from usertitle where usertitleid="&request("id"))
minarticle=rs("minarticle")
minuserclass=rs("usertitle")
if minarticle=-1 then
set rs=conn.execute("select top 1 * from usertitle order by minarticle")
else
set rs=conn.execute("select top 1 * from usertitle where MinArticle<"&minarticle&" and not Minarticle=-1 order by minarticle desc")
end if
if not (rs.eof and rs.bof) then
conn.execute("update [user] set userclass='"&rs("usertitle")&"',titlepic='"&rs("titlepic")&"' where userclass='"&minuserclass&"'")
end if
conn.execute("delete from usertitle where usertitleid="&request("id"))
response.write "ɾ���ɹ����������������û������н��и����û��ȼ�������"
set rs=nothing
end if
end sub
%>