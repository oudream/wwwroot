<!--#include file=conn.asp-->
<!-- #include file="inc/const.asp" -->
<title><%=Forum_info(0)%>--管理页面</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta NAME=GENERATOR Content="Microsoft FrontPage 3.0" CHARSET=GB2312>
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0" marginheight="0" marginwidth="0">
<%
	if not master or session("flag")="" then
		Errmsg=Errmsg+"<br>"+"<li>本页面为管理员专用，请<a href=admin_index.asp target=_top>登陆</a>后进入。<br><li>您没有管理本页面的权限。"
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
<B>相关组别ID信息</B>：<BR>
<%
	set rs=conn.execute("select * from usergroups order by usergroupid")
	do while not rs.eof
		response.write "用户组："&rs("title")&"，ID：<B>" & rs("usergroupid") & "</B><br>"
	rs.movenext
	loop
	rs.close
	set rs=nothing
%><BR>
相关用户组如果无对应等级名称，则注册用户自动按照文章升级<BR>
相关用户组的等级名称可以和用户组名不一样
</td>
<td width="50%" valign=top>
<B>在等级中设定用户组有什么用？</B><BR>
一般来说，只有注册用户拥有等级，所以在等级中一般都设定用户组ID为4对应注册用户组，如果设置成别的组ID，那么该用户在升级到这个等级的同时也将自动归入所设置的组<BR>
比如你新添加了一个用户组，并且给予了这个用户组某一些权限，那么你可以设置达到一定等级（帖子）的用户自动更新到这个用户组以使用这个用户组的权限。<BR>如果您想某个等级的用户不跟随帖子数而上升等级，那么就把最少发贴设置为<B>-1</B>，一般为特殊用户组需要这样的设置，设置某个级别最少发贴为<B>-1</B>后，该级别的用户将不能根据帖子增加而升级，别的用户也不能自动升级到该级别，只有在用户管理中方能更改其级别
</td>
</tr>
</table>
<form method="POST" action=admin_grade.asp?action=save>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th height="23" colspan="6" >用户等级设定</th>
</tr>
<tr> 
<td width="10%" class=forumHeaderBackgroundAlternate><b>等级<B></td>
<td width="25%" class=forumHeaderBackgroundAlternate><B>名称</B></td>
<td width="15%" class=forumHeaderBackgroundAlternate><B>最少发贴</B></td>
<td width="25%" class=forumHeaderBackgroundAlternate><B>图片</B></td>
<td width="15%" class=forumHeaderBackgroundAlternate><B>相关组ID</B></td>
<td width="10%" class=forumHeaderBackgroundAlternate><B>操作</B></td>
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
<td width="10%" class=Forumrow><a href="?action=del&id=<%=rs("usertitleid")%>">删除</a></td>
</tr>
<%
rs.movenext
loop
rs.close
set rs=nothing
%>
<tr> 
<td width="100%" colspan=6 class=Forumrow> 
<input type="submit" name="Submit" value="提 交">
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
	response.write "设置成功，请返回。"
	set rs=nothing
end sub

sub add()
%>
<form method="POST" action=admin_grade.asp?action=savenew>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th colspan="2">添加新的用户等级</th>
</tr>
<tr> 
<td width="40%" class=forumrow><b>等级<B></td>
<td width="60%" class=forumrow><input size=30 name="userclass" type=text></td>
</tr>
<tr>
<td width="40%" class=forumrow><B>所属用户组</B></td>
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
<td width="40%" class=forumrow><B>名称</B></td>
<td width="60%" class=forumrow><input size=30 name="usertitle" type=text></td>
</tr>
<tr>
<td width="40%" class=forumrow><B>最少发贴</B><BR>如果该等级是荣誉称号或者管理身份，这里可以填写-1，表示不跟随帖子增长而升级</td>
<td width="60%" class=forumrow><input size=30 name="minarticle" type=text></td>
</tr>
<tr>
<td width="40%" class=forumrow><B>图片</B></td>
<td width="60%" class=forumrow><input size=30 name="titlepic" type=text></td>
</tr>
<tr> 
<td width="100%" colspan=2 class=forumrow> 
<input type="submit" name="Submit" value="提 交">
</td>
</tr>
</table>
</form>
<%
end sub
sub savenew()
if request.form("userclass")="" then
	Errmsg="<br><li>请输入新的等级序号。"
	call dvbbs_error()
	exit sub
elseif not isnumeric(request.form("userclass")) then
	Errmsg="<br><li>新的等级序号只能是数字。"
	call dvbbs_error()
	exit sub
end if
if request.form("minarticle")="" then
	Errmsg="<br><li>请输入新的等级需要文章数。"
	call dvbbs_error()
	exit sub
elseif not isnumeric(request.form("minarticle")) then
	Errmsg="<br><li>新的等级文章数只能是数字。"
	call dvbbs_error()
	exit sub
end if
if request.form("titlepic")="" then
	Errmsg="<br><li>请输入新的等级图片。"
	call dvbbs_error()
	exit sub
end if
if request.form("usertitle")="" then
	Errmsg="<br><li>请输入新的等级名称。"
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
	Errmsg="<br><li>该等级名称已经存在。"
	call dvbbs_error()
	exit sub
end if
rs.close
set rs=nothing
response.write "添加成功！建议您到更新用户数据中进行更新操作！"
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
response.write "删除成功！建议您到更新用户数据中进行更新用户等级操作。"
set rs=nothing
end if
end sub
%>