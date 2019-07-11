<!--#include file=conn.asp-->
<!-- #include file="inc/const.asp" -->
<title><%=Forum_info(0)%>--管理页面</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta NAME=GENERATOR Content="Microsoft FrontPage 3.0" CHARSET=GB2312>
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0" marginheight="0" marginwidth="0">
<%
	Server.ScriptTimeOut=9999999
	if not master or session("flag")="" then
		Errmsg=Errmsg+"<br>"+"<li>本页面为管理员专用，请<a href=admin_index.asp target=_top>登陆</a>后进入。<br><li>您没有管理本页面的权限。"
		call dvbbs_error()
	else
		if request("action")="Nowused" then
		call nowused()
		elseif request("action")="update" then
		call update()
		elseif request("action")="del" then
		call del()
		elseif request("action")="CreatTable" then
		call creattable()
		elseif request("action")="search" then
		call search()
		elseif request("action")="update2" then
		call update2()
		elseif request("action")="update3" then
		call update3()
		else
		call main()
		end if
		conn.close
		set conn=nothing
	end if

sub main()
%>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<td height="23" colspan="2" class=Forumrow><B>说明</B>：<BR>您可以选择下列其中之一种模式进行帖子数据在不同表之间的转换。</td>
</tr>
<tr> 
<th height="23" colspan="2" align=left>模式一：搜索要转移的帖子</th>
</tr>
<FORM METHOD=POST ACTION="?action=search">
<tr> 
<td height="23" width="20%" class=Forumrow><B>搜索条件</B></td>
<td height="23" width="80%" class=Forumrow><input type="text" name="keyword">&nbsp;
<select name="tablename">
<%
for i=0 to ubound(AllPostTable)
	response.write "<option value="""&AllPostTable(i)&""">"&AllPostTableName(i)& "--" &AllPostTable(i)&"</option>"
next
%>
</select>
&nbsp;<input type="checkbox" name="username" value="yes" checked>用户&nbsp;<input type="checkbox" name="topic" value="yes">主题&nbsp;<input type="submit" name="submit" value="搜索"></td>
</tr>
</FORM>
<tr> 
<td height="23" colspan="2" class=Forumrow><B>注意</B>：这里仅搜索所在表的主题和发表用户数据，搜索后对搜索结果进行操作</td>
</tr>
<tr> 
<th height="23" colspan="2" align=left>模式二：在不同表转移数据</th>
</tr>
<FORM METHOD=POST ACTION="?action=update2">
<tr> 
<td height="23" width="100%" class=Forumrow colspan="2">&nbsp;
<select name="OutTablename">
<%
for i=0 to ubound(AllPostTable)
	response.write "<option value="""&AllPostTable(i)&""">"&AllPostTableName(i)& "--" &AllPostTable(i)&"</option>"
next
%>
</select>
 <input type="checkbox" name="top1000" value="yes" checked>前 <input type="checkbox" name="end1000" value="yes">后 <input type=text name="selnum" value="100" size=3>条 记录转移到
<select name="InTablename">
<%
for i=0 to ubound(AllPostTable)
	response.write "<option value="""&AllPostTable(i)&""">"&AllPostTableName(i)& "--" &AllPostTable(i)&"</option>"
next
%>
</select>
&nbsp;<input type="submit" name="submit" value="提交">
</td>
</tr>
</FORM>
<tr> 
<td height="23" colspan="2" class=Forumrow><B>注意</B>：最前N条记录指数据库中最早发表的帖子（如果平均每个帖子有5个回复，那么100个主题在这里的更新量将是500条记录），这通常要花很长的时间，更新的速度取决于您的服务器性能以及更新数据的多少。执行本步骤将消耗大量的服务器资源，建议您在访问人数较少的时候或者本地进行更新操作。</td>
</tr>
</table>
<%
end sub

sub nowused()
%>
<form method="POST" action="?action=update">
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<td height="23" colspan="5" class=Forumrow><B>说明</B>：<BR>下列数据表中选中的为当前论坛所使用来保存帖子数据的表，一般情况下每个表中的数据越少论坛帖子显示速度越快，当您下列单个帖子数据表中的数据有超过几万的帖子时不妨新添一个数据表来保存帖子数据（SQL版本用户建议每个表数据达到20万以后进行添加表操作），您会发现论坛速度快很多很多。<BR>您也可以将当前所使用的数据表在下列数据表中切换，当前所使用的帖子数据表即当前论坛用户发贴时默认的保存帖子数据表</td>
</tr>
<tr> 
<th height="23" colspan="5">当前数据表设定</th>
</tr>
<tr> 
<td width="20%" class=forumHeaderBackgroundAlternate><b>表名<B></td>
<td width="20%" class=forumHeaderBackgroundAlternate><B>说明</B></td>
<td width="20%" class=forumHeaderBackgroundAlternate><B>当前帖数</B></td>
<td width="20%" class=forumHeaderBackgroundAlternate><B>当前默认</B></td>
<td width="20%" class=forumHeaderBackgroundAlternate><B>删除</B></td>
</tr>
<%
for i=0 to ubound(AllPostTable)
%>
<tr> 
<td width="20%" class=Forumrow><%=AllPostTable(i)%></td>
<td width="20%" class=Forumrow><%=AllPostTableName(i)%></td>
<td width="20%" class=Forumrow>
<%
set rs=conn.execute("select count(*) from "&AllPostTable(i)&"")
response.write rs(0)
%>
</td>
<td width="20%" class=Forumrow><input value="<%=AllPostTable(i)%>" name="TableName" type="radio" <%if NowUseBBS=AllPostTable(i) then%>checked<%end if%>></td>
<td width="20%" class=Forumrow><a href="?action=del&tablename=<%=AllPostTable(i)%>"  onclick="{if(confirm('删除将包括该数据表所有帖子，本操作所删除的内容将不可恢复，确定删除吗?')){return true;}return false;}">删除</a></td>
</tr>
<%
next
%>
<tr> 
<td width="100%" colspan=5 class=Forumrow> 
<input type="submit" name="Submit" value="提 交">
</td>
</tr>
</form>
<FORM METHOD=POST ACTION="?action=CreatTable">
<tr> 
<th height="23" colspan="5">添加数据表</th>
</tr>
<tr> 
<td width="20%" class=Forumrow>添加的表名</td>
<td width="80%" class=Forumrow colspan=4><input type=text name="tablename" value="bbs<%=ubound(AllPostTable)+2%>">&nbsp;只能用bbs+数字表示，如bbs5，最后的数字最多不能超过9</td>
</tr>
<tr> 
<td width="20%" class=Forumrow>添加表的说明</td>
<td width="80%" class=Forumrow colspan=4><input type=text name="tablereadme">&nbsp;简单描述该表的用途，在搜索帖子和其他相关操作部分显示</td>
</tr>
<tr> 
<td width="100%" colspan=5 class=Forumrow> 
<input type="submit" name="Submit" value="提 交">
</td>
</tr>
</FORM>
</table>
<%
end sub

sub update()
	conn.execute("update config set NowUseBBS='"&request.form("TableName")&"'")
	response.write "更新成功！"
end sub

sub del()
	dim nAllPostTable,nAllPostTableName
	if request("tablename")=nowusebbs then
		errmsg="<br><li>当前正在使用的表不能删除。"
		call dvbbs_error()
		exit sub
	end if
	For i=0 to ubound(AllPostTable)
		if trim(request("tablename"))=trim(allposttable(i)) then
		nAllPostTableName=AllPostTableName(i)
		exit for
		end if
	Next
	set rs=conn.execute("select * from config where active=1")
	nAllPostTable=replace(rs("AllPostTable"),"|"&request("tablename"),"")
	nAllPostTableName=replace(rs("AllPostTableName"),"|"&nAllPostTableName,"")
	conn.execute("update config set AllPostTable='"&nAllPostTable&"',AllPostTableName='"&nAllPostTableName&"'")
	conn.execute("drop table "&request("tablename")&"")
	response.write "删除成功！"
end sub

sub CreatTable()
if request.form("tablename")="" then
	errmsg="<br><li>请输入表名。"
	call dvbbs_error()
	exit sub
elseif len(request.form("tablename"))<>4 then
	errmsg="<br><li>输入的表名不合法。"
	call dvbbs_error()
	exit sub
elseif not isnumeric(right(request.form("tablename"),1)) then
	errmsg="<br><li>输入的表名不合法。"
	call dvbbs_error()
	exit sub
elseif cint(right(request.form("tablename"),1))>9 or cint(right(request.form("tablename"),1))<0 then
	errmsg="<br><li>输入的表名不合法。"
	call dvbbs_error()
	exit sub
end if
if request.form("tablereadme")="" then
	errmsg="<br><li>请输入表的说明。"
	call dvbbs_error()
	exit sub
end if
for i=0 to ubound(AllPostTable)
	if AllPostTable(i)=request.form("tablename") then
		errmsg="<br><li>您输入的表名已经存在，请重新输入。"
		call dvbbs_error()
		exit sub
	end if
next

Dim NewAllPostTable,NewAllPostTableName
set rs=conn.execute("select allposttable,allposttablename from config where active=1")
NewAllPostTable=rs(0) & "|" & request.form("tablename")
NewAllPostTableName=rs(1) & "|" & request.form("tablereadme")

'Set conn = Server.CreateObject("ADODB.Connection")
'connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("dvbbs5.mdb")
'conn.open connstr
'建立新表
sql="CREATE TABLE "&request.form("tablename")&" (AnnounceID int IDENTITY (1, 1) NOT NULL CONSTRAINT PrimaryKey PRIMARY KEY,"&_
		"ParentID int default 0,"&_
		"BoardID int default 0,"&_
		"UserName varchar(50),"&_
		"PostUserID int default 0,"&_
		"Topic varchar(250),"&_
		"Body text,"&_
		"DateAndTime datetime default Now(),"&_
		"length int Default 0,"&_
		"RootID int Default 0,"&_
		"layer int Default 0,"&_
		"orders int Default 0,"&_
		"isbest int Default 0,"&_
		"ip varchar(20) NULL,"&_
		"Expression varchar(20) NULL,"&_
		"locktopic int Default 0,"&_
		"signflag int Default 0,"&_
		"emailflag int Default 0,"&_
		"isagree varchar(50) NULL)"
conn.execute(sql)

'添加索引
conn.execute("create index dispbbs on "&request.form("tablename")&" (postuserid,boardid,rootid,locktopic,announceid)")
conn.execute("create index save_2 on "&request.form("tablename")&" (rootid,orders)")
conn.execute("create index search on "&request.form("tablename")&" (boardid,dateandtime,topic,announceid)")

conn.execute("update config set AllPostTable='"&NewAllPostTable&"',AllPostTableName='"&NewAllPostTableName&"'")
response.write "添加表成功，请返回。"
end sub

sub update2()
dim trs
dim ForNum,TopNum
Dim orderby
if request.form("outtablename")=request.form("intablename") then
	Errmsg="<br><li>不能在相同数据表内转移数据。"
	call dvbbs_error()
	exit sub
end if
if (not isnumeric(request.form("selnum"))) or request.form("selnum")="" then
	Errmsg="<br><li>请填写正确的更新数量。"
	call dvbbs_error()
	exit sub
end if
if request.form("top1000")="yes" then
orderby=""
else
orderby=" desc"
end if
TopNum=Clng(request.form("selnum"))
if TopNum>100 then
	ForNum=int(TopNum/100)+1
	TopNum=100
else
	ForNum=1
end if

for i=1 to ForNum
set rs=conn.execute("select top "&TopNum&" topicid from topic where PostTable='"&request.form("outtablename")&"' order by topicid "&orderby&"")
if rs.eof and rs.bof then
	Errmsg="<br><li>您所选择导出的数据表已经没有任何内容"
	call dvbbs_error()
	exit sub
else
	do while not rs.eof
	'读取导出帖子数据表
	set trs=conn.execute("select * from "&request.form("outtablename")&" where rootid="&rs("topicid"))
	if not (trs.eof and trs.bof) then
	do while not trs.eof
	'插入导入帖子数据表
	conn.execute("insert into "&request("intablename")&"(Boardid,ParentID,username,topic,body,DateAndTime,length,rootid,layer,orders,ip,Expression,locktopic,signflag,emailflag,isbest,PostUserID) values ("&trs("boardid")&","&trs("parentid")&",'"&checkstr(trs("username"))&"','"&checkstr(trs("topic"))&"','"&checkstr(trs("body"))&"','"&trs("dateandtime")&"',"&trs("length")&","&trs("rootid")&","&trs("layer")&","&trs("orders")&",'"&trs("ip")&"','"&trs("Expression")&"',"&trs("locktopic")&","&trs("signflag")&","&trs("emailflag")&","&trs("isbest")&","&trs("postuserid")&")")
	'更新主题指定帖子表
	conn.execute("update topic set PostTable='"&request.form("inTableName")&"' where TopicID="&rs("topicid"))
	'删除导出帖子数据表对应数据
	conn.execute("delete from "&request.form("outTableName")&" where AnnounceID="&trs("Announceid"))
	trs.movenext
	loop
	end if
	rs.movenext
	loop
end if
next
set trs=nothing
set rs=nothing
response.write "更新成功！"
end sub

sub search()
dim keyword
dim totalrec
dim n
dim currentpage,page_count,Pcount
currentPage=request("page")
if currentpage="" or not isInteger(currentpage) then
	currentpage=1
else
	currentpage=clng(currentpage)
end if
if request("keyword")="" then
	Errmsg="<br><li>请输入您要查询的关键字。"
	call dvbbs_error()
	exit sub
else
	keyword=replace(request("keyword"),"'","")
end if
if request("username")="yes" then
sql="select * from topic where PostTable='"&request("tablename")&"' and postusername='"&keyword&"' order by LastPostTime desc"
elseif request("topic")="yes" then
sql="select * from topic where PostTable='"&request("tablename")&"' and title like '%"&keyword&"%' order by LastPostTime desc"
else
	Errmsg="<br><li>请选择您查询的方式。"
	call dvbbs_error()
	exit sub
end if
%>
<form method="POST" action="?action=update3">
<input type=hidden name="topic" value="<%=request("topic")%>">
<input type=hidden name="username" value="<%=request("username")%>">
<input type=hidden name="keyword" value="<%=keyword%>">
<input type=hidden name="tablename" value="<%=request("tablename")%>">
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<td height="23" colspan="6" class=Forumrow><B>说明</B>：<BR>您可以对下列的搜索结果进行转移数据表的操作，不能在相同表内进行转换操作。</td>
</tr>
<tr> 
<th height="23" colspan="6">搜索<%=request("tablename")%>结果</th>
</tr>
<tr> 
<td width="6%" class=forumHeaderBackgroundAlternate align=center><b>状态<B></td>
<td width="45%" class=forumHeaderBackgroundAlternate align=center><B>标题</B></td>
<td width="15%" class=forumHeaderBackgroundAlternate align=center><B>作者</B></td>
<td width="6%" class=forumHeaderBackgroundAlternate align=center><B>回复</B></td>
<td width="22%" class=forumHeaderBackgroundAlternate align=center><B>时间</B></td>
<td width="6%" class=forumHeaderBackgroundAlternate align=center><B>操作</B></td>
</tr>
<%
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
if rs.bof and rs.eof then
	response.write "<tr> <td class=Forumrow colspan=6 height=25>没有搜索到相关内容。</td></tr>"
else
	rs.PageSize = Forum_Setting(11)
	rs.AbsolutePage=currentpage
	page_count=0
	totalrec=rs.recordcount
	while (not rs.eof) and (not page_count = rs.PageSize)
%>
<tr> 
<td width="6%" class=Forumrow align=center>
<%
if rs("locktopic")=1 then
	response.write "锁定"
elseif rs("isvote")=1 then
	response.write "投票"
elseif rs("isbest")=1 then
	response.write "精华"
else
	response.write "正常"
end if
%>
</td>
<td width="45%" class=Forumrow><%=htmlencode(rs("title"))%></td>
<td width="15%" class=Forumrow align=center><a href="admin_user.asp?action=modify&userid=<%=rs("postuserid")%>"><%=htmlencode(rs("postusername"))%></a></td>
<td width="6%" class=Forumrow align=center><%=rs("child")%></td>
<td width="22%" class=Forumrow><%=rs("dateandtime")%></td>
<td width="6%" class=Forumrow align=center><input type="checkbox" name="topicid" value="<%=rs("topicid")%>"></td>
</tr>
<%
  	page_count = page_count + 1
	rs.movenext
	wend
	dim endpage
	Pcount=rs.PageCount
	response.write "<tr><td valign=middle nowrap colspan=2 class=forumrow height=25>&nbsp;&nbsp;分页： "

	if currentpage > 4 then
	response.write "<a href=""?page=1&action=search&keyword="&keyword&"&topic="&request("topic")&"&username="&request("username")&"&tablename="&request("tablename")&""">[1]</a> ..."
	end if
	if Pcount>currentpage+3 then
	endpage=currentpage+3
	else
	endpage=Pcount
	end if
	for i=currentpage-3 to endpage
	if not i<1 then
		if i = clng(currentpage) then
        response.write " <font color="&Forum_body(8)&">["&i&"]</font>"
		else
        response.write " <a href=""?page="&i&"&action=search&keyword="&keyword&"&topic="&request("topic")&"&username="&request("username")&"&tablename="&request("tablename")&""">["&i&"]</a>"
		end if
	end if
	next
	if currentpage+3 < Pcount then 
	response.write "... <a href=""?page="&Pcount&"&action=search&keyword="&keyword&"&topic="&request("topic")&"&username="&request("username")&"&tablename="&request("tablename")&""">["&Pcount&"]</a>"
	end if
	response.write "</td>"
	response.write "<td colspan=3 class=forumrow>所有结果<input type=checkbox name=allsearch value=yes>"
	response.write "&nbsp;<select name=toTablename>"

	for i=0 to ubound(AllPostTable)
		response.write "<option value="""&AllPostTable(i)&""">"&AllPostTableName(i)& "--" &AllPostTable(i)&"</option>"
	next

	response.write "</select>&nbsp;<input type=submit name=submit value=转换>"
	response.write "</td>"
	response.write "<td class=forumrow align=center><input type=checkbox name=chkall value=on onclick=""CheckAll(this.form)"">"
	response.write "</td></tr>"
end if
rs.close
set rs=nothing
response.write "</table></form><BR><BR>"
end sub

sub update3()
dim keyword,trs
if request.form("tablename")=request.form("totablename") then
	Errmsg="<br><li>不能在相同数据表内进行数据转换。"
	call dvbbs_error()
	exit sub
end if
if request.form("allsearch")="yes" then
	if request("keyword")="" then
		Errmsg="<br><li>请输入您要查询的关键字。"
		call dvbbs_error()
		exit sub
	else
		keyword=replace(request("keyword"),"'","")
	end if
	if request("username")="yes" then
		sql="select topicid from topic where PostTable='"&request("tablename")&"' and postusername='"&keyword&"' order by LastPostTime desc"
	elseif request("topic")="yes" then
		sql="select topicid from topic where PostTable='"&request("tablename")&"' and title like '%"&keyword&"%' order by LastPostTime desc"
	else
		Errmsg="<br><li>请选择您查询的方式。"
		call dvbbs_error()
		exit sub
	end if
else
	if request.form("topicid")="" then
		Errmsg="<br><li>请选择要转移的帖子。"
		call dvbbs_error()
		exit sub
	end if
	sql="select topicid from topic where PostTable='"&request("tablename")&"' and TopicID in ("&request.form("TopicID")&")"
end if

set rs=conn.execute(sql)
if rs.eof and rs.bof then
	Errmsg="<br><li>没有任何记录可转换。"
	call dvbbs_error()
	exit sub
else
	do while not rs.eof
	'取出原表数据
	set trs=conn.execute("select * from "&request("tablename")&" where rootid="&rs("topicid"))
	if not (trs.eof and trs.bof) then
	'插入新表
	do while not trs.eof
	conn.execute("insert into "&request("totablename")&"(Boardid,ParentID,username,topic,body,DateAndTime,length,rootid,layer,orders,ip,Expression,locktopic,signflag,emailflag,isbest,PostUserID) values ("&trs("boardid")&","&trs("parentid")&",'"&checkstr(trs("username"))&"','"&checkstr(trs("topic"))&"','"&checkstr(trs("body"))&"','"&trs("dateandtime")&"',"&trs("length")&","&trs("rootid")&","&trs("layer")&","&trs("orders")&",'"&trs("ip")&"','"&trs("Expression")&"',"&trs("locktopic")&","&trs("signflag")&","&trs("emailflag")&","&trs("isbest")&","&trs("postuserid")&")")
	trs.movenext
	loop
	'更新该主题表名
	conn.execute("update topic set PostTable='"&request("totablename")&"' where topicid="&rs("topicid"))
	'删除原表该帖子数据
	conn.execute("delete from "&request("tablename")&" where rootid="&rs("topicid"))
	end if
	rs.movenext
	loop
end if
set trs=nothing
set rs=nothing
response.write "更新成功！"
end sub
%>
<script language="javascript">
function CheckAll(form)  {
  for (var i=0;i<form.elements.length;i++)    {
    var e = form.elements[i];
    if (e.name != 'chkall')       e.checked = form.chkall.checked; 
   }
  }
</script>