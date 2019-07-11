<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!--#include file=md5.asp-->
<title><%=Forum_info(0)%>--管理页面</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta NAME=GENERATOR Content="Microsoft FrontPage 3.0" CHARSET=GB2312>
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0" marginheight="0" marginwidth="0">
<%
	dim str
	if not master or session("flag")="" then
		Errmsg=Errmsg+"<br>"+"<li>本页面为管理员专用，请<a href=admin_index.asp target=_top>登陆</a>后进入。<br><li>您没有管理本页面的权限。"
		call dvbbs_error()
	else
		call main()
		conn.close
		set conn=nothing
	end if

	sub main()
%>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th width="100%" class="tableHeaderText" colspan=2>论坛管理
</th>
</tr>
<tr>
<td class="forumrow" colspan=2>
<p><B>注意</B>：删除论坛同时将删除该论坛下所有帖子！删除分类同时删除下属论坛和其中帖子！ 操作时请完整填写表单信息。
</td>
</tr>
<tr>
<td class="forumrow">
<B>论坛操作选项</B></td>
<td class="forumrow"><a href="admin_board.asp">论坛管理首页</a> | <a href="admin_board.asp?action=addclass">新建论坛分类</a> | <a href="?action=orders">分类批量修改</a> | <a href="?action=boardorders">版面批量修改</a>
</td>
</tr>
</table>
<p></p>
<%
select case Request("action")
case "add"
	call add()
case "edit"
	call edit()
case "savenew"
	call savenew()
case "savedit"
	call savedit()
case "del"
	call del()
case "orders"
	call orders()
case "updatorders"
	call updateorders()
case "boardorders"
	call boardorders()
case "updatboardorders"
	call updateboardorders()
case "addclass"
	call addclass()
case "saveclass"
	call saveclass()
case "del1"
	call del1()
case "mode"
	call mode()
case "savemod"
	call savemod()
case "permission"
	call boardpermission()
case "editpermission"
	call editpermission()
case else
	call boardinfo()
end select
end sub

sub boardinfo()
dim rs_1,rs_2
dim sql_1,sql_2
	sql_2 = "select * from class order by orders"
	set rs_2=conn.execute(sql_2)
	do while not rs_2.Eof
%>
            <table width="95%" border="0" cellspacing="3" cellpadding="0" align=center>
              <tr>

                <td height="21"><B><%=rs_2("class")%></b>    <a href="admin_board.asp?action=add"><font color="<%=Forum_body(14)%>"><U>新增论坛</u></font></a> | <a href=admin_board.asp?action=del1&id=<%=rs_2("id")%> onclick="{if(confirm('删除将包括该分类下所有论坛的所有帖子，确定删除吗?')){return true;}return false;}"><font color="<%=Forum_body(14)%>"><U>删除分类</u></font></a></td>
              </tr>
            </table>
<%
		sql_1 = "select boardid,boardtype,readme,orders,lockboard from board where class="&rs_2("id")&" order by orders"
		set rs_1=conn.execute(sql_1)
		do while not rs_1.EOF 
%>
            <table width="95%" border="0" cellspacing="3" cellpadding="0" align=center>
              <tr> 
                <td height="18" width=50% ><li><%=rs_1("boardtype")%></font></td><td width=50% ><a href="admin_board.asp?action=edit&editid=<%=rs_1("boardid")%>"><font color="<%=Forum_body(14)%>"><U>编辑论坛</U></font></a><%if rs_1("lockboard")=2 then%> | <a href="admin_board.asp?action=mode&boardid=<%=rs_1("boardid")%>"><font color="<%=Forum_body(14)%>"><U>认证用户</U></font></a><%end if%> | <a href="admin_update.asp?action=updat&submit=更新论坛数据&boardid=<%=rs_1("boardid")%>" title="更新最后回复、帖子数、回复数"><font color="<%=Forum_body(14)%>"><U>更新帖子数据</U></font></a> | <a href="admin_update.asp?action=delboard&boardid=<%=rs_1("boardid")%>" onclick="{if(confirm('清空将包括该论坛所有帖子置于回收站，确定清空吗?')){return true;}return false;}"><font color="<%=Forum_body(14)%>"><U>清空</U></font></a> | <a href="admin_board.asp?action=del&editid=<%=rs_1("boardid")%>" onclick="{if(confirm('删除将包括该论坛所有帖子，确定删除吗?')){return true;}return false;}"><font color="<%=Forum_body(14)%>"><U>删除</U></a></td>
              </tr>
            </table>
<%
		rs_1.MoveNext
		loop
		rs_1.Close
%>
            <table width="95%" border="0" cellspacing="3" cellpadding="0" align=center>
              <tr>
                <td height="21" align=left><hr color=black height=1 width=70% size=1></td>
              </tr>
            </table>
<%
	rs_2.MoveNext 
	Loop
	rs_2.Close
set rs_1=nothing
set rs_2=nothing
end sub

sub add()
dim rs_c
set rs_c= server.CreateObject ("adodb.recordset")
sql = "select * from class"
rs_c.open sql,conn,1,1
	dim boardnum
	set rs = server.CreateObject ("Adodb.recordset")
	sql="select Max(boardid) from board"
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
	boardnum=1
	else
	boardnum=rs(0)+1
	end if
	if isnull(boardnum) then boardnum=1
	rs.close
%>
 <form action ="admin_board.asp?action=savenew" method=post>
<input type="hidden" name="newboardid" value=<%=boardnum%>>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th height=24 colspan=2><B>添加新论坛</th>
</tr>
<tr> 
<td width="52%" height=30 class="forumrow">论坛名称</td>
<td width="48%" class="forumrow"> 
<input type="text" name="boardtype" size="35">
</td>
</tr>
<tr> 
<td width="52%" height=24 class="forumrow">版面说明<BR>可以使用HTML代码</td>
<td width="48%" class="forumrow"> 
<textarea name="Readme" cols="40" rows="3"></textarea>
</td>
</tr>
<tr> 
<td width="52%" height=30 class="forumrow"><U>所属类别</U></td>
<td width="48%" class="forumrow"> 
<select name=class>
<% do while not rs_c.EOF%>
<option value="<%=rs_c("id")%>"><%=rs_c("class")%> 
<%
rs_c.MoveNext 
loop
rs_c.Close 
%>
</select>
</td>
</tr>
<tr> 
<td width="52%" height=30 class="forumrow"><U>使用设置模板</U><BR>相关模板中包含论坛颜色、设置、广告、图片<BR>等信息</td>
<td width="48%" class="forumrow"> 
<select name=sid>
<%
	sql = "select * from config order by active desc"
	rs_c.open sql,conn,1,1
	if rs_c.eof and rs_c.bof then
	response.write "<option value=>请先添加模板"
	else
	do while not rs_c.EOF
%>
<option value="<%=rs_c("id")%>"><%=rs_c("skinname")%> 
<%
	rs_c.MoveNext 
	loop
	end if
	rs_c.Close 
%>
</select>
</td>
</tr>
<tr> 
<td width="52%" height=30 class="forumrow"><U>论坛版主</U><BR>多斑竹添加请用|分隔，如：沙滩小子|wodeail</td>
<td width="48%" class="forumrow"> 
<input type="text" name="boardmaster" size="35">
</td>
</tr>
<tr> 
<td width="52%" height=30 class="forumrow"><U>首页显示论坛图片</U><BR>出现在首页论坛版面介绍左边<BR>请直接填写图片URL</td>
<td width="48%" class="forumrow">
<input type="text" name="indexIMG" size="35">
</td>
</tr><tr> 
<th height=24 colspan=2 align=left> ＝＝ <B>论坛选项</B></th>
</tr>
<tr> 
<td width="52%" height=30 class="forumrow"><U>论坛属性</U></td>
<td width="48%" class="forumrow"> 
<input type=radio name="lockboard" value="0" checked>开放 
<input type=radio name="lockboard" value="1">隐含 
<input type=radio name="lockboard" value="2">加密
<input type=radio name="lockboard" value="3">锁定
</td>
</tr>
<tr> 
<td width="52%" height=24 class="forumRow">&nbsp;</td>
<td width="48%" class="forumRow"> 
<input type="submit" name="Submit" value="添加论坛">
</td>
</tr>
</table>
</form>
<%
set rs=nothing
end sub

sub edit()
dim rs_c
set rs_c= server.CreateObject ("adodb.recordset")
sql = "select * from class"
rs_c.open sql,conn,1,1
set rs= server.CreateObject ("adodb.recordset")
sql = "select * from board where boardid="+CSTr(request("editid"))
rs.open sql,conn,1,1
%>        
 <form action ="admin_board.asp?action=savedit" method=post>       
<input type='hidden' name=editid value='<%=Request("editid")%>'>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th height=24 colspan=2>编辑论坛：<%=rs("boardtype")%></th>
</tr>
<tr> 
<td width="52%" height=30 class="forumrow">论坛名称</td>
<td width="48%" class="forumrow"> 
<input type="text" name="boardtype" size="35"  value = '<%=rs("boardtype")%>'>
</td>
</tr>
<tr> 
<td width="52%" height=24 class="forumrow">版面说明<BR>可以使用HTML代码</td>
<td width="48%" class="forumrow"> 
<textarea name="Readme" cols="40" rows="3"><%=rs("readme")%></textarea>
</td>
</tr>
<tr> 
<td width="52%" height=30 class="forumrow"><U>所属类别</U></td>
<td width="48%" class="forumrow"> 
<select name=class>
<% do while not rs_c.EOF%>
<option value=<%=rs_c("id")%> <% if cint(rs("class")) = rs_c("id") then%> selected <%end if%>><%=rs_c("class")%> 
<%
rs_c.MoveNext 
loop
rs_c.Close 
%>
</select>
</td>
</tr>
<tr> 
<td width="52%" height=30 class="forumrow"><U>使用设置模板</U><BR>相关模板中包含论坛颜色、设置、广告、图片<BR>等信息</td>
<td width="48%" class="forumrow"> 
<select name=sid>
<%
	sql = "select * from config order by active desc"
	rs_c.open sql,conn,1,1
	if rs_c.eof and rs_c.bof then
	response.write "<option value=>请先添加模板"
	else
	do while not rs_c.EOF
%>
<option value=<%=rs_c("id")%> <% if cint(rs("sid")) = rs_c("id") then%> selected <%end if%>><%=rs_c("skinname")%> 
<%
	rs_c.MoveNext 
	loop
	end if
	rs_c.Close 
%>
</select>
</td>
</tr>
<tr> 
<td width="52%" height=30 class="forumrow"><U>论坛版主</U><BR>多斑竹添加请用|分隔，如：沙滩小子|wodeail</td>
<td width="48%" class="forumrow"> 
<input type="text" name="boardmaster" size="35"  value='<%=rs("boardmaster")%>'>
</td>
</tr>
<tr> 
<td width="52%" height=30 class="forumrow"><U>首页显示论坛图片</U><BR>出现在首页论坛版面介绍左边<BR>请直接填写图片URL</td>
<td width="48%" class="forumrow">
<input type="text" name="indexIMG" size="35" value="<%=rs("indexIMG")%>">
</td>
</tr><tr> 
<th height=24 colspan=2 align=left> ＝＝ <B>论坛选项</B></th>
</tr>
<tr> 
<td width="52%" height=30 class="forumrow"><U>论坛属性</U></td>
<td width="48%" class="forumrow"> 
<input type=radio name="lockboard" value="0" <%if rs("lockboard")=0 then%>checked<%end if%>>开放 
<input type=radio name="lockboard" value="1" <%if rs("lockboard")=1 then%>checked<%end if%>>隐含
<input type=radio name="lockboard" value="2" <%if rs("lockboard")=2 then%>checked<%end if%>>加密
<input type=radio name="lockboard" value="3" <%if rs("lockboard")=3 then%>checked<%end if%>>锁定
</td>
</tr>
<tr> 
<td width="52%" height=24 class="forumrow">&nbsp;</td>
<td width="48%" class="forumrow"> 
<input type="submit" name="Submit" value="提交修改">
</td>
</tr>
</table>
</form>
<%
rs.close
set rs=nothing
set rs_c=nothing
end sub
sub mode()
dim boarduser
%>
 <form action ="admin_board.asp?action=savemod" method=post>
<table width="95%" class="tableBorder" cellspacing="1" cellpadding="1" align="center">
<tr bgcolor=<%=Forum_body(3)%>> 
<th width="52%" height=22>说明：</th>
<th width="48%">操作：</th>
</tr>
<tr> 
<td width="52%" height=22 class=forumrow><B>论坛名称</B></td>
<td width="48%" class=forumrow> 
<%
set rs= server.CreateObject ("adodb.recordset")
sql="select boardid,boardtype,boarduser from board where lockboard=2 and boardid="&request("boardid")
rs.open sql,conn,1,1
if rs.eof and rs.bof then
response.write "该版面并不存在或者该版面不是加密版面。"
else
response.write rs(1)
response.write "<input type=hidden value="&rs(0)&" name=boardid>"
boarduser=rs(2)
end if
rs.close
set rs=nothing
%>
</td>
</tr>
<tr> 
<td width="52%" class=forumrow><B>认证用户</B>：<br>
只有设定为认证论坛的论坛需要填写能够进入该版面的用户，每输入一个用户请确认用户名在论坛中存在，每个用户名用<B>回车</B>分开</font>
</td>
<td width="48%" class=forumrow> 
<textarea cols=35 rows=6 name="vipuser">
<%
if not isnull(boarduser) or boarduser<>"" then
	response.write replace(boarduser,",",chr(10))
end if
%>
</textarea>
</td>
</tr>
<tr> 
<td width="52%" height=22 class=forumrow>&nbsp;</td>
<td width="48%" class=forumrow> 
<input type="submit" name="Submit" value="设 定">
</td>
</tr>
</table>
</form>
<%
end sub

sub savemod()
dim boarduser
dim boarduser_1
dim userlen
dim updateinfo
response.write "<p>论坛设置成功！<br><br>"
if trim(request("vipuser"))<>"" then
	boarduser=request("vipuser")
	boarduser=split(boarduser,chr(13)&chr(10))
	for i = 0 to ubound(boarduser)
	if not (boarduser(i)="" or boarduser(i)=" ") then
		boarduser_1=""&boarduser_1&""&boarduser(i)&","
	end if
	next
	userlen=len(boarduser_1)
	if boarduser_1<>"" then
		boarduser=left(boarduser_1,userlen-1)
		response.write "<p>添加用户："&boarduser&"<br><br>"
		updateinfo=" boarduser='"&boarduser&"' "
	else
		response.write "<p><font color=red>你没有添加认证用户</font><br><br>"
	end if
end if
conn.execute("update board set "&updateinfo&" where boardid="&request("boardid"))
end sub

sub savenew()
dim Forum_info,Forum_setting,Forum_ubb,Forum_body,Forum_ads,Forum_user
dim Forum_pic,Forum_boardpic,Forum_TopicPic,Forum_statePic,Forum_copyright
	set rs = server.CreateObject ("adodb.recordset")
	if request("boardtype")="" then
		Errmsg=Errmsg+"<br>"+"<li>请输入论坛名称。"
		Founderr=true
	end if
	if request("class")="" then
		Errmsg=Errmsg+"<br>"+"<li>请选择论坛分类。"
		Founderr=true
	end if
	if request("boardmaster")="" then
		Errmsg=Errmsg+"<br>"+"<li>请输入版主姓名。"
		Founderr=true
	end if
	if request("readme")="" then
		Errmsg=Errmsg+"<br>"+"<li>请输入论坛说明。"
		Founderr=true
	end if
	if request("lockboard")="" then
		Errmsg=Errmsg+"<br>"+"<li>请选择论坛开放状态。"
		Founderr=true
	end if
	if founderr=true then
	response.write ""&Errmsg&""
	else
	dim boardid
	sql="select boardid from board where boardid="+cstr(request("newboardid"))
	rs.open sql,conn,1,1
	if not rs.eof and not rs.bof then
		response.write "您不能指定和别的论坛一样的序号。"
		exit sub
	else
		boardid=request("newboardid")
	end if
	rs.close
	sql = "select * from board"
	rs.Open sql,conn,1,3
	rs.AddNew
	rs("boardid") = Request("newboardid")
	rs("boardtype") = Request.Form ("boardtype")
	rs("class") = Request.Form  ("class")
	rs("boardmaster") = Request("boardmaster")
	rs("readme") = Request("readme")
	rs("lockboard") = Trim(Request("lockboard"))
	rs("indexIMG")=request.form("indexIMG")
	rs("lasttopicnum") = 0 
	rs("lastbbsnum") = 0 
	rs("lasttopicnum") = 0 
	rs("todaynum") = 0 
	rs("orders") = Request("newboardid")
	rs("LastPost")="$$"&Now()&"$$$$"
	rs("sid")=request("sid")

	rs.Update 
	rs.Close 
	call addmaster(Request("boardmaster"))
	response.write "<p>论坛添加成功！<br><br>"&str
	end if
	set rs=nothing
end sub

sub savedit()
	dim newboardid
	set rs = server.CreateObject ("adodb.recordset")
	sql = "select * from board where boardid="&request("editid")
	rs.Open sql,conn,1,3
	rs("boardtype") = Request.Form ("boardtype")
	rs("class") = Request.Form  ("class")
	rs("boardmaster") = Request("boardmaster")
	rs("readme") = Request("readme")
	rs("lockboard") = Trim(Request("lockboard"))
	rs("indexIMG")=request.form("indexIMG")
	rs("sid")=request.form("sid")
	rs.Update 
	rs.Close
	set rs=nothing
	call addmaster(Request("boardmaster"))
	response.write "<p>论坛修改成功！<br><br>"&str
end sub

sub del()
	set rs = server.CreateObject ("adodb.recordset")
	sql = "delete from board where boardid="+Cstr(Request("editid"))
	conn.execute(sql)
	sql = "delete from bbs1 where boardid="+cstr(Request("editid"))
	conn.execute(sql)
	set rs=nothing
	response.write "<p>论坛修改成功！"
end sub

sub del1()
	set rs = server.CreateObject ("adodb.recordset")
	sql = "delete from class where id="+Cstr(Request("id"))
	conn.execute(sql)
	sql = "delete from board where class="+Cstr(Request("id"))
	conn.execute(sql)
	sql="select boardid from board where class="+Cstr(Request("id"))
	rs.open sql,conn,1,1
	do while not rs.eof
	sql_1 = "delete from bbs1 where boardid="+cstr(rs("boardid"))
	conn.execute(sql_1)
	rs.movenext
	loop
	rs.close
	set rs=nothing
	response.write "<p>分类删除成功！"
end sub

sub orders()
%>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
	<tr> 
	<th height="22">论坛分类重新排序修改(请在相应论坛分类的排序表单内输入相应的排列序号)
	</th>
	</tr>
<form action=admin_board.asp?action=updatorders method=post>
	<tr>
	<td class="Forumrow">
<%
	set rs = server.CreateObject ("Adodb.recordset")
	sql="select * from class order by orders"
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
		response.write "还没有相应的论坛分类。"
	else
		do while not rs.eof
		response.write "<input type=text name=""classname"" size=25 value="""&rs("class")&""">"
		response.write "  <input type=text name=""OrderID"" size=4 value="""&rs("orders")&"""><input type=hidden name=""cID"" value="""&rs("id")&"""><BR>"
		rs.movenext
		loop
%>
<BR><input type=submit name=Submit value=修改>
<%
	end if
	rs.close
	set rs=nothing
%>
	</td>
	</tr></form>
</table>
<%
end sub

sub updateorders()
	dim cID,OrderID,ClassName
	for i=1 to request.form("cID").count
	cID=replace(request.form("cID")(i),"'","")
	OrderID=replace(request.form("OrderID")(i),"'","")
	ClassName=replace(request.form("ClassName")(i),"'","")
	conn.execute("update class set class='"&ClassName&"',Orders="&OrderID&" where id="&cID)
	next
	response.write "设置成功，请返回。"
end sub

sub boardorders()
%>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
	<tr> 
	<th height="22">论坛版面重新排序修改(请在相应论坛分类的排序表单内输入相应的排列序号)
	</th>
	</tr>
<form action=admin_board.asp?action=updatboardorders method=post>
	<tr>
	<td class=forumrow>
<%
	dim trs
	set rs = server.CreateObject ("Adodb.recordset")
	sql="select * from class order by orders"
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
		response.write "还没有相应的论坛分类。"
	else
		do while not rs.eof
		response.write "<BR>"&rs("orders")&".  "&rs("class")
		response.write "<BR><BR>"
		set trs=conn.execute("select boardid,boardtype,orders from board where class="&rs("id")&" order by orders")
			do while not trs.eof
			response.write "<input type=text name=""boardtype"" size=25 value="""&trs("boardtype")&""">"
			response.write "  <input type=text name=""OrderID"" size=4 value="""&trs("orders")&""">"
			response.write "<input type=hidden name=""bid"" value="""&trs("boardid")&"""><BR>"
			trs.movenext
			loop
			trs.close
		rs.movenext
		loop
%>
<BR><input type=submit name=Submit value=修改>
<%
	end if
	rs.close
	set rs=nothing
	set trs=nothing
%>
	</td>
	</tr></form>
	</table>
<%
end sub

sub updateboardorders()
	dim bID,OrderID,boardtype
	for i=1 to request.form("bID").count
	bID=replace(request.form("bID")(i),"'","")
	OrderID=replace(request.form("OrderID")(i),"'","")
	boardtype=replace(request.form("boardtype")(i),"'","")
	conn.execute("update board set boardtype='"&boardtype&"',Orders="&OrderID&" where boardid="&bID)
	next
	response.write "设置成功，请返回。"
end sub

sub addclass()
	dim classnum
	set rs = server.CreateObject ("Adodb.recordset")
	sql="select id from class"
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
	classnum=0
	else
	classnum=rs.recordcount
	end if
	rs.close
%>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
	<tr> 
	<th height="22">添加新的论坛分类
	</th>
	</tr>
<form action=admin_board.asp?action=saveclass method=post>
	<tr>
	<td class=forumrow>分类名：<input name=classname type=text size=25>  序号：
<input name=id type=text size=2 value=<%=classnum+1%>>   
<input type=submit name=Submit value=添加><BR><BR>
注意：请完整填写以下表单信息，注意不能和别的论坛分类有相同的排列序号。
	</td>
	</tr>
</form>
	</table>
<%
set rs=nothing
end sub

sub saveclass()
	set rs = server.CreateObject ("Adodb.recordset")
	if request("id")="" or request("classname")="" then
		response.write "您输入的序号和原来的相同，不必更新啦：）"
	else
	sql="select * from class where id="&cstr(request("id"))
	rs.open sql,conn,1,1
	if not rs.eof and not rs.bof then
		response.write "您输入的序号和其他分类序号相同，请重新输入。"
	else
		sql="insert into class(id,class) values("&request("id")&",'"&request("classname")&"')"
		conn.execute(sql)
		response.write "<p align=center>更新成功！</p>"
	end if
	end if
	set rs=nothing
end sub

sub delclass()

end sub

sub addmaster(s)
dim arr,i,rs,sql,pw
dim classname
set rs=conn.execute("select * from usergroups where usergroupid=3")
classname=rs("title")
randomize
pw=Cint(rnd*9000)+1000
arr=split(s,"|")
set rs=server.createobject("adodb.recordset")
for i=0 to Ubound(arr)
sql="select username,userpassword,userclass,UserGroupID,TitlePic from [user] where username='"& arr(i) &"'"
rs.open sql,conn,1,3
if rs.eof and rs.bof then
	rs.addnew
	rs("username")=arr(i)
	rs("userpassword")=md5(pw)
	rs("userclass")=classname
	rs("UserGroupID")=3
	rs.update
	str=str&"你添加了以下用户：<b>" &arr(i) &"</b> 密码：<b>"& pw &"</b><br><br>"
else
	if rs("UserGroupID")>3 then
		rs("userclass")=classname
		rs("UserGroupID")=3
		rs.update
	end if
end if
rs.close
next
set rs=nothing
end sub

sub boardpermission()
dim rs_1,rs_2
dim sql_1,sql_2
dim trs,ars
%>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
	<tr>
	<th height="21">编辑论坛权限</th>
	</tr>
	<tr>
	<td class=forumrow>您可以设置不同用户组在不同论坛内的权限，红色表示为该论坛该用户组使用的是用户定义属性
	</td>
	</tr>
	<tr>
	<td class=forumrow>
<%
	sql_2 = "select * from class order by orders"
	set rs_2=conn.execute(sql_2)
	do while not rs_2.Eof
%>
            <table width="100%" border="0" cellspacing="3" cellpadding="0">
              <tr>

                <td height="21"><B><%=rs_2("class")%></b></td>
              </tr>
            </table>
<%
		sql_1 = "select boardid,boardtype,readme,orders from board where class="&rs_2("id")&" order by orders"
		set rs_1=conn.execute(sql_1)
		do while not rs_1.EOF 
%>
            <table width="100%" border="0" cellspacing="3" cellpadding="0" style="font-size:9pt;line-height: 13pt">
              <tr> 
                <td height="18" width=50% valign=top><li><%=rs_1("boardtype")%></font></td>
				<td width=50% >
<%
		set trs=conn.execute("select title,usergroupid from usergroups order by usergroupid")
		do while not trs.eof
			set ars=conn.execute("select pid from BoardPermission where BoardID="&rs_1("boardid")&" and GroupID="&trs("UserGroupID"))
			if ars.eof and ars.bof then
			response.write "·" & trs("title") & "&nbsp;&nbsp;<a href=?action=editpermission&groupid="&trs("usergroupid")&"&reboardid="&rs_1("boardid")&">[编辑]</a><BR>"
			else
			response.write "·<font color=red>" & trs("title") & "</font>&nbsp;&nbsp;<a href=?action=editpermission&pid="&ars("pid")&"&reboardid="&rs_1("boardid")&">[编辑]</a><BR>"
			end if
		trs.movenext
		loop
%>
				</td>
              </tr>
            </table>
<%
		response.write "<hr color=black height=1 width=90% align=left size=1>"
		rs_1.MoveNext
		loop
		rs_1.Close
	rs_2.MoveNext 
	Loop
	rs_2.Close
%>  
	</td>
	</tr>
</table>
<%
set rs_1=nothing
set rs_2=nothing
set trs=nothing
set ars=nothing
end sub

sub editpermission()
if not isnumeric(request("groupid")) then
response.write "错误的参数！"
exit sub
end if
if request("groupaction")="yes" then
	dim GroupSetting
	GroupSetting=Request.Form("canview") & "," & Request.Form("canviewuserinfo") & "," & Request.Form("canviewpost") & "," & Request.Form("cannewpost") & "," & Request.Form("canreplymytopic") & "," & Request.Form("canreplytopic") & "," & Request.Form("canpostagree") & "," & Request.Form("canupload") & "," & Request.Form("canpostvote") & "," & Request.Form("canvote") & "," & Request.Form("caneditmytopic") & "," & Request.Form("candelmytopic") & "," & Request.Form("canmovemytopic") & "," & Request.Form("canclosemytopic") & "," & Request.Form("cansearch") & "," & Request.Form("canmailtopic") & "," & Request.Form("canmodify") & "," & Request.Form("cansmallpaper") & "," & Request.Form("candeltopic") & "," & Request.Form("canmovetopic") & "," & Request.Form("canclosetopic") & "," & Request.Form("cantoptopic") & "," & Request.Form("canawardtopic") & "," & Request.Form("canmodifytopic") & "," & Request.Form("canbesttopic") & "," & Request.Form("canAnnounce") & "," & Request.Form("canAdminAnnounce") & "," & Request.Form("canAdminPaper") & "," & Request.Form("canAdminUser") & "," & Request.Form("canDelUserTopic") & "," & Request.Form("canviewip") & "," & Request.Form("canadminip") & "," & Request.Form("cansendsms") & "," & Request.Form("Maxsendsms") & "," & Request.Form("Maxsmsbody") & "," & Request.Form("Maxsmsbox") & "," & Request.Form("canusetitle") & "," & Request.Form("canuseface") & "," & Request.Form("canusesign") & "," & Request.Form("canvieweven") & "," & Request.Form("canuploadnum") & "," & Request.Form("canviewbest") & "," & Request.Form("adminpermission") & "," & request.form("canaward") & "," & request.form("MaxUploadSize") & "," & request.form("canbatchtopic") & "," & request.form("smallpapermoney") & "," & request.form("postagreemoney")
	Set rs= Server.CreateObject("ADODB.Recordset")
	if request("isdefault")=1 then
	conn.execute("delete from BoardPermission where BoardID="&request("reBoardID")&" and GroupID="&request("GroupID"))
	else
	if request("pid")<>"" then
	sql="update BoardPermission set PSetting='"&GroupSetting&"' where pid="&request("pid")
	else
	sql="insert into BoardPermission (BoardID,GroupID,PSetting) values ("&request("reBoardID")&","&request("GroupID")&",'"&GroupSetting&"')"
	end if
	conn.execute(sql)
	end if
	set rs=nothing
	response.write "修改成功！返回<a href=?action=permission>论坛权限管理</a>"
else
Dim reGroupSetting,reBoardID,groupid
Dim Groupname,Boardname
if request("GroupID")<>"" then
set rs=conn.execute("select * from usergroups where usergroupid="&request("groupid"))
if rs.eof and rs.bof then
response.write "未找到该用户组！"
exit sub
end if
groupid=request("groupid")
reGroupSetting=split(rs("GroupSetting"),",")
reBoardID=request("reBoardID")
Groupname=rs("title")
else
set rs=conn.execute("select * from BoardPermission where pid="&request("pid"))
if rs.eof and rs.bof then
response.write "未找到该设置组！"
exit sub
end if
groupid=rs("groupid")
reGroupSetting=split(rs("PSetting"),",")
reBoardID=rs("boardid")
set rs=conn.execute("select title from UserGroups where usergroupid="&groupid)
groupname=rs("title")
end if
set rs=conn.execute("select boardtype from board where boardid="&reBoardID)
Boardname=rs("boardtype")
%>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<FORM METHOD=POST ACTION="?action=editpermission">
<input type=hidden name="groupid" value="<%=groupid%>">
<input type=hidden name="reBoardID" value="<%=reBoardID%>">
<input type=hidden name="pID" value="<%=request("pid")%>">
<tr> 
<th height="23" colspan="2" >编辑论坛用户组权限&nbsp;>> <%=boardname%>&nbsp;>> <%=groupname%></th>
</tr>
<tr> 
<td height="23" colspan="2" class=forumrow><input type=radio name="isdefault" value="1" <%if request("GroupID")<>"" then%>checked<%end if%>><B>使用用户组默认值</B> (注意: 这将删除任何之前所做的自定义设置)</td>
</tr>
<tr> 
<td height="23" colspan="2"  class=forumrow><input type=radio name="isdefault" value="0" <%if request("pID")<>"" then%>checked<%end if%>><B>使用自定义设置</B> </td>
</tr>
<tr> 
<th height="23" colspan="2" align=left>＝＝查看权限</th>
</tr>
<tr>
<td height="23" width="60%" class=forumrow>可以浏览论坛</td>
<td height="23" width="40%" class=forumrow>是<input name="canview" type=radio value="1" <%if reGroupSetting(0)=1 then%>checked<%end if%>>&nbsp;否<input name="canview" type=radio value="0" <%if reGroupSetting(0)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=forumrow>可以查看会员信息(包括其他会员的资料和会员列表)
</td>
<td height="23" width="40%" class=forumrow>是<input name="canviewuserinfo" type=radio value="1" <%if reGroupSetting(1)=1 then%>checked<%end if%>>&nbsp;否<input name="canviewuserinfo" type=radio value="0" <%if reGroupSetting(1)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=forumrow>可以查看其他人发布的主题
</td>
<td height="23" width="40%" class=forumrow>是<input name="canviewpost" type=radio value="1" <%if reGroupSetting(2)=1 then%>checked<%end if%>>&nbsp;否<input name="canviewpost" type=radio value="0" <%if reGroupSetting(2)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以浏览精华帖子
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canviewbest" type=radio value="1" <%if reGroupSetting(41)=1 then%>checked<%end if%>>&nbsp;否<input name="canviewbest" type=radio value="0" <%if reGroupSetting(41)=0 then%>checked<%end if%>></td>
</tr>
<tr> 
<th height="23" colspan="2" align=left>＝＝<b>发帖权限</b></th>
</tr>
<tr>
<td height="23" width="60%" class=forumrow>可以发布新主题</td>
<td height="23" width="40%" class=forumrow>是<input name="cannewpost" type=radio value="1" <%if reGroupSetting(3)=1 then%>checked<%end if%>>&nbsp;否<input name="cannewpost" type=radio value="0" <%if reGroupSetting(3)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=forumrow>可以回复自己的主题
</td>
<td height="23" width="40%" class=forumrow>是<input name="canreplymytopic" type=radio value="1" <%if reGroupSetting(4)=1 then%>checked<%end if%>>&nbsp;否<input name="canreplymytopic" type=radio value="0" <%if reGroupSetting(4)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=forumrow>可以回复其他人的主题
</td>
<td height="23" width="40%" class=forumrow>是<input name="canreplytopic" type=radio value="1" <%if reGroupSetting(5)=1 then%>checked<%end if%>>&nbsp;否<input name="canreplytopic" type=radio value="0" <%if reGroupSetting(5)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=forumrow>可以在论坛允许评分的时候参与评分(鲜花和鸡蛋)?
</td>
<td height="23" width="40%" class=forumrow>是<input name="canpostagree" type=radio value="1" <%if reGroupSetting(6)=1 then%>checked<%end if%>>&nbsp;否<input name="canpostagree" type=radio value="0" <%if reGroupSetting(6)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>参与评分所需金钱
</td>
<td height="23" width="40%" class=Forumrow><input name="postagreemoney" type=text size=4 value="<%=reGroupSetting(47)%>"></td>
</tr>
<tr>
<td height="23" width="60%" class=forumrow>可以上传附件
</td>
<td height="23" width="40%" class=forumrow>是<input name="canupload" type=radio value="1" <%if reGroupSetting(7)=1 then%>checked<%end if%>>&nbsp;否<input name="canupload" type=radio value="0" <%if reGroupSetting(7)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>最多上传文件个数
</td>
<td height="23" width="40%" class=Forumrow><input name="canuploadnum" type=text size=4 value="<%=reGroupSetting(40)%>"></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>上传文件大小限制
</td>
<td height="23" width="40%" class=Forumrow><input name="MaxUploadSize" type=text size=4 value="<%=reGroupSetting(44)%>"> KB</td>
</tr>
<tr>
<td height="23" width="60%" class=forumrow>可以发布新投票</td>
<td height="23" width="40%" class=forumrow>是<input name="canpostvote" type=radio value="1" <%if reGroupSetting(8)=1 then%>checked<%end if%>>&nbsp;否<input name="canpostvote" type=radio value="0" <%if reGroupSetting(8)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=forumrow>可以参与投票</td>
<td height="23" width="40%" class=forumrow>是<input name="canvote" type=radio value="1" <%if reGroupSetting(9)=1 then%>checked<%end if%>>&nbsp;否<input name="canvote" type=radio value="0" <%if reGroupSetting(9)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以发布小字报</td>
<td height="23" width="40%" class=Forumrow>是<input name="cansmallpaper" type=radio value="1"  <%if reGroupSetting(17)=1 then%>checked<%end if%>>&nbsp;否<input name="cansmallpaper" type=radio value="0" <%if reGroupSetting(17)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>发布小字报所需金钱</td>
<td height="23" width="40%" class=Forumrow><input name="smallpapermoney" type=text value="<%=reGroupSetting(46)%>" size=4></td>
</tr>
<tr> 
<th height="23" colspan="2" align=left>＝＝<b>帖子/主题编辑权限</b></th>
</tr>
<tr>
<td height="23" width="60%" class=forumrow>可以编辑自己的帖子
</td>
<td height="23" width="40%" class=forumrow>是<input name="caneditmytopic" type=radio value="1" <%if reGroupSetting(10)=1 then%>checked<%end if%>>&nbsp;否<input name="caneditmytopic" type=radio value="0" <%if reGroupSetting(10)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=forumrow>可以删除自己的帖子
</td>
<td height="23" width="40%" class=forumrow>是<input name="candelmytopic" type=radio value="1" <%if reGroupSetting(11)=1 then%>checked<%end if%>>&nbsp;否<input name="candelmytopic" type=radio value="0" <%if reGroupSetting(11)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=forumrow>可以移动自己的帖子到其他论坛
</td>
<td height="23" width="40%" class=forumrow>是<input name="canmovemytopic" type=radio value="1" <%if reGroupSetting(12)=1 then%>checked<%end if%>>&nbsp;否<input name="canmovemytopic" type=radio value="0" <%if reGroupSetting(12)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=forumrow>可以打开/关闭自己发布的主题
</td>
<td height="23" width="40%" class=forumrow>是<input name="canclosemytopic" type=radio value="1" <%if reGroupSetting(13)=1 then%>checked<%end if%>>&nbsp;否<input name="canclosemytopic" type=radio value="0" <%if reGroupSetting(13)=0 then%>checked<%end if%>></td>
</tr>
<tr> 
<th height="23" colspan="2" align=left>＝＝<b>其他权限</b></th>
</tr>
<tr>
<td height="23" width="60%" class=forumrow>可以搜索论坛
</td>
<td height="23" width="40%" class=forumrow>是<input name="cansearch" type=radio value="1" <%if reGroupSetting(14)=1 then%>checked<%end if%>>&nbsp;否<input name="cansearch" type=radio value="0" <%if reGroupSetting(14)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=forumrow>可以使用'发送本页给好友'功能
</td>
<td height="23" width="40%" class=forumrow>是<input name="canmailtopic" type=radio value="1" <%if reGroupSetting(15)=1 then%>checked<%end if%>>&nbsp;否<input name="canmailtopic" type=radio value="0" <%if reGroupSetting(15)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=forumrow>可以修改个人资料
</td>
<td height="23" width="40%" class=forumrow>是<input name="canmodify" type=radio value="1" <%if reGroupSetting(16)=1 then%>checked<%end if%>>&nbsp;否<input name="canmodify" type=radio value="0" <%if reGroupSetting(16)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>帖子中可以查看其它人头衔
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canusetitle" type=radio value="1" <%if reGroupSetting(36)=1 then%>checked<%end if%>>&nbsp;否<input name="canusetitle" type=radio value="0" <%if reGroupSetting(36)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>帖子中可以查看其它人头像
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canuseface" type=radio value="1"  <%if reGroupSetting(37)=1 then%>checked<%end if%>>&nbsp;否<input name="canuseface" type=radio value="0" <%if reGroupSetting(37)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>帖子中可以可以查看其它人签名
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canusesign" type=radio value="1"  <%if reGroupSetting(38)=1 then%>checked<%end if%>>&nbsp;否<input name="canusesign" type=radio value="0" <%if reGroupSetting(38)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以浏览论坛事件
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canvieweven" type=radio value="1"  <%if reGroupSetting(39)=1 then%>checked<%end if%>>&nbsp;否<input name="canvieweven" type=radio value="0" <%if reGroupSetting(39)=0 then%>checked<%end if%>></td>
</tr>
<tr> 
<th height="23" colspan="2" align=left>＝＝管理权限</th>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以删除其它人帖子
</td>
<td height="23" width="40%" class=Forumrow>是<input name="candeltopic" type=radio value="1" <%if reGroupSetting(18)=1 then%>checked<%end if%>>&nbsp;否<input name="candeltopic" type=radio value="0"  <%if reGroupSetting(18)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以移动其它人帖子
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canmovetopic" type=radio value="1" <%if reGroupSetting(19)=1 then%>checked<%end if%>>&nbsp;否<input name="canmovetopic" type=radio value="0"  <%if reGroupSetting(19)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以打开/关闭其它人帖子
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canclosetopic" type=radio value="1" <%if reGroupSetting(20)=1 then%>checked<%end if%>>&nbsp;否<input name="canclosetopic" type=radio value="0"  <%if reGroupSetting(20)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以固顶/解除固顶帖子
</td>
<td height="23" width="40%" class=Forumrow>是<input name="cantoptopic" type=radio value="1" <%if reGroupSetting(21)=1 then%>checked<%end if%>>&nbsp;否<input name="cantoptopic" type=radio value="0"  <%if reGroupSetting(21)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以奖励/惩罚发贴用户
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canawardtopic" type=radio value="1" <%if reGroupSetting(22)=1 then%>checked<%end if%>>&nbsp;否<input name="canawardtopic" type=radio value="0"  <%if reGroupSetting(22)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以奖励/惩罚用户
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canaward" type=radio value="1" <%if reGroupSetting(43)=1 then%>checked<%end if%>>&nbsp;否<input name="canaward" type=radio value="0"  <%if reGroupSetting(43)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以编辑其它人帖子
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canmodifytopic" type=radio value="1" <%if reGroupSetting(23)=1 then%>checked<%end if%>>&nbsp;否<input name="canmodifytopic" type=radio value="0" <%if reGroupSetting(23)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以加入/解除精华帖子
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canbesttopic" type=radio value="1" <%if reGroupSetting(24)=1 then%>checked<%end if%>>&nbsp;否<input name="canbesttopic" type=radio value="0"  <%if reGroupSetting(24)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以发布公告
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canAnnounce" type=radio value="1" <%if reGroupSetting(25)=1 then%>checked<%end if%>>&nbsp;否<input name="canAnnounce" type=radio value="0"  <%if reGroupSetting(25)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以管理公告
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canAdminAnnounce" type=radio value="1" <%if reGroupSetting(26)=1 then%>checked<%end if%>>&nbsp;否<input name="canAdminAnnounce" type=radio value="0"  <%if reGroupSetting(26)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以管理小字报
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canAdminPaper" type=radio value="1" <%if reGroupSetting(27)=1 then%>checked<%end if%>>&nbsp;否<input name="canAdminPaper" type=radio value="0"  <%if reGroupSetting(27)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以锁定/屏蔽/解除锁定用户
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canAdminUser" type=radio value="1" <%if reGroupSetting(28)=1 then%>checked<%end if%>>&nbsp;否<input name="canAdminUser" type=radio value="0"  <%if reGroupSetting(28)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以删除用户1－10天内所发帖子
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canDelUserTopic" type=radio value="1" <%if reGroupSetting(29)=1 then%>checked<%end if%>>&nbsp;否<input name="canDelUserTopic" type=radio value="0"  <%if reGroupSetting(29)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以查看来访IP及来源
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canviewip" type=radio value="1" <%if reGroupSetting(30)=1 then%>checked<%end if%>>&nbsp;否<input name="canviewip" type=radio value="0"  <%if reGroupSetting(30)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以限定IP来访
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canadminip" type=radio value="1" <%if reGroupSetting(31)=1 then%>checked<%end if%>>&nbsp;否<input name="canadminip" type=radio value="0"  <%if reGroupSetting(31)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以管理用户权限
</td>
<td height="23" width="40%" class=Forumrow>是<input name="adminpermission" type=radio value="1" <%if reGroupSetting(42)=1 then%>checked<%end if%>>&nbsp;否<input name="adminpermission" type=radio value="0"  <%if reGroupSetting(42)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以批量删除帖子（前台）
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canbatchtopic" type=radio value="1" <%if reGroupSetting(45)=1 then%>checked<%end if%>>&nbsp;否<input name="canbatchtopic" type=radio value="0"  <%if reGroupSetting(45)=0 then%>checked<%end if%>></td>
</tr>
<tr> 
<th height="23" colspan="2" align=left>＝＝短信权限</th>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以发送短信
</td>
<td height="23" width="40%" class=Forumrow>是<input name="cansendsms" type=radio value="1"  <%if reGroupSetting(32)=1 then%>checked<%end if%>>&nbsp;否<input name="cansendsms" type=radio value="0" <%if reGroupSetting(32)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>最多发送用户
</td>
<td height="23" width="40%" class=Forumrow><input name="Maxsendsms" size=5 type=text value="<%=reGroupSetting(33)%>"></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>短信内容大小限制
</td>
<td height="23" width="40%" class=Forumrow><input name="Maxsmsbody" size=5 type=text value="<%=reGroupSetting(34)%>"> byte</td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>信箱大小限制
</td>
<td height="23" width="40%" class=Forumrow><input name="Maxsmsbox" size=5 type=text value="<%=reGroupSetting(35)%>"> KB</td>
</tr>
<tr> 
<td height="23" width="60%" class=forumrow>
</td>
<td height="23" width="40%" class=forumrow><input type="submit" name="submit" value="提 交"></td>
</tr>
<input type=hidden value="yes" name="groupaction">
</FORM>
</table>
<%
end if
end sub
%>