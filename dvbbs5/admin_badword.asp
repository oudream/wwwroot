<!--#include file=conn.asp-->
<!-- #include file="inc/const.asp" -->
<title><%=Forum_info(0)%>--管理页面</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta NAME=GENERATOR Content="Microsoft FrontPage 3.0" CHARSET=GB2312>
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0" marginheight="0" marginwidth="0" bgcolor="#DDEEFF">
<%
	if not master or session("flag")="" then
		Errmsg=Errmsg+"<br>"+"<li>本页面为管理员专用，请<a href=admin_index.asp target=_top>登陆</a>后进入。<br><li>您没有管理本页面的权限。"
		call dvbbs_error()
	else
		call main()
		conn.close
		set conn=nothing
	end if

	sub main()
	dim sel

if request("action") = "savebadword" then
call savebadword()
else

%>

<form action="admin_badword.asp?action=savebadword" method=post>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center>
<tr> 
<td height="23" colspan="2" ><B>说明</B>：<BR>1、复选框中选择的为当前的使用设置模板，点击可查看该模板设置，点击别的模板直接查看该模板并修改设置。您可以将您下面的设置保存在多个论坛风格模板中<BR>2、您也可以将下面设定的信息保存并应用到具体的分论坛设置中，可多选<BR>3、如果您想在一个版面引用别的版面或模板的配置，只要点击该版面或模板名称，保存的时候选择要保存到的版面名称或模板名称即可。
</td>
</tr>
</table>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th width="100%" colspan=2 align=left>当前使用主模板（可将设置保存到下列模板中）</th>
</tr>
<tr>
<td colspan=2 class=forumrow>
<%
set rs = server.CreateObject ("adodb.recordset")
sql="select * from config"
rs.open sql,conn,1,1
do while not rs.eof
if request("skinid")="" then
	if request("boardid")="" then
	if rs("active")=1 then
	sel="checked"
	else
	sel=""
	end if
	else
	sel=""
	end if
else
	if rs("id")=cint(request("skinid")) then
	sel="checked"
	else
	sel=""
	end if
end if
response.write "<input type=checkbox name=skinid value="&rs("id")&" "&sel&"><a href=admin_badword.asp?skinid="&rs("id")&"><font color="&Forum_body(7)&">"&rs("skinname")&"</font></a>&nbsp;"
rs.movenext
loop
rs.close
set rs=nothing
%></font>
</td></tr>
<tr> 
<td width="100%" colspan=2 class=forumrow><B>当前使用该模板的分论坛</B><BR>以下分论坛将使用本模板设置，如需修改论坛使用模板，可到论坛管理中重新设置<BR>
<select size=1>
<%
set rs = server.CreateObject ("adodb.recordset")
sql="select * from board where sid="&skinid
rs.open sql,conn,1,1
if rs.eof and rs.bof then
response.write "<option>没有论坛使用该模板</option>"
else
do while not rs.eof
response.write "<option>" & rs("boardtype") & "</option>"
rs.movenext
loop
end if
rs.close
set rs=nothing
%></select>
</td></tr>
<%if request("reaction")="badword" then%>
<tr>
<th colspan=2 align=left height=23>帖子过滤字符</th>
</tr>
<tr>
<td class=forumrow width="20%">说明：</td>
<td class=forumrow width="80%">帖子过滤字符将过滤帖子内容中包含以下字符的内容，请您将要过滤的字符串添入，如果有多个字符串，请用“|”分隔开，例如：妈妈的|我靠|fuck</td>
</tr>
<tr>
<td class=forumrow width="20%">请输入过滤字符</td>
<td class=forumrow width="80%"><input type="text" name="badwords" value="<%=badwords%>" size="80"></td>
</tr>
<%elseif request("reaction")="splitreg" then%>
<tr>
<th colspan=2 align=left height=23>注册过滤字符</th>
</tr>
<tr>
<td class=forumrow width="20%">说明：</td>
<td class=forumrow width="80%">注册过滤字符将不允许用户注册包含以下字符的内容，请您将要过滤的字符串添入，如果有多个字符串，请用“,”分隔开，例如：沙滩,quest,木鸟</td>
</tr>
<tr>
<td class=forumrow width="20%">请输入过滤字符</td>
<td class=forumrow width="80%"><input type="text" name="splitwords" value="<%=splitwords%>" size="80"></td>
</tr>
<%end if%>
<input type=hidden value="<%=request("reaction")%>" name="reaction">
<tr> 
<td class=forumrow width="20%"></td>
<td width="80%" class=forumrow><input type="submit" name="Submit" value="提 交"></td>
</tr>
</table>
</form>
<%end if%>
<%
end sub

sub savebadword()
if request("skinid")<>"" then
if request("reaction")="badword" then
sql = "update config set badwords='"&request("badwords")&"' where id in ( "&request("skinid")&" )"
conn.execute(sql)
elseif request("reaction")="splitreg" then
sql = "update config set splitwords='"&request("splitwords")&"' where id in ( "&request("skinid")&" )"
conn.execute(sql)
end if
end if
response.write "更新成功！"
end sub
%>