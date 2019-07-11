<!--#include file=conn.asp-->
<!-- #include file="inc/const.asp" -->
<title><%=Forum_info(0)%>--管理页面</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta NAME=GENERATOR Content=""Microsoft FrontPage 3.0"" CHARSET=GB2312>
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0" marginheight="0" marginwidth="0">
<%
	if not master or session("flag")="" then
		Errmsg=Errmsg+"<br>"+"<li>本页面为管理员专用，请<a href=admin_index.asp target=_top>登陆</a>后进入。<br><li>您没有管理本页面的权限。"
		call dvbbs_error()
	else
		if request("action")="save" then
		call savegrade()
		else
		call grade()
		end if
		if founderr then call dvbbs_error()
		conn.close
		set conn=nothing
	end if

sub grade()
dim sel
%>
<form method="POST" action=admin_wealth.asp?action=save>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center>
<tr> 
<td height="23" colspan="2" ><B>说明</B>：<BR>1、复选框中选择的为当前的使用设置模板，点击可查看该模板设置，点击别的模板直接查看该模板并修改设置。您可以将您下面的设置保存在多个论坛风格模板中<BR>2、您也可以将下面设定的信息保存并应用到具体的分论坛设置中，可多选<BR>3、如果您想在一个版面引用别的版面或模板的配置，只要点击该版面或模板名称，保存的时候选择要保存到的版面名称或模板名称即可。</font></td>
</tr>
</table>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th width="100%" colspan=2><B>当前使用主模板</B>（可将设置保存到下列模板中）</th>
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
	skinid=rs("id")
	else
	sel=""
	end if
end if
response.write "<input type=checkbox name=skinid value="&rs("id")&" "&sel&"><a href=admin_wealth.asp?skinid="&rs("id")&">"&rs("skinname")&"</a>&nbsp;"
rs.movenext
loop
rs.close
set rs=nothing
%>
</td></tr>
<tr> 
<td width="100%" colspan=2 class=Forumrow><B>当前使用该模板的分论坛</B><BR>以下分论坛将使用本模板设置，如需修改论坛使用模板，可到论坛管理中重新设置<BR>
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
<tr> 
<th height="23" colspan="2">用户金钱设定</th>
</tr>
<tr> 
<td width="40%" class=Forumrow>注册金钱数</td>
<td width="60%" class=Forumrow> 
<input type="text" name="wealthReg" size="35" value="<%=Forum_user(0)%>">
</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>登陆增加金钱</td>
<td width="60%" class=Forumrow> 
<input type="text" name="wealthLogin" size="35" value="<%=Forum_user(4)%>">
</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>发帖增加金钱</td>
<td width="60%" class=Forumrow> 
<input type="text" name="wealthAnnounce" size="35" value="<%=Forum_user(1)%>">
</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>跟帖增加金钱</td>
<td width="60%" class=Forumrow> 
<input type="text" name="wealthReannounce" size="35" value="<%=Forum_user(2)%>">
</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>精华增加金钱</td>
<td width="60%" class=Forumrow> 
<input type="text" name="BestWealth" size="35" value="<%=Forum_user(15)%>">
</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>删帖减少金钱</td>
<td width="60%" class=Forumrow> 
<input type="text" name="wealthDel" size="35" value="<%=Forum_user(3)%>">
</td>
</tr>
<tr> 
<th height="23" colspan="2" >用户经验设定</th>
</tr>
<tr> 
<td width="40%" class=Forumrow>注册经验值</td>
<td width="60%" class=Forumrow> 
<input type="text" name="epReg" size="35" value="<%=Forum_user(5)%>">
</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>登陆增加经验值</td>
<td width="60%" class=Forumrow> 
<input type="text" name="epLogin" size="35" value="<%=Forum_user(9)%>">
</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>发帖增加经验值</td>
<td width="60%" class=Forumrow> 
<input type="text" name="epAnnounce" size="35" value="<%=Forum_user(6)%>">
</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>跟帖增加经验值</td>
<td width="60%" class=Forumrow> 
<input type="text" name="epReannounce" size="35" value="<%=Forum_user(7)%>">
</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>精华增加经验值</td>
<td width="60%" class=Forumrow> 
<input type="text" name="bestuserep" size="35" value="<%=Forum_user(17)%>">
</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>删帖减少经验值</td>
<td width="60%" class=Forumrow> 
<input type="text" name="epDel" size="35" value="<%=Forum_user(8)%>">
</td>
</tr>
<tr> 
<th height="23" colspan="2" >用户魅力设定</th>
</tr>
<tr> 
<td width="40%" class=Forumrow>注册魅力值</td>
<td width="60%" class=Forumrow> 
<input type="text" name="cpReg" size="35" value="<%=Forum_user(10)%>">
</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>登陆增加魅力值</td>
<td width="60%" class=Forumrow> 
<input type="text" name="cpLogin" size="35" value="<%=Forum_user(14)%>">
</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>发帖增加魅力值</td>
<td width="60%" class=Forumrow> 
<input type="text" name="cpAnnounce" size="35" value="<%=Forum_user(11)%>">
</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>跟帖增加魅力值</td>
<td width="60%" class=Forumrow> 
<input type="text" name="cpReannounce" size="35" value="<%=Forum_user(12)%>">
</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>精华增加魅力值</td>
<td width="60%" class=Forumrow> 
<input type="text" name="bestusercp" size="35" value="<%=Forum_user(16)%>">
</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>删帖减少魅力值</td>
<td width="60%" class=Forumrow> 
<input type="text" name="cpDel" size="35" value="<%=Forum_user(13)%>">
</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>&nbsp;</td>
<td width="60%" class=Forumrow> 
<div align="center"> 
<input type="submit" name="Submit" value="提 交">
</div>
</td>
</tr>
</table>
</form>
<%
end sub

sub savegrade()
if trim(request("skinid"))="" and trim(request("boardid"))="" then
Founderr=true
Errmsg=Errmsg+"<br>"+"<li>请选择保存的模板名称"
else
Forum_user=request.form("wealthReg") & "," & request.form("wealthAnnounce") & "," & request.form("wealthReannounce") & "," & request.form("wealthDel") & "," & request.form("wealthLogin") & "," & request.form("epReg") & "," & request.form("epAnnounce") & "," & request.form("epReannounce") & "," & request.form("epDel") & "," & request.form("epLogin") & "," & request.form("cpReg") & "," & request.form("cpAnnounce") & "," & request.form("cpReannounce") & "," & request.form("cpDel") & "," & request.form("cpLogin") & "," & request.form("BestWealth") & "," & request.form("BestuserCP") & "," & request.form("BestuserEP")
'response.write Forum_user
if request("skinid")<>"" then
sql = "update config set Forum_user='"&Forum_user&"' where id in ( "&request("skinid")&" )"
conn.execute(sql)
end if
response.write "设置成功。"
end if
end sub
%>