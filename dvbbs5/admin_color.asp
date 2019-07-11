<!--#include file =conn.asp-->
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
		if request("action")="save" then
			call saveconst()
		else
			call consted()
		end if
		if founderr then call dvbbs_error()
		conn.close
		set conn=nothing
	end if

sub consted()
dim sel
%>
<form method="POST" action=admin_color.asp?action=save>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center>
<tr> 
<td height="23" colspan="2" ><B>说明</B>：<BR>1、复选框中选择的为当前的使用设置模板，点击可查看该模板设置，点击别的模板直接查看该模板并修改设置。您可以将您下面的设置保存在多个论坛风格模板中<BR>2、您也可以将下面设定的信息保存并应用到具体的分论坛设置中，可多选<BR>3、如果您想在一个版面引用别的版面或模板的配置，只要点击该版面或模板名称，保存的时候选择要保存到的版面名称或模板名称即可。
<BR><B>小知识</B>：<BR>在CSS定义中，每项定义用分号（;）分开。常用的定义有：<BR>1、background-color表示背景颜色，如background-color: #FFFFFF;
</td>
</tr>
</table>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th width="100%" colspan=3>当前使用主模板（可将设置保存到下列模板中）</th>
</tr>
<tr>
<td colspan=3 class=forumrow>
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
response.write "<input type=checkbox name=skinid value="&rs("id")&" "&sel&"><a href=admin_color.asp?skinid="&rs("id")&">"&rs("skinname")&"</a>&nbsp;"
rs.movenext
loop
rs.close
set rs=nothing
%></font>
</td></tr>
<tr> 
<td width="100%" colspan=3 class=forumrow><B>当前使用该模板的分论坛</B><BR>以下分论坛将使用本模板设置，如需修改论坛使用模板，可到论坛管理中重新设置<BR>
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
<script>
function color(para_URL){var URL =new String(para_URL)
window.open(URL,'','width=300,height=220,noscrollbars')}
</SCRIPT>
<tr> 
<th height="23" colspan="3" align=left>论坛界面设置&nbsp;单击 <a href="javascript:color('color.asp')"><font color=#FFFFFF>这里</font></a> 使用万用颜色拾取器（修改以下设置必须具备一定网页知识）</th>
</tr>
<tr> 
<td width="45%" class=forumrow>论坛BODY标签<br>
控制整个论坛风格的背景颜色或者背景图片等</td>
<td width="5%" class=forumrow></td>
<td width="50%" class=forumrow> 
<input type="text" name="ForumBody" size="50" value="<%=Forum_body(11)%>">
</td>
</tr>
<tr> 
<td width="45%" class=forumrow>Body的CSS定义<BR>可定义内容为网页字体颜色、背景、浏览器边框等<BR><FONT COLOR="red">以下定义中，凡是CSS定义的就不一定局限于颜色上的配置了，您可以设置其中的字体、背景等</FONT></td>
<td width="5%" class=forumrow></td>
<td width="50%" class=forumrow> 
<textarea name="iebarcolor" cols="50" rows="3"><%=Forum_body(25)%></textarea>
</td>
</tr>
<tr> 
<td width="45%" class=forumrow>论坛连接总的CSS定义<BR>可定义内容为连接字体颜色、样式等<BR></td>
<td width="5%" class=forumrow></td>
<td width="50%" class=forumrow> 
<textarea name="ForumCssLink" cols="50" rows="3"><%=Forum_body(6)%></textarea>
</td>
</tr>
<tr> 
<td width="45%" class=forumrow>论坛表格总的CSS定义<BR>这里为论坛总的表格定义，为一般表格的风格设置（如页面中除了导航和信息列表外的部分：尾部、头部等），可定义内容为背景、字体颜色、样式等<BR></td>
<td width="5%" class=forumrow></td>
<td width="50%" class=forumrow> 
<textarea name="ForumCssTD" cols="50" rows="3"><%=Forum_body(10)%></textarea>
</td>
</tr>
<tr> 
<td width="45%" class=forumrow>顶部菜单表格CSS定义(Logo & Banner上方)<BR><FONT COLOR="#000099">调用：Class=TopDarkNav</FONT></td>
<td width="5%" class=forumrow></td>
<td width="50%" class=forumrow> 
<textarea name="NavDarkcolor" cols="50" rows="3"><%=Forum_body(24)%></textarea>
</td>
</tr>
<tr> 
<td width="45%" class=forumrow>顶部菜单表格CSS定义(Logo & Banner下方)<BR><FONT COLOR="#000099">调用：Class=TopLighNav</font></td>
<td width="5%" class=forumrow></td>
<td width="50%" class=forumrow> 
<textarea name="Navlighcolor" cols="50" rows="3"><%=Forum_body(23)%></textarea>
</td>
</tr>
<tr> 
<td width="45%" class=forumrow>顶部菜单表格CSS定义(导航菜单)<BR><FONT COLOR="#000099">调用：Class=TopLighNav1</font></td>
<td width="5%" class=forumrow></td>
<td width="50%" class=forumrow> 
<textarea name="Navlighcolor1" cols="50" rows="3"><%=Forum_body(26)%></textarea>
</td>
</tr>
<tr> 
<td width="45%" class=forumrow>顶部菜单表格CSS定义(Logo & banner部分)<BR><FONT COLOR="#000099">调用：Class=TopLighNav2</font></td>
<td width="5%" class=forumrow></td>
<td width="50%" class=forumrow> 
<textarea name="Navlighcolor2" cols="50" rows="3"><%=Forum_body(28)%></textarea>
</td>
</tr>
<tr> 
<td width="45%" class=forumrow>表格边框颜色定义<br>
在下面定义一和二CSS定义控制不到的地方，最好保持和边框CSS定义一中同样的颜色</td>
<td width="5%" bgcolor="<%=Forum_body(27)%>"></td>
<td width="50%" class=forumrow> 
<input type=text name="btablebackcolor" value="<%=Forum_body(27)%>" size=35>
</td>
</tr>
<tr> 
<td width="45%" class=forumrow>表格边框（体）CSS定义一<br>
一般页面，可定义内容为主表格背景、背景图、边框、宽度等<BR><FONT COLOR="#000099">调用：Class=TableBorder1</font></td>
<td width="5%"  class=forumrow></td>
<td width="50%" class=forumrow> 
<textarea name="Tablebackcolor" cols="50" rows="3"><%=Forum_body(0)%></textarea>
</td>
</tr>
<tr> 
<td width="45%" class=forumrow>表格边框（体）CSS定义二<br>
分论坛导航表格边框及部分页面，可定义内容为主表格背景、背景图、边框、宽度等<BR><FONT COLOR="#000099">调用：Class=TableBorder2</font></td>
<td width="5%" class=forumrow></td>
<td width="50%" class=forumrow> 
<textarea name="aTablebackcolor" cols="50" rows="3"><%=Forum_body(1)%></textarea>
</td>
</tr>
<tr> 
<td width="45%" class=forumrow>标题表格CSS定义一（深背景）<br>
一般为表格的标题状态栏颜色或者背景上的定义，可定义内容为背景、背景图、字体及其颜色等<BR><FONT COLOR="#000099">调用：th</font></td>
<td width="5%" class=forumrow></td>
<td width="50%" class=forumrow> 
<textarea name="Tabletitlecolor" cols="50" rows="3"><%=Forum_body(2)%></textarea>
</td>
</tr>
<tr> 
<td width="45%" class=forumrow>标题表格字体连接CSS定义（深背景）<br>
如果标题栏为深背景，在这里必须特别定义在标题栏中连接的颜色和字体的颜色一样，否则在这里将采用默认的连接颜色<BR>注意：这里的#TableTitleLink不能去掉，可以将几个项目分开写来达到连接的不同效果<BR><FONT COLOR="#000099">调用：id=TableTitleLink</font></td>
<td width="5%" class=forumrow></td>
<td width="50%" class=forumrow> 
<textarea name="TableTitleLink" cols="50" rows="3"><%=Forum_body(7)%></textarea>
</td>
</tr>
<tr> 
<td width="45%" class=forumrow>标题表格CSS定义二（浅背景）<br>
为二级标题状态栏颜色或者背景上的定义，如首页的论坛分类栏，可定义内容为背景、背景图、字体及其颜色等<BR><FONT COLOR="#000099">调用：Class=TableTitle2</font></td>
<td width="5%" class=forumrow></td>
<td width="50%" class=forumrow> 
<textarea name="aTabletitlecolor" cols="50" rows="3"><%=Forum_body(3)%></textarea>
</td>
</tr>
<tr> 
<td width="45%" class=forumrow>表格体CSS定义一<BR>可定义内容为背景、背景图、字体及其颜色等<BR><FONT COLOR="#000099">调用：Class=TableBody1</font></td>
<td width="5%" class=forumrow></td>
<td width="50%" class=forumrow> 
<textarea name="Tablebodycolor" cols="50" rows="3"><%=Forum_body(4)%></textarea>
</td>
</tr>
<tr> 
<td width="45%" class=forumrow>表格体颜色二(1和2颜色在显示中穿插)<BR>可定义内容为背景、背景图、字体及其颜色等<BR><FONT COLOR="#000099">调用：Class=TableBody2</font></td>
<td width="5%" class=forumrow></td>
<td width="50%" class=forumrow> 
<textarea name="aTablebodycolor" cols="50" rows="3"><%=Forum_body(5)%></textarea>
</td>
</tr>
<tr> 
<td width="45%" class=forumrow>表格宽度</td>
<td width="5%" class=forumrow></td>
<td width="50%" class=forumrow> 
<input type="text" name="TableWidth" size="35" value="<%=Forum_body(12)%>">
</td>
</tr>
<tr> 
<td width="45%" class=forumrow>警告提醒语句的颜色</td>
<td width="5%" bgcolor="<%=Forum_body(8)%>"></td>
<td width="50%" class=forumrow> 
<input type="text" name="AlertFontColor" size="35" value="<%=Forum_body(8)%>">
</td>
</tr>
<tr> 
<td width="45%" class=forumrow>显示帖子的时候，相关帖子，转发帖子，回复等的颜色</td>
<td width="5%" bgcolor="<%=Forum_body(9)%>"></td>
<td width="50%" class=forumrow> 
<input type="text" name="ContentTitle" size="35" value="<%=Forum_body(9)%>">
</td>
</tr>
<tr> 
<td width="45%" class=forumrow>首页连接颜色<BR>如版面连接</td>
<td width="5%" bgcolor="<%=Forum_body(14)%>"></td>
<td width="50%" class=forumrow> 
<input type="text" name="BoardLinkColor" size="35" value="<%=Forum_body(14)%>">
</td>
</tr>
<tr> 
<td width="45%" class=forumrow>一般用户名称字体颜色</td>
<td width="5%" bgcolor="<%=Forum_body(15)%>"></td>
<td width="50%" class=forumrow> 
<input type="text" name="user_fc" size="35" value="<%=Forum_body(15)%>">
</td>
</tr>
<tr> 
<td width="45%" class=forumrow>一般用户名称上的光晕颜色</td>
<td width="5%" bgcolor="<%=Forum_body(16)%>"></td>
<td width="50%" class=forumrow> 
<input type="text" name="user_mc" size="35" value="<%=Forum_body(16)%>">
</td>
</tr>
<tr> 
<td width="45%" class=forumrow>版主名称字体颜色</td>
<td width="5%" bgcolor="<%=Forum_body(17)%>"></td>
<td width="50%" class=forumrow> 
<input type="text" name="bmaster_fc" size="35" value="<%=Forum_body(17)%>">
</td>
</tr>
<tr> 
<td width="45%" class=forumrow>版主名称上的光晕颜色</td>
<td width="5%" bgcolor="<%=Forum_body(18)%>"></td>
<td width="50%" class=forumrow> 
<input type="text" name="bmaster_mc" size="35" value="<%=Forum_body(18)%>">
</td>
</tr>
<tr> 
<td width="45%" class=forumrow>管理员名称字体颜色</td>
<td width="5%" bgcolor="<%=Forum_body(19)%>"></td>
<td width="50%" class=forumrow> 
<input type="text" name="master_fc" size="35" value="<%=Forum_body(19)%>">
</td>
</tr>
<tr> 
<td width="45%" class=forumrow>版主名称上的光晕颜色</td>
<td width="5%" bgcolor="<%=Forum_body(20)%>"></td>
<td width="50%" class=forumrow> 
<input type="text" name="master_mc" size="35" value="<%=Forum_body(20)%>">
</td>
</tr>
<tr> 
<td width="45%" class=forumrow>贵宾名称字体颜色</td>
<td width="5%" bgcolor="<%=Forum_body(21)%>"></td>
<td width="50%" class=forumrow> 
<input type="text" name="vip_fc" size="35" value="<%=Forum_body(21)%>">
</td>
</tr>
<tr> 
<td width="45%" class=forumrow>贵宾名称上的光晕颜色</td>
<td width="5%" bgcolor="<%=Forum_body(22)%>"></td>
<td width="50%" class=forumrow> 
<input type="text" name="vip_mc" size="35" value="<%=Forum_body(22)%>">
</td>
</tr>
<tr> 
<td width="45%" class=Forumrow>&nbsp;</td>
<td width="5%" class=Forumrow></td>
<td width="50%" class=Forumrow> 
<input type="submit" name="Submit" value="提 交">
</td>
</tr>
</table>
</form>
<%
end sub

sub saveconst()
dim Forum_body
if trim(request("skinid"))="" and trim(request("boardid"))="" then
Founderr=true
Errmsg=Errmsg+"<br>"+"<li>请选择保存的模板名称"
else
Forum_body=request.form("Tablebackcolor") & "|||" & request.form("aTablebackcolor") & "|||" & request.form("Tabletitlecolor") & "|||" & request.form("aTabletitlecolor") & "|||" & request.form("Tablebodycolor") & "|||" & request.form("aTablebodycolor") & "|||" & request.form("ForumCssLink") & "|||" & request.form("TableTitleLink") & "|||" & request.form("AlertFontColor") & "|||" & request.form("ContentTitle") & "|||" & request.form("ForumCssTD") & "|||" & request.form("ForumBody") & "|||" & request.form("TableWidth") & "|||" & request.form("BodyFontColor") & "|||" & request.form("BoardLinkColor") & "|||" & request.form("user_fc") & "|||" & request.form("user_mc") & "|||" & request.form("bmaster_fc") & "|||" & request.form("bmaster_mc") & "|||" & request.form("master_fc") & "|||" & request.form("master_mc") & "|||" & request.form("vip_fc") & "|||" & request.form("vip_mc") & "|||" & request.form("NavLighColor") & "|||" & request.form("NavDarkColor") & "|||" & request.form("IEbarColor") & "|||" & request.form("Navlighcolor1") & "|||" & request.form("btablebackcolor") & "|||" & request.form("Navlighcolor2")
'response.write Forum_body
if request("skinid")<>"" then
sql = "update config set Forum_body='"&Forum_body&"' where id in ( "&request("skinid")&" )"
conn.execute(sql)
end if
if request("boardid")<>"" then
sql = "update board set Forum_body='"&Forum_body&"' where boardid in ( "&request("boardid")&" )"
conn.execute(sql)
end if
response.write "设置成功。"
end if
end sub
%>