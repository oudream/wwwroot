<!--#include file =conn.asp-->
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
<form method="POST" action=admin_ads.asp?action=save>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center>
<tr> 
<td height="23" colspan="2"><B>说明</B>：<BR>1、复选框中选择的为当前的使用设置模板，点击可查看该模板设置，点击别的模板直接查看该模板并修改设置。您可以将您下面的设置保存在多个论坛风格模板中<BR>2、您也可以将下面设定的信息保存并应用到具体的分论坛设置中，可多选<BR>3、如果您想在一个版面引用别的版面或模板的配置，只要点击该版面或模板名称，保存的时候选择要保存到的版面名称或模板名称即可。</td>
</tr>
</table>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th width="100%" colspan=2 class="tableHeaderText"><B>当前使用主模板</B>（可将设置保存到下列模板中）
</th>
</tr>
<tr>
<td colspan=2 class="forumrow">
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
response.write "<input type=checkbox name=skinid value="&rs("id")&" "&sel&"><a href=admin_ads.asp?skinid="&rs("id")&">"&rs("skinname")&"</a>&nbsp;"
rs.movenext
loop
rs.close
set rs=nothing
%></font>
</td></tr>
<tr> 
<td width="100%" colspan=2  class="forumrow"><B>当前使用该模板的分论坛</B><BR>以下分论坛将使用本模板设置，如需修改论坛使用模板，可到论坛管理中重新设置<BR>
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
<th height="23" colspan="2" class="tableHeaderText"><b>论坛广告设置</b>（如为设置分论坛，就是分论坛首页广告，下属页面为帖子显示页面）</th>
</tr>
<tr> 
<td width="40%" class="forumrow"><B>首页顶部广告代码</B></font></td>
<td width="60%" class="forumrow"> 
<textarea name="index_ad_t" cols="50" rows="3"><%=Forum_ads(0)%></textarea>
</td>
</tr>
<tr> 
<td width="40%" class="forumrow"><B>首页尾部广告代码</B></font></td>
<td width="60%" class="forumrow"> 
<textarea name="index_ad_f" cols="50" rows="3"><%=Forum_ads(1)%></textarea>
</td>
</tr>
<tr> 
<td width="40%" class="forumrow"><B>开启首页浮动广告</B></font></td>
<td width="60%" class="forumrow"> 
<input type=radio name="index_moveFlag" value=0 <%if Forum_ads(2)=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="index_moveFlag" value=1 <%if Forum_ads(2)=1 then%>checked<%end if%>>打开&nbsp;
</td>
</tr>
<tr> 
<td width="40%" class="forumrow"><B>论坛首页浮动广告图片地址</B></font></td>
<td width="60%" class="forumrow"> 
<input type="text" name="MovePic" size="35" value="<%=Forum_ads(3)%>">
</td>
</tr>
<tr> 
<td width="40%" class="forumrow"><B>论坛首页浮动广告连接地址</B></font></td>
<td width="60%" class="forumrow"> 
<input type="text" name="MoveUrl" size="35" value="<%=Forum_ads(4)%>">
</td>
</tr>
<tr> 
<td width="40%" class="forumrow"><B>论坛首页浮动广告图片宽度</B></font></td>
<td width="60%" class="forumrow"> 
<input type="text" name="move_w" size="3" value="<%=Forum_ads(5)%>">&nbsp;象素
</td>
</tr>
<tr> 
<td width="40%" class="forumrow"><B>论坛首页浮动广告图片高度</B></font></td>
<td width="60%" class="forumrow"> 
<input type="text" name="move_h" size="3" value="<%=Forum_ads(6)%>">&nbsp;象素
</td>
</tr>
<input type=hidden name="Board_moveFlag" value=0>
<tr> 
<td width="40%" class="forumrow"><B>开启首页右下固定广告</B></font></td>
<td width="60%" class="forumrow"> 
<input type=radio name="index_fixupFlag" value=0 <%if Forum_ads(13)=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="index_fixupFlag" value=1 <%if Forum_ads(13)=1 then%>checked<%end if%>>打开&nbsp;
</td>
</tr>
<tr> 
<td width="40%" class="forumrow"><B>论坛首页右下固定广告图片地址</B></font></td>
<td width="60%" class="forumrow"> 
<input type="text" name="fixupPic" size="35" value="<%=Forum_ads(8)%>">
</td>
</tr>
<tr> 
<td width="40%" class="forumrow"><B>论坛首页右下固定广告连接地址</B></font></td>
<td width="60%" class="forumrow"> 
<input type="text" name="fixupUrl" size="35" value="<%=Forum_ads(9)%>">
</td>
</tr>
<tr> 
<td width="40%" class="forumrow"><B>论坛首页右下固定广告图片宽度</B></font></td>
<td width="60%" class="forumrow"> 
<input type="text" name="fixup_w" size="3" value="<%=Forum_ads(10)%>">&nbsp;象素
</td>
</tr>
<tr> 
<td width="40%" class="forumrow"><B>论坛首页右下固定广告图片高度</B></font></td>
<td width="60%" class="forumrow"> 
<input type="text" name="fixup_h" size="3" value="<%=Forum_ads(11)%>">&nbsp;象素
</td>
</tr>
<input type=hidden name="Board_fixupFlag" value=0>
<tr> 
<td width="40%" class="forumrow">&nbsp;</td>
<td width="60%" class="forumrow"> 
<div align="center"> 
<input type="submit" name="Submit" value="提 交">
</div>
</td>
</tr>
</table>
</form>
<%
end sub

sub saveconst()
if trim(request("skinid"))="" and trim(request("boardid"))="" then
Founderr=true
Errmsg=Errmsg+"<br>"+"<li>请选择保存的模板名称"
else
Forum_ads=request("index_ad_t") & "$" & request("index_ad_f") & "$" & request("index_moveFlag") & "$" & request("MovePic") & "$" & request("MoveUrl") & "$" & request("move_w") & "$" & request("move_h") & "$" & request("Board_moveFlag") & "$" & request("fixupPic") & "$" & request("FixupUrl") & "$" & request("Fixup_w") & "$" & request("Fixup_h") & "$" & request("Board_fixupFlag") & "$" & request("index_fixupFlag")
'response.write Forum_ads
if request("skinid")<>"" then
sql = "update config set Forum_ads='"&CheckStr(Forum_ads)&"' where id in ( "&request("skinid")&" )"
conn.execute(sql)
end if
response.write "设置成功。"
end if
end sub
%>