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
<form method="POST" action=admin_pic.asp?action=save>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center>
<tr> 
<td height="23" colspan="2" ><B>说明</B>：<BR>1、复选框中选择的为当前的使用设置模板，点击可查看该模板设置，点击别的模板直接查看该模板并修改设置。您可以将您下面的设置保存在多个论坛风格模板中<BR>2、您也可以将下面设定的信息保存并应用到具体的分论坛设置中，可多选<BR>3、如果您想在一个版面引用别的版面或模板的配置，只要点击该版面或模板名称，保存的时候选择要保存到的版面名称或模板名称即可。<br>4、以下图片均保存于论坛pic目录中，如要更换也请将图放于该目录
</td>
</tr>
</table>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th width="100%" colspan=2 align=left>当前使用主模板（可将设置保存到下列模板中）</th>
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
response.write "<input type=checkbox name=skinid value="&rs("id")&" "&sel&"><a href=admin_pic.asp?skinid="&rs("id")&">"&rs("skinname")&"</a>&nbsp;"
rs.movenext
loop
rs.close
set rs=nothing
%>
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
<tr bgcolor=<%=Forum_body(3)%>> 
<th height="23" colspan="2" align=left><font color=<%=Forum_body(6)%>>在线图片设置</th>
</tr>
<tr> 
<td width="40%" class=forumrow>论坛管理员</td>
<td width="60%" class=forumrow> 
<input type="text" name="pic_om" size="20" value="<%=Forum_pic(0)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_pic(0)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>论坛版主</td>
<td width="60%" class=forumrow> 
<input type="text" name="pic_ob" size="20" value="<%=Forum_pic(1)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_pic(1)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>论坛贵宾</td>
<td width="60%" class=forumrow> 
<input type="text" name="pic_ov" size="20" value="<%=Forum_pic(2)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_pic(2)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>普通会员</td>
<td width="60%" class=forumrow> 
<input type="text" name="pic_ou" size="20" value="<%=Forum_pic(3)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_pic(3)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>客人或隐身会员</td>
<td width="60%" class=forumrow> 
<input type="text" name="pic_oh" size="20" value="<%=Forum_pic(4)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_pic(4)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>突出显示自己的颜色</td>
<td width="60%" class=forumrow> 
<input type="text" name="online_mc" size="20" value="<%=Forum_pic(5)%>">&nbsp;&nbsp;<font color=<%=Forum_pic(5)%>>自己</font>
</td>
</tr>
<tr bgcolor=<%=Forum_body(3)%>> 
<th height="23" colspan="2" align=left>论坛图例</th>
</tr>
<tr> 
<td width="40%" class=forumrow>常规论坛－－有新帖子</td>
<td width="60%" class=forumrow> 
<input type="text" name="F_mode1_n" size="20" value="<%=Forum_pic(7)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_pic(7)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>常规论坛－－无新帖子</td>
<td width="60%" class=forumrow> 
<input type="text" name="F_mode1_o" size="20" value="<%=Forum_pic(6)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_pic(6)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>认证论坛－－有新帖子</td>
<td width="60%" class=forumrow> 
<input type="text" name="F_mode5_n" size="20" value="<%=Forum_pic(15)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_pic(15)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>认证论坛－－无新帖子</td>
<td width="60%" class=forumrow> 
<input type="text" name="F_mode5_o" size="20" value="<%=Forum_pic(14)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_pic(14)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>联盟论坛</td>
<td width="60%" class=forumrow> 
<input type="text" name="F_link" size="20" value="<%=Forum_boardpic(0)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_boardpic(0)%>>
</td>
</tr>
<tr bgcolor=<%=Forum_body(3)%>> 
<th height="23" colspan="2" align=left>论坛主体图标</th>
</tr>
<tr> 
<td width="40%" class=forumrow>发表帖子</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_post" size="20" value="<%=Forum_boardpic(1)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_boardpic(1)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>发表投票</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_vote" size="20" value="<%=Forum_boardpic(2)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_boardpic(2)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>回复帖子</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_reply" size="20" value="<%=Forum_boardpic(4)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_boardpic(4)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>小字报</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_paper" size="20" value="<%=Forum_boardpic(3)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_boardpic(3)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>帮助</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_help" size="20" value="<%=Forum_boardpic(5)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_boardpic(5)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>返回首页</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_gohome" size="20" value="<%=Forum_boardpic(6)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_boardpic(6)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>返回上级目录</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_folder" size="20" value="<%=Forum_boardpic(7)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_boardpic(7)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>当前目录</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_ofolder" size="20" value="<%=Forum_boardpic(8)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_boardpic(8)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>新的短消息</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_newmsg" size="20" value="<%=Forum_boardpic(9)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_boardpic(9)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>我发表的主题</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_mytopic" size="20" value="<%=Forum_boardpic(10)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_boardpic(10)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>树形浏览帖子模式</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_treeview" size="20" value="<%=Forum_boardpic(11)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_boardpic(11)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>平板形浏览帖子模式</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_flatview" size="20" value="<%=Forum_boardpic(12)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_boardpic(12)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>下一篇帖子</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_nexttopic" size="20" value="<%=Forum_boardpic(13)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_boardpic(13)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>上一篇帖子</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_backTopic" size="20" value="<%=Forum_boardpic(14)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_boardpic(14)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>刷新浏览</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_reflash" size="20" value="<%=Forum_boardpic(15)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_boardpic(15)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>论坛公告</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_call" size="20" value="<%=Forum_boardpic(16)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_boardpic(16)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>打开的主题</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_opentopic" size="20" value="<%=Forum_statePic(0)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_statePic(0)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>热门的主题</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_hotTopic" size="20" value="<%=Forum_statePic(1)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_statePic(1)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>锁定的主题</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_closeTopic" size="20" value="<%=Forum_statePic(2)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_statePic(2)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>固顶的主题</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_istop" size="20" value="<%=Forum_statePic(3)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_statePic(3)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>精华的主题</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_Tisbest" size="20" value="<%=Forum_statePic(4)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_statePic(4)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>投票的主题</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_isvote" size="20" value="<%=Forum_statePic(12)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_statePic(12)%>>
</td>
</tr>
<tr bgcolor=<%=Forum_body(3)%>> 
<th height="23" colspan="2" align=left>论坛帖子图标</th>
</tr>
<tr> 
<td width="40%" class=forumrow>保存当页为文件</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_save" size="20" value="<%=Forum_TopicPic(0)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_TopicPic(0)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>报告本帖给版主</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_report" size="20" value="<%=Forum_TopicPic(1)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_TopicPic(1)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>显示可打印的版本</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_print" size="20" value="<%=Forum_TopicPic(2)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_TopicPic(2)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>把本帖加入论坛收藏</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_bbsfav" size="20" value="<%=Forum_TopicPic(4)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_TopicPic(4)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>把本帖打包邮递</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_EmailPost" size="20" value="<%=Forum_TopicPic(3)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_TopicPic(3)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>发送本页面给朋友</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_EmailPage" size="20" value="<%=Forum_TopicPic(5)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_TopicPic(5)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>把本帖加入IE收藏</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_iefav" size="20" value="<%=Forum_TopicPic(6)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_TopicPic(6)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>发送短信给好友</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_sms" size="20" value="<%=Forum_TopicPic(7)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_TopicPic(7)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>加为好友</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_friend" size="20" value="<%=Forum_TopicPic(21)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_TopicPic(21)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>浏览用户信息</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_userinfo" size="20" value="<%=Forum_TopicPic(9)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_TopicPic(9)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>搜索用户发表帖子</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_search" size="20" value="<%=Forum_TopicPic(8)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_TopicPic(8)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>用户邮箱</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_Email" size="20" value="<%=Forum_TopicPic(10)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_TopicPic(10)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>用户OICQ</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_oicq" size="20" value="<%=Forum_TopicPic(11)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_TopicPic(11)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>用户ICQ</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_icq" size="20" value="<%=Forum_TopicPic(12)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_TopicPic(12)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>用户MSN</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_msn" size="20" value="<%=Forum_TopicPic(13)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_TopicPic(13)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>用户主页</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_homepage" size="20" value="<%=Forum_TopicPic(14)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_TopicPic(14)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>引用帖子</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_quote" size="20" value="<%=Forum_TopicPic(15)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_TopicPic(15)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>编辑帖子</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_edit" size="20" value="<%=Forum_TopicPic(16)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_TopicPic(16)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>删除帖子</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_delete" size="20" value="<%=Forum_TopicPic(17)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_TopicPic(17)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>复制帖子</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_copy" size="20" value="<%=Forum_TopicPic(18)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_TopicPic(18)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>加入精华帖子</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_isbest" size="20" value="<%=Forum_TopicPic(19)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_TopicPic(19)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>用户IP</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_ip" size="20" value="<%=Forum_TopicPic(20)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_TopicPic(20)%>>
</td>
</tr>
<tr bgcolor=<%=Forum_body(3)%>> 
<td width="40%" class=forumrow>&nbsp;</td>
<td width="60%" class=forumrow> 
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
Forum_pic=request("pic_om") & "," & request("pic_ob") & "," & request("pic_ov") & "," & request("pic_ou") & "," & request("pic_oh") & "," & request("online_mc") & "," & request("F_mode1_o") & "," & request("F_mode1_n") & "," & request("F_mode2_o") & "," & request("F_mode2_n") & "," & request("F_mode3_o") & "," & request("F_mode3_n") & "," & request("F_mode4_n") & "," & request("F_mode4_n") & "," & request("F_mode5_o") & "," & request("F_mode5_n") & "," & request("F_mode6_o") & "," & request("F_mode6_n") & ",ifolder.gif,foldernew.gif"

Forum_boardpic=request("F_link") & "," & request("P_post") & "," & request("P_vote") & "," & request("P_paper") & "," & request("P_reply") & "," & request("P_help") & "," & request("P_gohome") & "," & request("P_folder") & "," & request("P_ofolder") & "," & request("P_newmsg") & "," & request("P_mytopic") & "," & request("P_treeview") & "," & request("P_flatview") & "," & request("P_nexttopic") & "," & request("P_backTopic") & "," & request("P_reflash") & "," & request("P_call")

Forum_TopicPic=request("P_save") & "," & request("P_report") & "," & request("P_print") & "," & request("P_EmailPost") & "," & request("P_bbsfav") & "," & request("P_EmailPage") & "," & request("P_iefav") & "," & request("P_sms") & "," & request("P_search") & "," & request("P_userinfo") & "," & request("P_Email") & "," & request("P_oicq") & "," & request("P_icq") & "," & request("P_msn") & "," & request("P_homepage") & "," & request("P_quote") & "," & request("P_edit") & "," & request("P_delete") & "," & request("P_copy") & "," & request("P_isbest") & "," & request("P_ip") & "," & request("P_friend")

Forum_statePic=request("P_opentopic") & "," & request("P_hotTopic") & "," & request("P_closeTopic") & "," & request("P_istop") & "," & request("P_Tisbest") & "," & request("P_newTopic") & "," & request("P_userlist") & "," & request("P_Top20") & "," & request("P_nowTime") & "," & request("P_online_s") & "," & request("P_birth") & "," & request("P_userfrom") & "," & request("P_isvote")

'Forum_ubb=request("bold") & "," & request("italicize") & "," & request("underline") & "," & request("center") & "," & request("url1") & "," & request("email1") & "," & request("image") & "," & request("swf") & "," & request("Shockwave") & "," & request("rm") & "," & request("mp") & "," & request("qt") & "," & request("quote1") & "," & request("fly") & "," & request("move") & "," & request("glow") & "," & request("shadow")
'response.write Forum_Topicpic & "<br>" & Forum_boardpic & "<br>" & Forum_statePic
if request("skinid")<>"" then
sql = "update config set Forum_pic='"&Forum_pic&"',Forum_boardpic='"&Forum_boardpic&"',Forum_TopicPic='"&Forum_TopicPic&"',Forum_statePic='"&Forum_statePic&"' where id in ( "&request("skinid")&" )"
conn.execute(sql)
end if
response.write "设置成功。"
end if
end sub
%>