<!--#include file =conn.asp-->
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
		call saveconst()
		else
		call consted()
		end if
		conn.close
		set conn=nothing
	end if


sub consted()
dim sel
%>
<form method="POST" action=admin_setting.asp?action=save>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center>
<tr> 
<td height=23 colspan=2 align=left><B>说明</B>：这里的设置代表了所有论坛中所有的设置，包括总设置和分论坛，在这里您可以设置几套不同的方案以供各个论坛使用。<b>下列设置如有和用户组功能重复的，以用户组设置为准</b></td>
</tr>
<tr> 
<td height=23 colspan=2 align=left>
<table width=85% align=center>
<tr>
<td width=50% ><a name=top></a><a href="#setting1">[打开或者关闭论坛]</a></td>
<td width=50% ><a href="#setting2">[灌水检查选项]</a></td>
</tr>
<tr>
<td width=50% ><a href="#setting3">[基本信息]</a></td>
<td width=50% ><a href="#setting4">[联系资料]</a></td>
</tr>
<tr>
<td width=50% ><a href="#setting5">[发帖选项]</a></td>
<td width=50% ><a href="#setting6">[悄悄话选项]</a></td>
</tr>
<tr>
<td width=50% ><a href="#setting7">[论坛首页选项]</a></td>
<td width=50% ><a href="#setting8">[用户与注册选项]</a></td>
</tr>
<tr>
<td width=50% ><a href="#setting9">[投票选项]</a></td>
<td width=50% ><a href="#setting10">[服务器选项(脚本超时、组件等)]</a></td>
</tr>
<tr>
<td width=50% ><a href="#setting11">[搜索和查看资料选项]</a></td>
<td width=50% ><a href="#setting12">[在线和用户来源]</a></td>
</tr>
<tr>
<td width=50% ><a href="#setting13">[邮件选项]</a></td>
<td width=50% ><a href="#setting14">[上传设置]</a></td>
</tr>
<tr>
<td width=50% ><a href="#setting15">[用户选项(签名、头衔、排行等)]</a></td>
<td width=50% ><a href="#setting16">[浏览帖子选项]</a></td>
</tr>
<tr>
<td width=50% ><a href="#setting17">[防刷新机制]</a></td>
<td width=50% ><a href="#setting18">[版面设置]</a></td>
</tr>
<tr>
<td width=50% ><a href="#setting19">[门派设置]</a></td>
<td width=50% ></td>
</tr>
</table>
</td>
</tr> 
</table>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th width="100%" colspan=2  class="tableHeaderText">当前使用设置（可将当前设置保存到下列设置模板中，选中即可）
</th></tr>
<tr> 
<td width="100%" colspan=2 class="forumrow">
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
response.write "<input type=checkbox name=skinid value="&rs("id")&" "&sel&"><a href=admin_setting.asp?skinid="&rs("id")&">"&rs("skinname")&"</a>&nbsp;"
rs.movenext
loop
rs.close
set rs=nothing
%></font>
</td></tr>
<tr> 
<td width="100%" colspan=2 class="Forumrow"><B>当前使用该设置的分论坛</B><BR>以下分论坛将使用本设置，如需修改论坛使用设置，可到论坛管理中重新设置<BR>
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
<th height="23" colspan="2" align=left><a name="setting1"><b>打开或者关闭论坛</b></a>[<a href="#top"><FONT color="#FFFFFF">顶部</FONT></a>]</th>
</tr>
<tr> 
<td width="50%" class=Forumrow><U>论坛当前状态</U><BR>维护期间可设置关闭论坛使用</td>
<td width="50%" class=Forumrow> 
<input type=radio name="ForumStop" value=0 <%if cint(Forum_Setting(21))=0 then%>checked<%end if%>>打开&nbsp;
<input type=radio name="ForumStop" value=1 <%if cint(Forum_Setting(21))=1 then%>checked<%end if%>>关闭&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow><U>维护说明</U><BR>在论坛关闭情况下显示，支持html语法</td>
<td width="50%" class=Forumrow> 
<textarea name="StopReadme" cols="40" rows="3"><%=StopReadme%></textarea>
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>首页模式</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="IpFlag" value=0 <%if cint(Forum_Setting(4))=0 then%>checked<%end if%>>新版样式&nbsp;
<input type=radio name="IpFlag" value=1 <%if cint(Forum_Setting(4))=1 then%>checked<%end if%>>旧版样式&nbsp;
</td>
</tr>
<tr> 
<th align=left height="23" colspan="2" ><a name="setting2"><b>灌水检查选项</b></a>[<a href="#top"><FONT color="#FFFFFF">顶部</font></a>]</th>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>防灌水机制</U><BR>如选择打开请填写下面的限制发贴间隔时间<BR>对版主和管理员无效</td>
<td width="50%" class=Forumrow>  
<input type=radio name="relaypost" value=0 <%if cint(Forum_Setting(42))=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="relaypost" value=1 <%if cint(Forum_Setting(42))=1 then%>checked<%end if%>>打开&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>防灌水时间间隔</U><BR>填写该项目请确认您打开了防灌水机制</td>
<td width="50%" class=Forumrow>  
<input type="text" name="relaypostTime" size="3" value="<%=Forum_Setting(43)%>">&nbsp;秒
</td>
</tr>
<tr> 
<th height=23 colspan=2 align=left><a name="setting3"><b>基本信息</b></a>[<a href="#top"><font color=#FFFFFF>顶部</font></a>]</th>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>论坛名称</U></td>
<td width="50%" class=Forumrow>  
<input type="text" name="ForumName" size="35" value="<%=Forum_info(0)%>">
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>论坛的url</U></td>
<td width="50%" class=Forumrow>  
<input type="text" name="ForumURL" size="35" value="<%=Forum_info(1)%>">
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>主页名称</U></td>
<td width="50%" class=Forumrow>  
<input type="text" name="CompanyName" size="35" value="<%=Forum_info(2)%>">
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>主页URL</U></td>
<td width="50%" class=Forumrow>  
<input type="text" name="HostUrl" size="35" value="<%=Forum_info(3)%>">
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>论坛首页Logo地址</U><BR>显示在论坛首页，添加论坛如果没有填写logo地址，则使用该内容</td>
<td width="50%" class=Forumrow>  
<input type="text" name="Logo" size="35" value="<%=Forum_info(6)%>">
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>论坛图片目录</U></td>
<td width="50%" class=Forumrow>  
<input type="text" name="Picurl" size="35" value="<%=Forum_info(7)%>">
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>论坛表情目录</U></td>
<td width="50%" class=Forumrow>  
<input type="text" name="Faceurl" size="35" value="<%=Forum_info(8)%>">
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>版权信息</U></td>
<td width="50%" class=Forumrow>  
<input type="text" name="Copyright" size="35" value="<%=Copyright%>">
</td>
</tr>
<tr> 
<th height=23 colspan=2 align=left><a name="setting4"><b>联系资料</b></a>[<a href="#top"><font color=#FFFFFF>顶部</font></a>]</th>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>论坛管理员Email</U><BR>给用户发送邮件时，显示的来源Email信息</td>
<td width="50%" class=Forumrow>  
<input type="text" name="SystemEmail" size="35" value="<%=Forum_info(5)%>">
</td>
</tr>
<tr> 
<th height=23 colspan=2 align=left><a name="setting5"><b>发贴选项</b></a>[<a href="#top"><font color=#FFFFFF>顶部</font></a>]</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>发贴模式</U><BR>高级模式为显示所有项目</td>
<td width="50%" class=Forumrow>  
<input type=radio name="TopicFlag" value=0 <%if cint(Forum_Setting(17))=0 then%>checked<%end if%>>普通&nbsp;
<input type=radio name="TopicFlag" value=1 <%if cint(Forum_Setting(17))=1 then%>checked<%end if%>>高级&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>是否开启UBB代码</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="strAllowForumCode" value=0 <%if cint(Forum_Setting(34))=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="strAllowForumCode" value=1 <%if cint(Forum_Setting(34))=1 then%>checked<%end if%>>打开&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>是否开启HTML代码</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="strAllowHTML" value=0 <%if cint(Forum_Setting(35))=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="strAllowHTML" value=1 <%if cint(Forum_Setting(35))=1 then%>checked<%end if%>>打开&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>是否开启贴图标签</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="strIMGInPosts" value=0 <%if cint(Forum_Setting(36))=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="strIMGInPosts" value=1 <%if cint(Forum_Setting(36))=1 then%>checked<%end if%>>打开&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>是否开启表情标签</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="strIcons" value=0 <%if cint(Forum_Setting(37))=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="strIcons" value=1 <%if cint(Forum_Setting(37))=1 then%>checked<%end if%>>打开&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>是否开启Flash标签</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="strflash" value=0 <%if cint(Forum_Setting(38))=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="strflash" value=1 <%if cint(Forum_Setting(38))=1 then%>checked<%end if%>>打开&nbsp;
</td>
</tr>
<tr> 
<td class=Forumrow width="50%"><U>帖子内容最大的字节数</U><BR>请输入数字</td>
<td class=Forumrow width="50%"> 
<input type="text" name="AnnounceMaxBytes" size="6" value="<%=Forum_Setting(13)%>">&nbsp;字节
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>发新帖后返回</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="PostRetrun" value=1 <%if Forum_Setting(69)=1 then%>checked<%end if%>>首页&nbsp;
<input type=radio name="PostRetrun" value=2 <%if Forum_Setting(69)=2 then%>checked<%end if%>>论坛&nbsp;
<input type=radio name="PostRetrun" value=3 <%if Forum_Setting(69)=3 then%>checked<%end if%>>帖子&nbsp;
</td>
</tr>
<tr> 
<th height=23 colspan=2 align=left><a name="setting6"><b>悄悄话选项</b></a>[<a href="#top"><font color=#FFFFFF>顶部</font></a>]</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>新短消息弹出窗口</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="openmsg" value=0 <%if cint(Forum_Setting(10))=0 then%>checked<%end if%>>否&nbsp;
<input type=radio name="openmsg" value=1 <%if cint(Forum_Setting(10))=1 then%>checked<%end if%>>是&nbsp;
</td>
</tr>
<tr> 
<th height=23 colspan=2 align=left><a name="setting7"><b>论坛首页选项</b></a>[<a href="#top"><font color=#FFFFFF>顶部</font></a>]</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>快速登陆位置</U></td>
<td width="50%" class=Forumrow>  
<select name="FastLogin">
<option value="1" <%if cint(Forum_Setting(31))=1 then%>selected<%end if%>>顶部
<option value="2" <%if cint(Forum_Setting(31))=2 then%>selected<%end if%>>底部
<option value="0" <%if cint(Forum_Setting(31))="0" then%>selected<%end if%>>不显示
</select>
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>是否显示过生日会员</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="BirthFlag" value=0 <%if cint(Forum_Setting(29))=0 then%>checked<%end if%>>否&nbsp;
<input type=radio name="BirthFlag" value=1 <%if cint(Forum_Setting(29))=1 then%>checked<%end if%>>是&nbsp;
</td>
</tr>
<tr> 
<th height=23 colspan=2 align=left><a name="setting8"><b>用户与注册选项</b></a>[<a href="#top"><font color=#FFFFFF>顶部</font></a>]</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>同一IP注册间隔时间</U><BR>如不想限制可填写0</td>
<td width="50%" class=Forumrow>  
<input type="text" name="RegTime" size="3" value="<%=Forum_Setting(22)%>">&nbsp;秒
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>Email通知密码</U><BR>确认您的站点支持发送mail，所包含密码为系统随机生成</td>
<td width="50%" class=Forumrow>  
<input type=radio name="EmailReg" value=0 <%if cint(Forum_Setting(23))=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="EmailReg" value=1 <%if cint(Forum_Setting(23))=1 then%>checked<%end if%>>打开&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>一个Email只能注册一个帐号</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="EmailRegOne" value=0 <%if cint(Forum_Setting(24))=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="EmailRegOne" value=1 <%if cint(Forum_Setting(24))=1 then%>checked<%end if%>>打开&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>注册需要管理员认证</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="RegFlag" value=0 <%if cint(Forum_Setting(25))=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="RegFlag" value=1 <%if cint(Forum_Setting(25))=1 then%>checked<%end if%>>打开&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>发送注册信息邮件</U><BR>请确认您打开了邮件功能</td>
<td width="50%" class=Forumrow>  
<input type=radio name="SendRegEmail" value=0 <%if cint(Forum_Setting(47))=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="SendRegEmail" value=1 <%if cint(Forum_Setting(47))=1 then%>checked<%end if%>>打开&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>开启短信欢迎新注册用户</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="smsflag" value=0 <%if cint(Forum_Setting(46))=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="smsflag" value=1 <%if cint(Forum_Setting(46))=1 then%>checked<%end if%>>打开&nbsp;
</td>
</tr>
<tr> 
<th height=23 colspan=2 align=left><a name="setting9"><b>投票选项</b></a>[<a href="#top"><font color=#FFFFFF>顶部</font></a>]</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>论坛投票项数目</U><BR>请填写数字</td>
<td width="50%" class=Forumrow>  
<input type="text" name="vote_num" size="3" value="<%=Forum_Setting(39)%>">
</td>
</tr>
<tr> 
<th height=23 colspan=2 align=left><a name="setting10"><b>服务器选项</b></a>[<a href="#top"><font color=#FFFFFF>顶部</font></a>]</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>论坛所在时区</U></td>
<td width="50%" class=Forumrow>  
<input type="text" name="GMT" size="35" value="<%=Forum_info(9)%>">
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>服务器时差</U></td>
<td width="50%" class=Forumrow>  
<select name="TimeAdjust">
<%for i=-23 to 23%>
<option value="<%=i%>" <%if cint(i)=cint(Forum_Setting(0)) then%>selected<%end if%>><%=i%>
<%next%>
</select>
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>脚本超时时间</U><BR>默认为300，一般不做更改</td>
<td width="50%" class=Forumrow>  
<input type="text" name="ScriptTimeOut" size="3" value="<%=Forum_Setting(1)%>">&nbsp;秒
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>是否显示页面执行时间</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="runtime" value=0 <%if cint(Forum_Setting(30))=0 then%>checked<%end if%>>否&nbsp;
<input type=radio name="runtime" value=1 <%if cint(Forum_Setting(30))=1 then%>checked<%end if%>>是&nbsp;
</td>
</tr>
<tr> 
<th height=23 colspan=2 align=left><a name="setting11"><b>搜索和查看资料选项</b></a>[<a href="#top"><font color=#FFFFFF>顶部</font></a>]</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>客人是否允许查看会员资料</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="ViewUser_g" value=0 <%if cint(Forum_Setting(27))=0 then%>checked<%end if%>>否&nbsp;
<input type=radio name="ViewUser_g" value=1 <%if cint(Forum_Setting(27))=1 then%>checked<%end if%>>是&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>用户是否允许查看会员资料</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="ViewUser_u" value=0 <%if cint(Forum_Setting(28))=0 then%>checked<%end if%>>否&nbsp;
<input type=radio name="ViewUser_u" value=1 <%if cint(Forum_Setting(28))=1 then%>checked<%end if%>>是&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>是否允许未登陆用户搜索</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="Search_G" value=0 <%if cint(Forum_Setting(48))=0 then%>checked<%end if%>>允许&nbsp;
<input type=radio name="Search_G" value=1 <%if cint(Forum_Setting(48))=1 then%>checked<%end if%>>不允许&nbsp;
</td>
</tr>
<tr> 
<th height=23 colspan=2 align=left><a name="setting12"><b>在线和用户来源</b></a>[<a href="#top"><font color=#FFFFFF>顶部</font></a>]</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>在线名单显示客人在线</U><BR>为节省资源建议关闭</td>
<td width="50%" class=Forumrow>  
<input type=radio name="online_g" value=0 <%if cint(Forum_Setting(15))=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="online_g" value=1 <%if cint(Forum_Setting(15))=1 then%>checked<%end if%>>打开&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>在线名单显示用户在线</U><BR>为节省资源建议关闭</td>
<td width="50%" class=Forumrow>  
<input type=radio name="online_u" value=0 <%if cint(Forum_Setting(14))=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="online_u" value=1 <%if cint(Forum_Setting(14))=1 then%>checked<%end if%>>打开&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>删除不活动用户时间</U><BR>可设置删除多少分钟内不活动用户<BR>单位：分钟，请输入数字</td>
<td width="50%" class=Forumrow>  
<input type="text" name="kicktime" size="3" value="<%=Forum_Setting(8)%>">&nbsp;分钟
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>允许同时在线数</U><BR>如不想限制，可设置为999999</td>
<td width="50%" class=Forumrow>  
<input type="text" name="online_n" size="6" value="<%=Forum_Setting(26)%>">&nbsp;人
</td>
</tr>
<tr> 
<th height=23 colspan=2 align=left><a name="setting13"><b>邮件选项</b></a>[<a href="#top"><font color=#FFFFFF>顶部</font></a>]</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>发送邮件组件</U><BR>如果您的服务器不支持下列组件，请选择不支持</td>
<td width="50%" class=Forumrow>  
<select name="EmailFlag">
<option value="0" <%if Forum_Setting(2)=0 then%>selected<%end if%>>不支持 
<option value="1" <%if Forum_Setting(2)=1 then%>selected<%end if%>>JMAIL 
<option value="2" <%if Forum_Setting(2)=2 then%>selected<%end if%>>CDONTS 
<option value="3" <%if Forum_Setting(2)=3 then%>selected<%end if%>>ASPEMAIL 
</select>
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>SMTP Server地址</U><BR>只有在论坛使用设置中打开了发送邮件功能，该填写内容方有效</td>
<td width="50%" class=Forumrow>  
<input type="text" name="SMTPServer" size="35" value="<%=Forum_info(4)%>">
</td>
</tr>
<tr> 
<th height=23 colspan=2 align=left><a name="setting14"><b>上传设置</b></a>[<a href="#top"><font color=#FFFFFF>顶部</font></a>]</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>头像上传</U></td>
<td width="50%" class=Forumrow> 
<input type=radio name="uploadFlag" value=0 <%if Forum_Setting(7)=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="uploadFlag" value=1 <%if Forum_Setting(7)=1 then%>checked<%end if%>>打开&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>上传文件类型</U><BR>使用逗号隔开</td>
<td width="50%" class=Forumrow>  
<input type="text" name="Forum_upload" size="35" value="<%=Forum_upload%>">
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>上传文件大小</U><BR>请填写数字</td>
<td width="50%" class=Forumrow>  
<input type="text" name="uploadsize" size="6" value="<%=Forum_Setting(33)%>"> K
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>帖子上传图片</U><BR>用于帖子图片的上传</td>
<td width="50%" class=Forumrow>  
<input type=radio name="Uploadpic" value=0 <%if cint(Forum_Setting(3))=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="Uploadpic" value=1 <%if cint(Forum_Setting(3))=1 then%>checked<%end if%>>打开&nbsp;
</td>
</tr>
<tr> 
<td class=Forumrow width="50%"><U>帖子图片上传路径</U><BR>相对论坛目录路径</td>
<td class=Forumrow width="50%"> 
<input type="text" name="UpLoadPath" size="10" value="<%=Forum_Setting(64)%>">&nbsp;如：uploadFace
</td>
</tr>
<tr> 
<th height=23 colspan=2 align=left><a name="setting15"><b>用户选项</b></a>[<a href="#top"><font color=#FFFFFF>顶部</font></a>]</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>用户签名是否开启UBB代码</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="UserubbCode" value=0 <%if Forum_Setting(65)=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="UserubbCode" value=1 <%if Forum_Setting(65)=1 then%>checked<%end if%>>打开&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>用户签名是否开启HTML代码</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="UserHtmlCode" value=0 <%if Forum_Setting(66)=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="UserHtmlCode" value=1 <%if Forum_Setting(66)=1 then%>checked<%end if%>>打开&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>用户是否开启贴图标签</U><BR>包括图片、flash、多媒体等</td>
<td width="50%" class=Forumrow>  
<input type=radio name="UserImgCode" value=0 <%if Forum_Setting(67)=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="UserImgCode" value=1 <%if Forum_Setting(67)=1 then%>checked<%end if%>>打开&nbsp;
</td>
</tr>
<tr> 
<td class=Forumrow width="50%"><U>用户排行列表个数</U></td>
<td class=Forumrow width="50%"> 
<input type="text" name="TopUserNum" size="6" value="<%=Forum_Setting(68)%>">&nbsp;个
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>用户头衔</U><BR>是否允许用户自定义头衔</td>
<td width="50%" class=Forumrow>  
<input type=radio name="TitleFlag" value=0 <%if cint(Forum_Setting(6))=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="TitleFlag" value=1 <%if cint(Forum_Setting(6))=1 then%>checked<%end if%>>打开&nbsp;
</td>
</tr><tr> 
<td width="50%" class=Forumrow> <U>用户可删除和锁定自己发布帖子</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="UserPostAdmin" value=0 <%if Forum_Setting(70)=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="UserPostAdmin" value=1 <%if Forum_Setting(70)=1 then%>checked<%end if%>>开放&nbsp;
</td>
</tr>
<tr> 
<th height=23 colspan=2 align=left><a name="setting16"><b>浏览帖子选项</b></a>[<a href="#top"><font color=#FFFFFF>顶部</font></a>]</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>论坛风格</U><BR>讨论区风格为树形</td>
<td width="50%" class=Forumrow>  
<input type=radio name="DvbbsSkin" value=0 <%if Forum_Setting(62)=0 then%>checked<%end if%>>默认风格&nbsp;
<input type=radio name="DvbbsSkin" value=1 <%if Forum_Setting(62)=1 then%>checked<%end if%>>讨论区风格&nbsp;
</td>
</tr>
<tr> 
<td class=Forumrow width="50%"><U>讨论区风格预览字数</U><BR>取消预览请填写0</td>
<td class=Forumrow width="50%"> 
<input type="text" name="SkinFontNum" size="6" value="<%=Forum_Setting(63)%>">&nbsp;字
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>浏览贴子每页显示贴子数</U></td>
<td width="50%" class=Forumrow>  
<input type="text" name="Maxtitlelist" size="3" value="<%=Forum_Setting(12)%>">&nbsp;条
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>帖子正文字号</U></td>
<td width="50%" class=Forumrow>  
<input type="text" name="FontSize" size="3" value="<%=Forum_Setting(59)%>">&nbsp;Pt，如15
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>帖子正文行距</U></td>
<td width="50%" class=Forumrow>  
<input type="text" name="FontHeight" size="3" value="<%=Forum_Setting(60)%>">&nbsp;倍行距，如1.5
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>浏览帖子页面用户资料显示</U><BR>一为不显示经验财产等<BR>二为不显示用户资料<BR>三为两者都不显示，四为两者都显示</td>
<td width="50%" class=Forumrow>  
<input type=radio name="BbsUserInfo" value=1 <%if Forum_Setting(61)=1 then%>checked<%end if%>>模式一&nbsp;
<input type=radio name="BbsUserInfo" value=2 <%if Forum_Setting(61)=2 then%>checked<%end if%>>模式二&nbsp;
<input type=radio name="BbsUserInfo" value=3 <%if Forum_Setting(61)=3 then%>checked<%end if%>>模式三&nbsp;
<input type=radio name="BbsUserInfo" value=4 <%if Forum_Setting(61)=4 then%>checked<%end if%>>模式四&nbsp;
</td>
</tr>
<tr> 
<th height=23 colspan=2 align=left><a name="setting17"><b>防刷新机制</b></a>[<a href="#top"><font color=#FFFFFF>顶部</font></a>]</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>防刷新机制</U><BR>如选择打开请填写下面的限制刷新时间<BR>对版主和管理员无效</td>
<td width="50%" class=Forumrow>  
<input type=radio name="ReflashFlag" value=0 <%if cint(Forum_Setting(19))=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="ReflashFlag" value=1 <%if cint(Forum_Setting(19))=1 then%>checked<%end if%>>打开&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>浏览刷新时间间隔</U><BR>填写该项目请确认您打开了防刷新机制<BR>仅对帖子列表和显示帖子页面起作用</td>
<td width="50%" class=Forumrow>  
<input type="text" name="ReflashTime" size="3" value="<%=Forum_Setting(20)%>">&nbsp;秒
</td>
</tr>
<tr> 
<th height=23 colspan=2 align=left><a name="setting18"><b>版面设置</b></a>[<a href="#top"><font color=#FFFFFF>顶部</font></a>]</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>客人浏览论坛</U><BR>可设置是否允许客人浏览论坛</td>
<td width="50%" class=Forumrow>  
<input type=radio name="guestlogin" value=0 <%if cint(Forum_Setting(9))=0 then%>checked<%end if%>>允许&nbsp;
<input type=radio name="guestlogin" value=1 <%if cint(Forum_Setting(9))=1 then%>checked<%end if%>>不允许&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>每页显示最多纪录</U><BR>用于论坛所有和分页有关的项目</td>
<td width="50%" class=Forumrow>  
<input type="text" name="MaxAnnouncePerPage" size="3" value="<%=Forum_Setting(11)%>">&nbsp;条
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>论坛小字报</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="SmallPaper" value=0 <%if Forum_Setting(56)=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="SmallPaper" value=1 <%if Forum_Setting(56)=1 then%>checked<%end if%>>开启&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>小字报是否对客人开放</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="SmallPaper_g" value=0 <%if Forum_Setting(57)=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="SmallPaper_g" value=1 <%if Forum_Setting(57)=1 then%>checked<%end if%>>开启&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>发布小字报扣除金钱数</U><br>如选择客人可发表，则只扣除会员金钱</td>
<td width="50%" class=Forumrow> 
<input type="text" name="SmallPaper_m" size="3" value="<%=Forum_Setting(58)%>">
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>主版主可增删副版主</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="bmflag_1" value=0 <%if cint(Forum_Setting(49))=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="bmflag_1" value=1 <%if cint(Forum_Setting(49))=1 then%>checked<%end if%>>打开&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>主版主可修改颜色配置</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="bmflag_2" value=0 <%if cint(Forum_Setting(50))=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="bmflag_2" value=1 <%if cint(Forum_Setting(50))=1 then%>checked<%end if%>>打开&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>所有版主均可修改颜色配置</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="bmflag_4" value=0 <%if cint(Forum_Setting(52))=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="bmflag_4" value=1 <%if cint(Forum_Setting(52))=1 then%>checked<%end if%>>打开&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>版面配色公开</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="viewcolor" value=0 <%if cint(Forum_Setting(55))=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="viewcolor" value=1 <%if cint(Forum_Setting(55))=1 then%>checked<%end if%>>打开&nbsp;
</td>
</tr><tr> 
<td width="50%" class=Forumrow> <U>是否公开版面事件</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="bbsEven" value=0 <%if Forum_Setting(71)=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="bbsEven" value=1 <%if Forum_Setting(71)=1 then%>checked<%end if%>>开放&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>是否公开版面事件中的操作人</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="bbsEvenView" value=0 <%if Forum_Setting(72)=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="bbsEvenView" value=1 <%if Forum_Setting(72)=1 then%>checked<%end if%>>开放&nbsp;
</td>
</tr>
<tr> 
<th height=23 colspan=2 align=left><a name="setting19"><b>门派设置</b></a>[<a href="#top"><font color=#FFFFFF>顶部</font></a>]</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>是否开启论坛门派</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="GroupFlag" value=0 <%if cint(Forum_Setting(32))=0 then%>checked<%end if%>>否&nbsp;
<input type=radio name="GroupFlag" value=1 <%if cint(Forum_Setting(32))=1 then%>checked<%end if%>>是&nbsp;
</td>
</tr>



<tr bgcolor=<%=Forum_body(3)%>> 
<td width="50%" class=Forumrow> &nbsp;</td>
<td width="50%" class=Forumrow>  
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
dim Forum_copyright
if trim(request("skinid"))="" and trim(request("boardid"))="" then
Founderr=true
Errmsg=Errmsg+"<br>"+"<li>请选择保存的模板名称"
else
Forum_Setting=request.Form("TimeAdjust") & "," & request.Form("ScriptTimeOut") & "," & request.Form("EmailFlag") & "," & request.Form("Uploadpic") & "," & request.Form("IpFlag") & "," & request.Form("FromFlag") & "," & request.Form("TitleFlag") & "," & request.Form("uploadFlag") & "," & request.Form("kicktime") & "," & request.Form("guestlogin") & "," & request.Form("openmsg") & "," & request.Form("MaxAnnouncePerPage") & "," & request.Form("Maxtitlelist") & "," & request.Form("AnnounceMaxBytes") & "," & request.Form("online_u") & "," & request.Form("online_g") & "," & request.Form("LinkFlag") & "," & request.Form("TopicFlag") & "," & request.Form("VoteFlag") & "," & request.Form("ReflashFlag") & "," & request.Form("ReflashTime") & "," & request.Form("ForumStop") & "," & request.Form("RegTime") & "," & request.Form("EmailReg") & "," & request.Form("EmailRegOne") & "," & request.Form("RegFlag") & "," & request.Form("online_n") & "," & request.Form("ViewUser_g") & "," & request.Form("ViewUser_u") & "," & request.Form("BirthFlag") & "," & request.Form("runtime") & "," & request.Form("FastLogin") & "," & request.Form("GroupFlag") & "," & request.Form("uploadsize") & "," & request.Form("strAllowForumCode") & "," & request.Form("strAllowHTML") & "," & request.Form("strIMGInPosts") & "," & request.Form("strIcons") & "," & request.Form("strflash") & "," & request.Form("vote_num") & "," & request.Form("facenum") & "," & request.Form("imgnum") & "," & request.Form("relaypost") & "," & request.Form("relayposttime") & "," & request.Form("facename") & "," & request.Form("imgname") & "," & request.Form("smsflag") & "," & request.Form("SendRegEmail") & "," & request.Form("Search_G") & "," & request.Form("bmflag_1") & "," & request.Form("bmflag_2") & "," & request.Form("bmflag_3") & "," & request.Form("bmflag_4") & "," & request.Form("bmflag_5") & "," & request.Form("RegFaceNum") & "," & request.Form("viewcolor") & "," & request.Form("SmallPaper") & "," & request.Form("SmallPaper_g") & "," & request.Form("SmallPaper_m")
Forum_Setting=Forum_Setting & "," & request.Form("FontSize") & "," & request.Form("FontHeight") & "," & request.Form("BbsUserInfo") & "," & request.Form("DvbbsSkin") & "," & request.Form("SkinFontNum") & "," & request.Form("UpLoadPath") & "," & request.Form("UserubbCode") & "," & request.Form("UserHtmlCode") & "," & request.Form("UserImgCode") & "," & request.Form("TopUserNum") & "," & request.Form("PostRetrun") & "," & request.Form("UserPostAdmin") & "," & request.Form("bbsEven") & "," & request.Form("bbsEvenView")

Forum_info=Request("ForumName") & "," & Request("ForumURL") & "," & Request("CompanyName") & "," & Request("HostUrl") & "," & Request("SMTPServer") & "," & Request("SystemEmail") & "," & Request("Logo") & "," & Request("Picurl") & "," & Request("Faceurl") & "," & Request("GMT")
Forum_copyright=request("copyright")
if request("skinid")<>"" then
sql="update config set Forum_Setting='"&Forum_Setting&"',StopReadme='"&request.Form("StopReadme")&"',Forum_upload='"&request.Form("Forum_upload")&"',Forum_info='"&Forum_info&"',Forum_copyright='"&Forum_copyright&"' where id in ( "&request("skinid")&" )"
conn.execute(sql)
end if
response.write "设置成功。"
end if
end sub
%>