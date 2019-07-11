<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<%
stats="帖子管理"
dim replyid
dim id
dim Lasttopic,Lastpost
dim lastrootid,lastpostuser
dim ip,url
dim title,content
dim TotalUseTable
dim Topic,TopicUsername,TopicUserID
dim canlocktopic,candeltopic,canmovetopic,cantoptopic,canbesttopic,canawardtopic
canlocktopic=false
candeltopic=false
canmovetopic=false
cantoptopic=false
canbesttopic=false
canawardtopic=false
ip=Request.ServerVariables("REMOTE_ADDR")
if not founduser then
	founderr=true
	Errmsg=Errmsg+"<br>"+"<li>请登陆后进行操作。"
end if
if request("boardid")="" then
	founderr=true
	Errmsg=Errmsg+"<br>"+"<li>请指定论坛版面。"
elseif not isInteger(request("boardid")) then
	founderr=true
	Errmsg=Errmsg+"<br>"+"<li>非法的版面参数。"
else
	boardid=clng(boardid)
end if
if request("id")="" then
	founderr=true
	Errmsg=Errmsg+"<br>"+"<li>请指定相关贴子。"
elseif not isInteger(request("id")) then
	founderr=true
	Errmsg=Errmsg+"<br>"+"<li>非法的贴子参数。"
else
	id=request("id")
end if
if isInteger(request("replyid")) then
	replyid=request("replyid")
end if

dim doWealth,douserEP,douserCP
dim doWealthMsg,douserEPMsg,douserCPMsg,allMsg
if not isnumeric(request("doWealth")) or request("doWealth")="0" or request("doWealth")="" then
	doWealth=0
	doWealthMsg=""
else
	doWealth=request("doWealth")
	doWealthMsg="金钱" & request("doWealth") & "，"
end if
if not isnumeric(request("douserEP")) or request("douserEP")="0" or request("douserEP")="" then
	douserEP=0
	douserEPMsg=""
else
	douserEP=request("douserEP")
	douserEPMsg="经验" & request("douserEP") & "，"
end if
if not isnumeric(request("douserCP")) or request("douserCP")="0" or request("douserCP")="" then
	douserCP=0
	douserCPMsg=""
else
	douserCP=request("douserCP")
	douserCPMsg="魅力" & request("douserCP")
end if
if doWealthMsg="" and douserEPMsg="" and douserCPMsg="" then
	allmsg="没有对用户进行分值操作"
else
	allmsg="用户操作：" & doWealthMsg & douserEPMsg & douserCPMsg
end if
if not founderr then
	set rs=conn.execute("select title,postusername,postuserid,PostTable from topic where topicid="&id)
	if rs.eof and rs.bof then
		Errmsg=Errmsg+"<br><li>没有找到相关帖子！"
		founderr=true
	else
		Topic=rs(0)
		Topicusername=rs(1)
		Topicuserid=rs(2)
		TotalUseTable=rs(3)
	end if
	'判断用户是否有锁定/解除锁定权限
	if (master or superboardmaster or boardmaster) and Cint(GroupSetting(20))=1 then
		canlocktopic=true
	else
		canlocktopic=false
	end if
	if Cint(GroupSetting(20))=1 and UserGroupID>3 then
		canlocktopic=true
	end if
	if (Cint(GroupSetting(13))=1 and TopicUsername=membername) then
		canlocktopic=true
	end if
	if FoundUserPer and Cint(GroupSetting(13))=1 and TopicUsername=membername then
		canlocktopic=true
	elseif FoundUserPer and Cint(GroupSetting(13))=0 and TopicUsername=membername then
		canlocktopic=false
	end if
	if FoundUserPer and Cint(GroupSetting(20))=1 and TopicUsername<>membername then
		canlocktopic=true
	elseif FoundUserPer and Cint(GroupSetting(20))=0 and TopicUsername<>membername then
		canlocktopic=false
	end if
	'判断用户是否有移动帖子权限
	if (master or superboardmaster or boardmaster) and Cint(GroupSetting(19))=1 then
		canmovetopic=true
	else
		canmovetopic=false
	end if
	if Cint(GroupSetting(19))=1 and UserGroupID>3 then
		canlocktopic=true
	end if
	if (Cint(GroupSetting(12))=1 and TopicUsername=membername) then
		canmovetopic=true
	end if
	if FoundUserPer and Cint(GroupSetting(12))=1 and TopicUsername=membername then
		canmovetopic=true
	elseif FoundUserPer and Cint(GroupSetting(12))=0 and TopicUsername=membername then
		canmovetopic=false
	end if
	if FoundUserPer and Cint(GroupSetting(19))=1 and TopicUsername<>membername then
		canmovetopic=true
	elseif FoundUserPer and Cint(GroupSetting(19))=0 and TopicUsername<>membername then
		canmovetopic=false
	end if
	'判断用户是否有删除帖子权限
	if (master or superboardmaster or boardmaster) and Cint(GroupSetting(18))=1 then
		candeltopic=true
	else
		candeltopic=false
	end if
	if Cint(GroupSetting(18))=1 and UserGroupID>3 then
		canlocktopic=true
	end if
	if (Cint(GroupSetting(11))=1 and TopicUsername=membername) then
		candeltopic=true
	end if
	if FoundUserPer and Cint(GroupSetting(11))=1 and TopicUsername=membername then
		candeltopic=true
	elseif FoundUserPer and Cint(GroupSetting(11))=0 and TopicUsername=membername then
		candeltopic=false
	end if
	if FoundUserPer and Cint(GroupSetting(18))=1 and TopicUsername<>membername then
		candeltopic=true
	elseif FoundUserPer and Cint(GroupSetting(18))=0 and TopicUsername<>membername then
		candeltopic=false
	end if
	'判断用户是否有固顶/解除固顶帖子权限
	if (master or superboardmaster or boardmaster) and Cint(GroupSetting(21))=1 then
		cantoptopic=true
	else
		cantoptopic=false
	end if
	if Cint(GroupSetting(21))=1 and UserGroupID>3 then
		canlocktopic=true
	end if
	if FoundUserPer and Cint(GroupSetting(21))=1 then
		cantoptopic=true
	elseif FoundUserPer and Cint(GroupSetting(21))=0 then
		cantoptopic=false
	end if
	'判断用户是否有奖励/惩罚帖子权限
	if (master or superboardmaster or boardmaster) and Cint(GroupSetting(22))=1 then
		canawardtopic=true
	else
		canawardtopic=false
	end if
	if Cint(GroupSetting(22))=1 and UserGroupID>3 then
		canlocktopic=true
	end if
	if FoundUserPer and Cint(GroupSetting(22))=1 then
		canawardtopic=true
	elseif FoundUserPer and Cint(GroupSetting(22))=0 then
		canawardtopic=false
	end if
	'判断用户是否有加入/解除精华帖子权限
	if (master or superboardmaster or boardmaster) and Cint(GroupSetting(24))=1 then
		canbesttopic=true
	else
		canbesttopic=false
	end if
	if Cint(GroupSetting(24))=1 and UserGroupID>3 then
		canlocktopic=true
	end if
	if FoundUserPer and Cint(GroupSetting(24))=1 then
		canbesttopic=true
	elseif FoundUserPer and Cint(GroupSetting(24))=0 then
		canbesttopic=false
	end if
end if
if founderr then
	call nav()
	call head_var(1)
	call dvbbs_error()
else
	call nav()
	call head_var(2)
	select case request("action")
	case "lock"
		if not canlocktopic then
			Errmsg=Errmsg+"<br><li>您没有执行此操作的权限。"
			founderr=true
		else
			call lock()
		end if
	case "unlock"
		if not canlocktopic then
			Errmsg=Errmsg+"<br><li>您没有执行此操作的权限。"
			founderr=true
		else
			call unlock()
		end if
	case "delet"
		if not candeltopic then
			Errmsg=Errmsg+"<br><li>您没有执行此操作的权限。"
			founderr=true
		else
			call delete()
		end if
	case "move"
		if not canmovetopic then
			Errmsg=Errmsg+"<br><li>您没有执行此操作的权限。"
			founderr=true
		else
			call Tmove()
		end if
	case "copy"
		call copy()
	case "istop"
		if not cantoptopic then
			Errmsg=Errmsg+"<br><li>您没有执行此操作的权限。"
			founderr=true
		else
			call istop()
		end if
	case "notop"
		if not cantoptopic then
			Errmsg=Errmsg+"<br><li>您没有执行此操作的权限。"
			founderr=true
		else
			call notop()
		end if
	case "dele"
		call dele()
	case "isbest"
		if not canbesttopic then
			Errmsg=Errmsg+"<br><li>您没有执行此操作的权限。"
			founderr=true
		else
			call isbest()
		end if
	case "nobest"
		if not canbesttopic then
			Errmsg=Errmsg+"<br><li>您没有执行此操作的权限。"
			founderr=true
		else
			call nobest()
		end if
	case "award"
		if not canawardtopic then
			Errmsg=Errmsg+"<br><li>您没有执行此操作的权限。"
			founderr=true
		else
			call award()
		end if
	case "punish"
		if not canawardtopic then
			Errmsg=Errmsg+"<br><li>您没有执行此操作的权限。"
			founderr=true
		else
			call punish()
		end if
	case else
		call main()
	end select
	if founderr then call dvbbs_error()
end if
call activeonline()
call footer()
	
sub main()
dim doWealth,douserEP,douserCP
dim seldisable,reaction
dim postusername
select case request("action")
case "锁定"
	doWealth=0
	douserEP=0
	douserCP=0
	seldisable=""
	reaction="lock"
	if not canlocktopic then
		Errmsg=Errmsg+"<br><li>您没有执行此操作的权限。"
		founderr=true
		exit sub
	end if
case "解锁"
	doWealth=0
	douserEP=0
	douserCP=0
	seldisable=""
	reaction="unlock"
	if not canlocktopic then
		Errmsg=Errmsg+"<br><li>您没有执行此操作的权限。"
		founderr=true
		exit sub
	end if
case "删除主题"
	doWealth=-Forum_user(3)
	douserEP=-Forum_user(8)
	douserCP=-Forum_user(13)
	seldisable=""
	reaction="delet"
	if not candeltopic then
		Errmsg=Errmsg+"<br><li>您没有执行此操作的权限。"
		founderr=true
		exit sub
	end if
case "删除跟帖"
	doWealth=-Forum_user(3)
	douserEP=-Forum_user(8)
	douserCP=-Forum_user(13)
	seldisable=""
	reaction="dele"
	set rs=conn.execute("select topic,username,postuserid from "&TotalUseTable&" where Announceid="&replyid)
	if rs.eof and rs.bof then
		Errmsg=Errmsg+"<br><li>没有找到相关跟贴！"
		founderr=true
		exit sub
	end if
	Topic=rs(0)
	TopicUsername=rs(1)
	TopicUserID=rs(2)
	'判断用户是否有删除帖子权限
	if (Cint(GroupSetting(11))=1 and TopicUsername=membername) then
		candeltopic=true
	end if
	if FoundUserPer and Cint(GroupSetting(11))=1 and TopicUsername=membername then
		candeltopic=true
	elseif FoundUserPer and Cint(GroupSetting(11))=0 and TopicUsername=membername then
		candeltopic=false
	end if
	if FoundUserPer and Cint(GroupSetting(18))=1 and TopicUsername<>membername then
		candeltopic=true
	elseif FoundUserPer and Cint(GroupSetting(18))=0 and TopicUsername<>membername then
		candeltopic=false
	end if
	if not candeltopic then
		Errmsg=Errmsg+"<br><li>您没有执行此操作的权限。"
		founderr=true
		exit sub
	end if
case "精华"
	doWealth=Forum_user(15)
	douserEP=Forum_user(17)
	douserCP=Forum_user(16)
	seldisable=""
	reaction="isbest"
	set rs=conn.execute("select topic,username,postuserid from "&TotalUseTable&" where Announceid="&replyid)
	if rs.eof and rs.bof then
		Errmsg=Errmsg+"<br><li>没有找到相关跟贴！"
		founderr=true
		exit sub
	end if
	Topic=rs(0)
	TopicUsername=rs(1)
	TopicUserID=rs(2)
	if not canbesttopic then
		Errmsg=Errmsg+"<br><li>您没有执行此操作的权限。"
		founderr=true
		exit sub
	end if
case "解除精华"
	doWealth=-Forum_user(15)
	douserEP=-Forum_user(17)
	douserCP=-Forum_user(16)
	seldisable=""
	reaction="nobest"
	set rs=conn.execute("select topic,username,postuserid from "&TotalUseTable&" where Announceid="&replyid)
	if rs.eof and rs.bof then
		Errmsg=Errmsg+"<br><li>没有找到相关跟贴！"
		founderr=true
		exit sub
	end if
	Topic=rs(0)
	TopicUsername=rs(1)
	TopicUserID=rs(2)
	if not canbesttopic then
		Errmsg=Errmsg+"<br><li>您没有执行此操作的权限。"
		founderr=true
		exit sub
	end if
case "复制"
	doWealth=0
	douserEP=0
	douserCP=0
	seldisable="disabled"
	reaction="copy"
	set rs=conn.execute("select topic,username,postuserid from "&TotalUseTable&" where Announceid="&replyid)
	if rs.eof and rs.bof then
		Errmsg=Errmsg+"<br><li>没有找到相关跟贴！"
		founderr=true
		exit sub
	end if
	Topic=rs(0)
	TopicUsername=rs(1)
	TopicUserID=rs(2)
	'判断用户是否有移动帖子权限
	if (Cint(GroupSetting(12))=1 and TopicUsername=membername) then
		canmovetopic=true
	end if
	if FoundUserPer and Cint(GroupSetting(12))=1 and TopicUsername=membername then
		canmovetopic=true
	elseif FoundUserPer and Cint(GroupSetting(12))=0 and TopicUsername=membername then
		canmovetopic=false
	end if
	if FoundUserPer and Cint(GroupSetting(19))=1 and TopicUsername<>membername then
		canmovetopic=true
	elseif FoundUserPer and Cint(GroupSetting(19))=0 and TopicUsername<>membername then
		canmovetopic=false
	end if
	if not canmovetopic then
		Errmsg=Errmsg+"<br><li>您没有执行此操作的权限。"
		founderr=true
		exit sub
	end if
case "固顶"
	doWealth=0
	douserEP=0
	douserCP=0
	seldisable=""
	reaction="istop"
	if not cantoptopic then
		Errmsg=Errmsg+"<br><li>您没有执行此操作的权限。"
		founderr=true
		exit sub
	end if
case "解固"
	doWealth=0
	douserEP=0
	douserCP=0
	seldisable=""
	reaction="notop"
	if not cantoptopic then
		Errmsg=Errmsg+"<br><li>您没有执行此操作的权限。"
		founderr=true
		exit sub
	end if
case "移动"
	doWealth=0
	douserEP=0
	douserCP=0
	seldisable="disabled"
	reaction="move"
	if not canmovetopic then
		Errmsg=Errmsg+"<br><li>您没有执行此操作的权限。"
		founderr=true
		exit sub
	end if
case "奖励"
	doWealth=5
	douserEP=1
	douserCP=2
	seldisable=""
	reaction="award"
	if not canawardtopic then
		Errmsg=Errmsg+"<br><li>您没有执行此操作的权限。"
		founderr=true
		exit sub
	end if
case "惩罚"
	doWealth=-5
	douserEP=-1
	douserCP=-2
	seldisable=""
	reaction="punish"
	if not canawardtopic then
		Errmsg=Errmsg+"<br><li>您没有执行此操作的权限。"
		founderr=true
		exit sub
	end if
case else
	Errmsg=Errmsg+"<br><li>请指定相关参数！"
	founderr=true
	exit sub
end select
%>
<FORM METHOD=POST ACTION="admin_postings.asp?action=<%=reaction%>">
<table style="width:70%" cellspacing="1" cellpadding="3" align="center" class=tableborder1>
  <tr> 
    <th height=24>论坛帖子管理中心－－您要进行的操作是<%=request("action")%></th>
  </tr>   
  <tr> 
    <td class=tablebody1 height=24><b>
      操作理由</b>：  
	  <select name="title" size=1>
<option value="">自定义</option>
<option value="灌水">灌水</option>
<option value="广告">广告</option>
<option value="奖励">奖励</option>
<option value="好文章">好文章</option>
<option value="内容不符">内容不符</option>
<option value="重复发贴">重复发帖</option>
	  </select>
	  <input type="text" name="content" size=60>  *</td>
  </tr>   
  <tr> 
    <td class=tablebody1 height=24><b>
      用户操作</b>：  金钱
	<select name="doWealth" size=1 <%=seldisable%>>

<%for i=-50 to 50%>
<option value="<%=i%>" <%if cint(doWealth)=cint(i) then%>selected<%end if%>><%=i%></option>
<%next%>
	</select>&nbsp;魅力
	<select name="douserCP" size=1 <%=seldisable%>>

<%for i=-50 to 50%>
<option value="<%=i%>" <%if cint(dousercp)=cint(i) then%>selected<%end if%>><%=i%></option>
<%next%>
	</select>&nbsp;经验
	<select name="douserEP" size=1 <%=seldisable%>>

<%for i=-50 to 50%>
<option value="<%=i%>" <%if cint(douserep)=cint(i) then%>selected<%end if%>><%=i%></option>
<%next%>
	</select>
  *</td>
  </tr> 
<input type=hidden value="<%=id%>" name="id">
<input type=hidden value="<%=replyid%>" name="replyid">
<input type=hidden value="<%=boardid%>" name="boardid">
  <tr> 
    <td height=24 class=tablebody2>
      请慎重使用版主的管理职能，版主所有操作将被记录&nbsp;<input type="submit" name=submit value="确认操作"></td>
  </tr>   
</table>
</FORM>
<%
	set rs=nothing
end sub

sub award()
title=CheckStr(htmlencode(request.form("title")))
content=CheckStr(htmlencode(request.form("content")))
content="原因：" & title & content
if TopicUserName=membername then
	Errmsg=Errmsg+"<br><li>版主和管理员不能对自己的帖子实行奖励操作。"
	founderr=true
	exit sub
end if
if htmlencode(request.form("title"))="" and htmlencode(request.form("content"))="" then
	Errmsg=Errmsg+"<br><li>请写明操作原因。"
	founderr=true
	exit sub
end if
sql="insert into log (l_announceid,l_boardid,l_touser,l_username,l_content,l_ip) values ("&id&","&boardid&",'"&TopicUsername&"','"&membername&"','奖励用户《"&topic&"》，"&content& "，"&allmsg&"','"&ip&"')"
conn.execute(sql)
if allmsg<>"" then
	conn.execute("update [user] set userWealth=userWealth+"&doWealth&",userCP=userCP+"&douserCP&",userEP=userEP+"&douserEP&" where userid="&TopicUserID)
end if
sucmsg="奖励用户《"&topic&"》，"&content& "，"&allmsg&""
call dvbbs_suc()
end sub

sub punish()
title=CheckStr(htmlencode(request.form("title")))
content=CheckStr(htmlencode(request.form("content")))
content="原因：" & title & content
if htmlencode(request.form("title"))="" and htmlencode(request.form("content"))="" then
	Errmsg=Errmsg+"<br><li>请写明操作原因。"
	founderr=true
	exit sub
end if
sql="insert into log (l_announceid,l_boardid,l_touser,l_username,l_content,l_ip) values ("&id&","&boardid&",'"&TopicUsername&"','"&membername&"','惩罚用户《"&topic&"》，"&content& "，"&allmsg&"','"&ip&"')"
conn.execute(sql)
if allmsg<>"" then
	conn.execute("update [user] set userWealth=userWealth+"&doWealth&",userCP=userCP+"&douserCP&",userEP=userEP+"&douserEP&" where userid="&TopicUserID)
end if
sucmsg="惩罚用户《"&topic&"》，"&content& "，"&allmsg&""
call dvbbs_suc()
end sub

sub lock()
title=CheckStr(htmlencode(request.form("title")))
content=CheckStr(htmlencode(request.form("content")))
content="原因：" & title & content
if htmlencode(request.form("title"))="" and htmlencode(request.form("content"))="" then
	Errmsg=Errmsg+"<br><li>请写明操作原因。"
	founderr=true
	exit sub
end if
sql="update topic set locktopic=1 where boardid="&boardid&" and topicid="&id
conn.Execute(sql)
sql="insert into log (l_announceid,l_boardid,l_touser,l_username,l_content,l_ip) values ("&id&","&boardid&",'"&TopicUsername&"','"&membername&"','锁定帖子《"&topic&"》，"&content& "，"&allmsg&"','"&ip&"')"
conn.execute(sql)
if allmsg<>"" then
	conn.execute("update [user] set userWealth=userWealth+"&doWealth&",userCP=userCP+"&douserCP&",userEP=userEP+"&douserEP&" where userid="&TopicUserID)
end if
sucmsg="锁定帖子《"&topic&"》，"&content& "，"&allmsg&""
call dvbbs_suc()
end sub

sub unlock()
title=CheckStr(htmlencode(request.form("title")))
content=CheckStr(htmlencode(request.form("content")))
content="原因：" & title & content
if htmlencode(request.form("title"))="" and htmlencode(request.form("content"))="" then
	Errmsg=Errmsg+"<br><li>请写明操作原因。"
	founderr=true
	exit sub
end if
sql="update topic set locktopic=0 where boardid="&boardid&" and topicid="&id
conn.Execute(sql)

sql="insert into log (l_announceid,l_boardid,l_touser,l_username,l_content,l_ip) values ("&id&","&boardid&",'"&TopicUsername&"','"&membername&"','解除锁定《"&topic&"》，"&content& "，"&allmsg&"','"&ip&"')"
conn.execute(sql)
if allmsg<>"" then
	conn.execute("update [user] set userWealth=userWealth+"&doWealth&",userCP=userCP+"&douserCP&",userEP=userEP+"&douserEP&" where userid="&TopicUserID)
end if
sucmsg="解除锁定《"&topic&"》，"&content& "，"&allmsg&""
call dvbbs_suc()
end sub

sub istop()
title=CheckStr(htmlencode(request.form("title")))
content=CheckStr(htmlencode(request.form("content")))
content="原因：" & title & content
if htmlencode(request.form("title"))="" and htmlencode(request.form("content"))="" then
	Errmsg=Errmsg+"<br><li>请写明操作原因。"
	founderr=true
	exit sub
end if
	
sql="update topic set istop=1 where boardid="&boardid&" and topicid="&id
conn.Execute(sql)

sql="insert into log (l_announceid,l_boardid,l_touser,l_username,l_content,l_ip) values ("&id&","&boardid&",'"&TopicUsername&"','"&membername&"','固顶帖子《"&topic&"》，"&content& "，"&allmsg&"','"&ip&"')"
conn.execute(sql)
if allmsg<>"" then
	conn.execute("update [user] set userWealth=userWealth+"&doWealth&",userCP=userCP+"&douserCP&",userEP=userEP+"&douserEP&" where userid="&TopicUserID)
end if
sucmsg="固顶帖子《"&topic&"》，"&content& "，"&allmsg&""
call dvbbs_suc()
end sub

sub notop()
title=CheckStr(htmlencode(request.form("title")))
content=CheckStr(htmlencode(request.form("content")))
content="原因：" & title & content
if htmlencode(request.form("title"))="" and htmlencode(request.form("content"))="" then
	Errmsg=Errmsg+"<br><li>请写明操作原因。"
	founderr=true
	exit sub
end if
sql="update topic set istop=0 where boardid="&boardid&" and topicid="&id
conn.Execute(sql)

sql="insert into log (l_announceid,l_boardid,l_touser,l_username,l_content,l_ip) values ("&id&","&boardid&",'"&topicusername&"','"&membername&"','解除固顶《"&topic&"》，"&content& "，"&allmsg&"','"&ip&"')"
conn.execute(sql)
if allmsg<>"" then
	conn.execute("update [user] set userWealth=userWealth+"&doWealth&",userCP=userCP+"&douserCP&",userEP=userEP+"&douserEP&" where userid="&topicuserid)
end if
sucmsg="解除固顶《"&topic&"》，"&content& "，"&allmsg&""
call dvbbs_suc()
end sub

sub isbest()
dim datetimestr
set rs=conn.execute("select * from "&TotalUseTable&" where Announceid="&replyid)
if rs.eof and rs.bof then
	Errmsg=Errmsg+"<br><li>没有找到相关帖子！"
	founderr=true
	exit sub
end if
topic=rs("topic")
topicusername=rs("username")
topicuserid=rs("postuserid")
if topic="" then
	topic=left(replace(replace(rs("body"),"'",""),chr(10),","),26)
else
	topic=replace(topic,"'","")
end if
datetimestr=replace(replace(rs("dateandtime"),"上午",""),"下午","")

title=CheckStr(htmlencode(request.form("title")))
content=CheckStr(htmlencode(request.form("content")))
content="原因：" & title & content
if htmlencode(request.form("title"))="" and htmlencode(request.form("content"))="" then
	Errmsg=Errmsg+"<br><li>请写明操作原因。"
	founderr=true
	exit sub
end if
sql="update "&TotalUseTable&" set isbest=1 where boardid="&boardid&" and announceid="&cstr(replyid)
conn.Execute(sql)
conn.execute("update topic set isbest=1 where boardid="&boardid&" and topicid="&id)

sql="insert into bestTopic (title,boardid,Announceid,rootid,postusername,postuserid,dateandtime,expression) values ('"&topic&"',"&rs("boardid")&","&rs("Announceid")&","&rs("rootid")&",'"&topicusername&"',"&rs("postuserid")&",'"&datetimestr&"','"&rs("expression")&"')"
conn.execute(sql)

sql="insert into log (l_announceid,l_boardid,l_touser,l_username,l_content,l_ip) values ("&id&","&boardid&",'"&topicusername&"','"&membername&"','加入精华《"&topic&"》，"&content& "，"&allmsg&"','"&ip&"')"
conn.execute(sql)
conn.execute("update [user] set userWealth=userWealth+"&doWealth&",userCP=userCP+"&douserCP&",userEP=userEP+"&douserEP&",userIsBest=userisBest+1 where userid="&topicuserid)
sucmsg="加入精华《"&topic&"》，"&content& "，"&allmsg&""
call dvbbs_suc()
set rs=nothing
end sub

sub nobest()
set rs=conn.execute("select topic,username,postuserid from "&TotalUseTable&" where Announceid="&replyid)
if rs.eof and rs.bof then
	Errmsg=Errmsg+"<br><li>没有找到相关帖子！"
	founderr=true
	exit sub
else
	topic=rs(0)
	topicusername=rs(1)
	topicuserid=rs(2)
	if topic="" then
		topic="本帖子为回复帖子"
	end if
end if
set rs=nothing
title=CheckStr(htmlencode(request.form("title")))
content=CheckStr(htmlencode(request.form("content")))
content="原因：" & title & content
if htmlencode(request.form("title"))="" and htmlencode(request.form("content"))="" then
	Errmsg=Errmsg+"<br><li>请写明操作原因。"
	founderr=true
	exit sub
end if

sql="update "&TotalUseTable&" set isbest=0 where boardid="&boardid&" and announceid="&replyid
conn.Execute(sql)
conn.execute("update topic set isbest=0 where boardid="&boardid&" and topicid="&id)

conn.execute("delete from besttopic where Announceid="&replyID)

sql="insert into log (l_announceid,l_boardid,l_touser,l_username,l_content,l_ip) values ("&id&","&boardid&",'"&topicusername&"','"&membername&"','解除精华《"&topic&"》，"&content& "，"&allmsg&"','"&ip&"')"
conn.execute(sql)

conn.execute("update [user] set userWealth=userWealth+"&doWealth&",userCP=userCP+"&douserCP&",userEP=userEP+"&douserEP&",userIsBest=userisBest-1 where userid="&topicuserid)
sucmsg="解除精华《"&topic&"》，"&content& "，"&allmsg&""
call dvbbs_suc()
end sub

sub dele()
set rs=conn.execute("select topic,username,postuserid from "&TotalUseTable&" where Announceid="&replyid)
if rs.eof and rs.bof then
	Errmsg=Errmsg+"<br><li>没有找到相关帖子！"
	founderr=true
	exit sub
else
	topic=rs(0)
	topicusername=rs(1)
	topicuserid=rs(2)
	if topic="" then
		topic="本帖子为回复帖子"
	end if
end if
set rs=nothing

'判断用户是否有删除帖子权限
if (Cint(GroupSetting(11))=1 and TopicUsername=membername) then
	candeltopic=true
end if
if FoundUserPer and Cint(GroupSetting(11))=1 and TopicUsername=membername then
	candeltopic=true
elseif FoundUserPer and Cint(GroupSetting(11))=0 and TopicUsername=membername then
	candeltopic=false
end if
if FoundUserPer and Cint(GroupSetting(18))=1 and TopicUsername<>membername then
	candeltopic=true
elseif FoundUserPer and Cint(GroupSetting(18))=0 and TopicUsername<>membername then
	candeltopic=false
end if
if not candeltopic then
	Errmsg=Errmsg+"<br><li>您没有执行此操作的权限。"
	founderr=true
	exit sub
end if

title=CheckStr(htmlencode(request.form("title")))
content=CheckStr(htmlencode(request.form("content")))
content="原因：" & title & content
if htmlencode(request.form("title"))="" and htmlencode(request.form("content"))="" then
	Errmsg=Errmsg+"<br><li>请写明操作原因。"
	founderr=true
	exit sub
end if

dim LastPostime
sql="update "&TotalUseTable&" set locktopic=2 where boardid="&boardid&" and announceid="&replyid
conn.Execute(sql)
sql="select Max(dateandtime) from "&TotalUseTable&" where boardid="&boardid&" and rootid="&id&" and not locktopic=2"
set rs=conn.Execute(sql)
LastPostime=rs(0)

isLastPost

call LastCount(boardid)
call BoardNumSub(boardid,0,1,1)

call AllboardNumSub(1,1,0)

sql="update topic set child=child-1,LastPostTime='"&LastPostime&"' where boardid="&boardid&" and topicid="&id
conn.execute(sql)
	
sql="insert into log (l_announceid,l_boardid,l_touser,l_username,l_content,l_ip) values ("&id&","&boardid&",'"&topicusername&"','"&membername&"','删除帖子《"&topic&"》，"&content& "，"&allmsg&"','"&ip&"')"
conn.execute(sql)
if allmsg<>"" then
	conn.execute("update [user] set userWealth=userWealth+"&doWealth&",userCP=userCP+"&douserCP&",userEP=userEP+"&douserEP&",article=article-1,userDel=userDel-1 where userid="&topicuserid)
end if
sucmsg="删除帖子《"&topic&"》，"&content& "，"&allmsg&""
call dvbbs_suc()
end sub

sub delete()
dim voteid,isvote
set rs=conn.execute("select title,postusername,postuserid,PollID,isvote from topic where topicid="&id)
if rs.eof and rs.bof then
	Errmsg=Errmsg+"<br><li>没有找到相关帖子！"
	founderr=true
	exit sub
else
	topic=rs(0)
	topicusername=rs(1)
	topicuserid=rs(2)
	voteid=rs(3)
	isvote=rs(4)
	if topic="" then
		topic="本帖子为回复帖子"
	end if
end if
set rs=nothing

title=CheckStr(htmlencode(request.form("title")))
content=CheckStr(htmlencode(request.form("content")))
content="原因：" & title & content
if htmlencode(request.form("title"))="" and htmlencode(request.form("content"))="" then
	Errmsg=Errmsg+"<br><li>请写明操作原因。"
	founderr=true
	exit sub
end if
	
dim todaynum,postnum
sql="select count(*) from "&TotalUseTable&" where rootid="&id
set rs=conn.execute(sql)
postNum=rs(0)
sql="select count(*) from "&TotalUseTable&" where rootid="&id&" and dateandtime>#"&date()&"#"
set rs=conn.execute(sql)
todayNum=rs(0)
		
if allmsg<>"" then
	sql="select postuserid from "&TotalUseTable&" where rootid="&id
	set rs=conn.execute(sql)
	do while not rs.eof
	conn.execute("update [user] set userWealth=userWealth+"&doWealth&",userCP=userCP+"&douserCP&",userEP=userEP+"&douserEP&",article=article-1,userDel=userDel-1 where userid="&rs(0))
	rs.movenext
	loop
end if
set rs=nothing
	
sql="update "&TotalUseTable&" set locktopic=2 where rootid="&id
conn.Execute(sql)
if isvote=1 then
	conn.execute("update topic set locktopic=2,isvote=0,VoteTotal=0 where topicid="&id)
	conn.execute("delete from vote where voteid="&voteid)
	conn.execute("delete from voteuser where voteid="&voteid)
else
	conn.execute("update topic set locktopic=2 where topicid="&id)
end if
	
call LastCount(boardid)
call BoardNumSub(boardid,1,postNum,todayNum)
call AllboardNumSub(todayNum,postNum,1)

sql="insert into log (l_announceid,l_boardid,l_touser,l_username,l_content,l_ip) values ("&id&","&boardid&",'"&topicusername&"','"&membername&"','删除主题《"&topic&"》，"&content& "，"&allmsg&"','"&ip&"')"
conn.execute(sql)
sucmsg="删除主题《"&topic&"》，"&content& "，"&allmsg&""
call dvbbs_suc()
end sub

sub Tmove()
title=CheckStr(htmlencode(request.form("title")))
content=CheckStr(htmlencode(request.form("content")))
content="原因：" & title & content
if htmlencode(request.form("title"))="" and htmlencode(request.form("content"))="" then
	Errmsg=Errmsg+"<br><li>请写明操作原因。"
	founderr=true
	exit sub
end if
dim newboardid,newParentID,nrs
dim newtopic
set rs=server.createobject("adodb.recordset")
if request("checked")="yes" then
	if request("boardid")=request("newboardid") then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>不能在相同版面内进行转移操作。"
		exit sub
	elseif not isInteger(request("newboardid")) then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>非法的版面参数。"
		exit sub
	else
		newboardid=request("newboardid")
	end if

	sql="select * from topic where boardid="&boardid&" and topicid="&id
	set rs=conn.execute(sql)
	if rs.eof and rs.bof then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>您选择的贴子并不存在。"
		exit sub
	else
		if request.form("isdispmove")="yes" then
			newtopic=CheckStr(request.form("topic")) & "-->" & membername & "转移"
		else
			newtopic=CheckStr(request.form("topic"))
		end if
		if request("leavemessage")="yes" then
			sql="insert into topic (Title,Boardid,PostUsername,PostUserid,DateAndTime,Expression,LastPost,LastPostTime,child,hits,isvote,isbest,votetotal,PostTable) values ('"&newtopic&"',"&newboardid&",'"&rs("postusername")&"',"&rs("postuserid")&",'"&rs("dateandtime")&"','"&rs("Expression")&"','"&rs("LastPost")&"','"&rs("LastPosttime")&"',"&rs("child")&","&rs("hits")&","&rs("isvote")&","&rs("isbest")&","&rs("votetotal")&",'"&NowUseBBS&"')"
			conn.execute(sql)
		end if
	end if
	
	if request("leavemessage")="yes" then
		conn.execute("update topic set locktopic=1 where topicid="&id)
		set rs=conn.execute("select top 1 topicid from topic order by topicid desc")
		newparentid=rs(0)
		sql="select * from "&TotalUseTable&" where rootid="&id&" and not locktopic=2 order by Announceid"
		set rs=conn.execute(sql)
		do while not rs.eof
			Sql="insert into "&NowUseBBS&"(Boardid,ParentID,username,topic,body,DateAndTime,length,rootid,layer,orders,ip,Expression,locktopic,signflag,emailflag,isbest,postuserid) values "&_
				"("&_
				newboardid&","&rs("parentid")&",'"&_
				rs("username")&"','"&_
				rs("topic")&"','"&_
				checkstr(rs("body"))&"','"&_
				rs("DateAndTime")&"','"&_
				rs("length")&"',"&newParentID&","&rs("layer")&","&rs("orders")&",'"&rs("ip")&"','"&_
				rs("Expression")&"',"&rs("locktopic")&","&rs("signflag")&","&rs("emailflag")&","&rs("isbest")&","&rs("postuserid")&")"
				'response.write sql
				conn.execute(sql)
		rs.movenext
		loop
		
	elseif request("leavemessage")="no" then
		if request.form("isdispmove")="yes" then
			newtopic=CheckStr(request.form("topic")) & "-->" & membername & "转移"
		else
			newtopic=CheckStr(request.form("topic"))
		end if
		conn.execute("update topic set title='"&newtopic&"',boardid="&newboardid&" where topicid="&id)
		sql="update "&TotalUseTable&" set topic='"&newtopic&"' where announceid="&replyid
		conn.execute(sql)
		sql="update "&TotalUseTable&" set boardid="&newboardid&" where rootid="&id
		conn.Execute(sql)							
	else
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>请选择相应操作。"
		exit sub
	end if
	dim postNum,todayNum
	'计算该帖子的回复数量，用来统计对应版面帖子数
	sql="select count(*) from "&TotalUseTable&" where rootid="&id
	set rs=conn.execute(sql)
	postNum=rs(0)
	'计算该帖子中今日回复的数量
	sql="select count(*) from "&TotalUseTable&" where rootid="&id&" and dateandtime>=#"&date()&"#"
	set rs=conn.execute(sql)
	todayNum=rs(0)
	set rs=nothing
	'更新论坛贴子数据
	call LastCount(boardid)
	call BoardNumSub(boardid,1,postNum,todayNum)
	call LastCount(newboardid)
	call BoardNumAdd(newboardid,1,postNum,todayNum)
	'更新论坛数据结束
	sql="insert into log (l_announceid,l_boardid,l_touser,l_username,l_content,l_ip) values ("&id&","&boardid&",'"&topicusername&"','"&membername&"','转移主题《"&topic&"》，"&content& "，"&allmsg&"','"&ip&"')"
	conn.execute(sql)
	sucmsg="转移主题《"&topic&"》，"&content& "，"&allmsg&""
	call dvbbs_suc()
else
%>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
                        <tr>
                        <th valign=middle colspan=2>
                        <form action="admin_postings.asp" method="post">
                        <input type=hidden name="action" value="move">
                        <input type=hidden name="checked" value="yes">
                        <input type=hidden name="boardid" value="<%=request("boardid")%>">
                        <input type=hidden name="replyid" value="<%=request("replyid")%>">
                        <input type=hidden name="id" value="<%=request("id")%>">
						<input type=hidden value="<%=CheckStr(htmlencode(request.form("title")))%>" name="title">
						<input type=hidden value="<%=CheckStr(htmlencode(request.form("content")))%>" name="content">
						<input type=hidden value="<%=doWealth%>" name="doWealth">
						<input type=hidden value="<%=dousercp%>" name="dousercp">
						<input type=hidden value="<%=douserep%>" name="douserep">
                        移动主题</th></tr>

                                    <tr>
                                    <td class=tablebody1 valign=middle>
                                    <b>移动选项</td>
                                    <td class=tablebody1 valign=middle>
                                    <input name="leavemessage" type="radio" value="yes"> 移动并保留一个已经锁定的主题在原论坛<br>
									<input name="leavemessage" type="radio" value="no" checked> 移动并将此主题从原论坛中删除<br>
									<input name="isdispmove" type="checkbox" value="yes"> 移动后的主题是否显示操作者名称
                                    </td>
                                    </tr>
                                    <tr>
                                    <td class=tablebody1 valign=middle>
                                    <b>帖子主题</td>
                                    <td class=tablebody1 valign=middle>
   <input name="topic" value="<%=topic%>" size=50>
                                    </td>
                                    </tr>
                            <tr>
                        <td class=tablebody1 valign=top><b>移动到：</b></td>
                        <td class=tablebody1 valign=top>
<%
	sql="select boardid,boardtype from board"
	set rs=conn.execute(sql)
%>
<select name="newboardid" size="1">
<option value="">选择一个论坛</option>
<%
	do while not rs.eof
        response.write "<option value='"+CStr(rs("BoardID"))+"'>"+rs("Boardtype")+"</option>"+chr(13)+chr(10)
	rs.movenext
	loop
	rs.close
	set rs=nothing
%>        
          </select>&nbsp;注意：不能在相同论坛内进行移动操作
			</td>
                        </tr>
                    <tr>
                <td class=tablebody2 valign=middle colspan=2 align=center><input type=submit name="submit" value="提 交"></td></tr></form></table>
            </table>
            </td></tr>
            </table>
<%
end if
end sub

sub copy()
set rs=conn.execute("select topic,username,postuserid from "&TotalUseTable&" where Announceid="&replyid)
if rs.eof and rs.bof then
	Errmsg=Errmsg+"<br><li>没有找到相关帖子！"
	founderr=true
	exit sub
else
	topic=rs(0)
	topicusername=rs(1)
	topicuserid=rs(2)
	if topic="" then
		topic="本帖子为回复帖子"
	end if
end if
set rs=nothing

'判断用户是否有移动帖子权限
if (Cint(GroupSetting(12))=1 and TopicUsername=membername) then
	canmovetopic=true
end if
if FoundUserPer and Cint(GroupSetting(12))=1 and TopicUsername=membername then
	canmovetopic=true
elseif FoundUserPer and Cint(GroupSetting(12))=0 and TopicUsername=membername then
	canmovetopic=false
end if
if FoundUserPer and Cint(GroupSetting(19))=1 and TopicUsername<>membername then
	canmovetopic=true
elseif FoundUserPer and Cint(GroupSetting(19))=0 and TopicUsername<>membername then
	canmovetopic=false
end if
if not canmovetopic then
	Errmsg=Errmsg+"<br><li>您没有执行此操作的权限。"
	founderr=true
	exit sub
end if
title=CheckStr(htmlencode(request.form("title")))
content=CheckStr(htmlencode(request.form("content")))
content="原因：" & title & content
if htmlencode(request.form("title"))="" and htmlencode(request.form("content"))="" then
	Errmsg=Errmsg+"<br><li>请写明操作原因。"
	founderr=true
	exit sub
end if
	
if request("checked")="yes" then
	dim newboardid
	dim todaynum,postnum
	sql="select count(*) from "&TotalUseTable&" where rootid="&id
	set rs=conn.execute(sql)
	postNum=rs(0)
	sql="select count(*) from "&TotalUseTable&" where rootid="&id&" and dateandtime>#"&date()&"#"
	set rs=conn.execute(sql)
	todayNum=rs(0)
	if request("boardid")=request("newboardid") then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>不能在相同版面内进行转移操作。"
		exit sub
	elseif not isInteger(request("newboardid")) then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>非法的版面参数。"
		exit sub
	else
		newboardid=request("newboardid")
	end if

	sql="select boardid from "&TotalUseTable&" where announceid="&replyid&" and boardid="&cstr(boardid)
	set rs=conn.execute(sql)
	if rs.eof and rs.bof then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>您选择的贴子并不存在。"
		exit sub
	end if

	dim newtopic,trs
	set rs=server.createobject("adodb.recordset")
	sql="select * from "&TotalUseTable&" where announceid="&replyid
	rs.open sql,conn,1,1
	if request.form("isdispmove")="yes" then
		newtopic=CheckStr(request.form("topic")) & "-->" & membername & "添加"
	else
		newtopic=CheckStr(request.form("topic"))
	end if
	sql="insert into topic (Title,Boardid,PostUsername,PostUserid,DateAndTime,Expression,LastPost,LastPostTime,child,hits,isvote,isbest,votetotal,PostTable) values ('"&newtopic&"',"&newboardid&",'"&rs("username")&"',"&rs("postuserid")&",Now(),'"&rs("Expression")&"','$$$$$$',Now(),0,0,0,0,0,'"&NowUseBBS&"')"
	conn.execute(sql)
	set trs=conn.execute("select top 1 topicid from topic order by topicid desc")
	Sql="insert into "&NowUseBBS&"(Boardid,ParentID,username,topic,body,DateAndTime,length,rootid,layer,orders,ip,Expression,locktopic,signflag,emailflag,isbest,postuserid) values "&_
				"("&_
				newboardid&",0,'"&_
				rs("username")&"','"&_
				newtopic&"','"&_
				rs("body")&"','"&_
				rs("DateAndTime")&"','"&_
				rs("length")&"',"&trs(0)&",1,0,'"&rs("ip")&"','"&_
				rs("Expression")&"',"&rs("locktopic")&","&rs("signflag")&","&rs("emailflag")&","&rs("isbest")&","&rs("postuserid")&")"
	conn.execute(sql)
	rs.close
	set rs=nothing
	
	'更新论坛贴子数据
	call LastCount(Newboardid)
	call BoardNumAdd(newboardid,1,postNum,todayNum)
	call AllboardNumAdd(todayNum,postNum,1)
	
	sql="insert into log (l_announceid,l_boardid,l_touser,l_username,l_content,l_ip) values ("&id&","&boardid&",'"&topicusername&"','"&membername&"','拷贝帖子《"&topic&"》，"&content& "，"&allmsg&"','"&ip&"')"
	conn.execute(sql)
	'更新论坛数据结束
	sucmsg="拷贝帖子《"&topic&"》，"&content& "，"&allmsg&""
	call dvbbs_suc()
else
%>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1 >
                        <tr>
                        <th valign=middle colspan=2>
                        <form action="admin_postings.asp" method="post">
                        <input type=hidden name="action" value="copy">
                        <input type=hidden name="checked" value="yes">
                        <input type=hidden name="boardid" value="<%=request("boardid")%>">
                        <input type=hidden name="replyid" value="<%=request("replyid")%>">
                        <input type=hidden name="id" value="<%=request("id")%>">
<input type=hidden value="<%=CheckStr(htmlencode(request.form("title")))%>" name="title">
<input type=hidden value="<%=CheckStr(htmlencode(request.form("content")))%>" name="content">
<input type=hidden value="<%=doWealth%>" name="doWealth">
<input type=hidden value="<%=dousercp%>" name="dousercp">
<input type=hidden value="<%=douserep%>" name="douserep">
                        复制帖子</th></tr>

                                    <tr>
                                    <td class=tablebody1 valign=middle>
                                    <b>操作说明</td>
                                    <td class=tablebody1 valign=middle>
该帖子将复制到别的论坛，成为新的帖子，所以如该帖子无主题请指定该帖子的主题，原帖子将保留在原来论坛
                                    </td>
                                    </tr>
                                    <tr>
                                    <td class=tablebody1 valign=middle>
                                    <b>帖子主题</td>
                                    <td class=tablebody1 valign=middle>
   <input name="topic" value="<%=topic%>" size=50>&nbsp;<input name="isdispmove" type="checkbox" value="yes"> 移动后的主题是否显示操作者名称
                                    </td>
                                    </tr>
                            <tr>
                        <td class=tablebody1 valign=top><b>移动到：</b></td>
                        <td class=tablebody1 valign=top>
<%
	set rs=server.createobject("adodb.recordset")
	sql="select boardid,boardtype from board"
	rs.open sql,conn,1,1
%>
<select name="newboardid" size="1">
<option value="">选择一个论坛</option>
<%
	do while not rs.eof
        response.write "<option value='"+CStr(rs("BoardID"))+"'>"+rs("Boardtype")+"</option>"+chr(13)+chr(10)
	rs.movenext
	loop
	rs.close
	set rs=nothing
%>        
          </select>&nbsp;注意：相同论坛间不能进行复制操作
			</td>
                        </tr>
                    <tr>
                <td class=tablebody2 valign=middle colspan=2 align=center><input type=submit name="submit" value="提 交"></td></tr></form></table></td></tr></table>
            </table>
            </td></tr>
            </table>
<%
	end if
	end sub

	sub upTopic()
	if request.Form("checked")="yes" then

	else
%>
            <table cellpadding=0 cellspacing=0 border=0 width=95% bgcolor=<%=Forum_body(0)%> align=center>
                <tr>
                    <td>
                    <table cellpadding=6 cellspacing=1 border=0 width=100%>
                        <tr>
                        <td bgcolor=<%=Forum_body(2)%> valign=middle align=center colspan=2>
                        <form action="admin_postings.asp" method="post">
                        <input type=hidden name="action" value="upTopic">
                        <input type=hidden name="checked" value="yes">
                        <input type=hidden name="boardid" value="<%=request("boardid")%>">
                        <input type=hidden name="rootid" value="<%=request("rootid")%>">
                        <input type=hidden name="id" value="<%=request("id")%>">
<input type=hidden value="<%=CheckStr(htmlencode(request.form("title")))%>" name="title">
<input type=hidden value="<%=CheckStr(htmlencode(request.form("content")))%>" name="content">
<input type=hidden value="<%=doWealth%>" name="doWealth">
<input type=hidden value="<%=dousercp%>" name="dousercp">
<input type=hidden value="<%=douserep%>" name="douserep">
                        <b>提升帖子</b></td></tr>

                                    <tr>
                                    <td bgcolor=<%=Forum_body(4)%> valign=middle>
                                    <b>操作说明</td>
                                    <td bgcolor=<%=Forum_body(4)%> valign=middle>
将版面某一主题和某一主题的位置互换，也就是把某一比较靠后的主题和靠前的主题替换，要求输入两个主题的RootID号，主题的RootID号可通过点开主题看浏览器地址栏获得
                                    </td>
                                    </tr>
                            <tr>
                        <td bgcolor=<%=Forum_body(4)%> valign=top><b>请填写Rootid(数字)：</b></td>
                        <td bgcolor=<%=Forum_body(4)%> valign=top>
						靠后主题 <input type="text" name="OldRootid" value="<%=request("rootid")%>">&nbsp;&nbsp;
						靠前主题 <input type="text" name="NewRootid">
						</td>
                        </tr>
                    <tr>
                <td bgcolor=<%=Forum_body(3)%> valign=middle colspan=2 align=center><input type=submit name="submit" value="提 交"></td></tr></form></table></td></tr></table>
            </table>
            </td></tr>
            </table>
<%
	end if
	end sub

'更新指定论坛信息
function LastCount(boardid)
dim LastTopic,body,LastRootid,LastPostTime,LastPostUser
dim LastPost,uploadpic_n,Lastpostuserid,Lastid
set rs=conn.execute("select top 1 topic,body,Announceid,dateandtime,username,postuserid,rootid from "&NowUseBBS&" where boardid="&boardid&" and not locktopic=2 order by announceid desc")
if not(rs.eof and rs.bof) then
	Lasttopic=replace(rs(0),"'","")
	body=rs(1)
	LastRootid=rs(2)
	LastPostTime=rs(3)
	LastPostUser=rs(4)
	LastPostUserid=rs(5)
	Lastid=rs(6)
else
	LastTopic="无"
	LastRootid=0
	LastPostTime=now()
	LastPostUser="无"
	LastPostUserid=0
	Lastid=0
end if
if Lasttopic="" or isnull(Lasttopic) then LastTopic=replace(left(body,20),"'","")
set rs=nothing

LastPost=LastPostUser & "$" & LastRootid & "$" & LastPostTime & "$" & replace(left(LastTopic,20),"$","") & "$" & uploadpic_n & "$" & LastPostUserID & "$" & LastID
sql="update board set LastPost='"&LastPost&"' where boardid="&boardid
conn.execute(sql)
end function
	
'版面发帖数增加
sub BoardNumAdd(boardid,topicNum,postNum,todayNum)
sql="update board set lastbbsnum=lastbbsnum+"&postNum&",lasttopicNum=lasttopicNum+"&topicNum&",todayNum=todayNum+"&todayNum&" where boardid="&boardid
conn.execute(sql)
end sub
	
'版面发帖数减少
sub BoardNumSub(boardid,topicNum,postNum,todayNum)
sql="update board set lastbbsnum=lastbbsnum-"&postNum&",lasttopicNum=lasttopicNum-"&topicNum&",todayNum=todayNum-"&todayNum&" where boardid="&boardid
conn.execute(sql)
end sub
	
'所有论坛发帖数增加
function AllboardNumAdd(todayNum,postNum,topicNum)
sql="update config set TodayNum=todayNum+"&todaynum&",BbsNum=bbsNum+"&postNum&",TopicNum=topicNum+"&TopicNum
conn.execute(sql)
end function

'所有论坛发帖数减少
function AllboardNumSub(todayNum,postNum,topicNum)
sql="update config set TodayNum=todayNum-"&todaynum&",BbsNum=bbsNum-"&postNum&",TopicNum=topicNum-"&TopicNum
conn.execute(sql)
end function

'判断是否为帖子最后回复
function isLastPost()
dim LastTopic,body,LastRootid,LastPostTime,LastPostUser
dim LastPost,uploadpic_n,LastPostUserID,LastID
isLastPost=false
'取得当前主题最后回复id
sql="select LastPost from topic where topicid="&id
set rs=conn.execute(sql)
if not (rs.eof and rs.bof) then
	if not isnull(rs(0)) and rs(0)<>"" then
		if Clng(split(rs(0),"$")(1))=Clng(replyid) then isLastPost=true
	end if
end if
if isLastPost then
	sql="select top 1 topic,body,Announceid,dateandtime,username,PostUserid,rootid from "&TotalUseTable&" where rootid="&id&" and not locktopic=2 order by Announceid desc"
	set rs=conn.execute(sql)
	if not(rs.eof and rs.bof) then
		body=rs(1)
		LastRootid=rs(2)
		LastPostTime=rs(3)
		LastPostUser=replace(rs(4),"$","")
		LastTopic=left(replace(body,"$",""),20)
		LastPostUserID=rs(5)
		LastID=rs(6)
	else
		LastTopic="无"
		LastRootid=0
		LastPostTime=now()
		LastPostUser="无"
		LastPostUserID=0
		LastID=0
	end if
	set rs=nothing
	LastPost=LastPostUser & "$" & LastRootid & "$" & LastPostTime & "$" & replace(left(LastTopic,20),"$","") & "$" & uploadpic_n & "$" & LastPostUserID & "$" & LastID
	conn.execute("update topic set LastPost='"&LastPost&"' where topicid="&id)
end if
end function
%>