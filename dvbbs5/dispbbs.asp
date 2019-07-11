<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/char_board.asp" -->
<!--#include file="inc/birthday.asp"-->
<!-- #include file="inc/ubbcode.asp" -->
<%
'=========================================================
' File: dispbbs.asp
' Version:5.0
' Date: 2002-9-7
' Script Written by satan
'=========================================================
' Copyright (C) 2001,2002 AspSky.Net. All rights reserved.
' Web: http://www.aspsky.net,http://www.dvbbs.net
' Email: info@aspsky.net,eway@aspsky.net
'=========================================================
Dim AnnounceID
Dim ReplyID
Dim Star,nSkin,SkinPic,Skiname
Dim Topic_1,IsTop,IsBest,IsVote
Dim UserName,view,times
Dim onlineUserList
Dim userhiddensql
Dim Page_Count,TotalRec,abgcolor,bgcolor
Dim TopicCount
Dim Pcount,endpage
Dim isagree,noagree
Dim PostUserName,PostUserid
Dim pollid
Dim TotalUseTable
Dim canreply,mycanreply
Dim LockTopic
Page_count=0
canreply=false
i=1
if boardmaster or master then
	guestlist=""
	userhiddensql=""
else
	guestlist=" lockboard<>1 and "
	userhiddensql=" and userhidden=2"
end if

stats="浏览帖子"
if BoardID="" or (not isInteger(BoardID)) or BoardID="0" then
	Errmsg=Errmsg+"<br>"+"<li>错误的版面参数！请确认您是从有效的连接进入。"
	founderr=true
else
	BoardID=clng(BoardID)
end if
if request("id")="" then
	Errmsg=Errmsg+"<br>"+"<li>请指定相关贴子。"
	founderr=true
elseif not isInteger(request("id")) then
	Errmsg=Errmsg+"<br>"+"<li>非法的贴子参数。"
	founderr=true
else
	AnnounceID=request("id")
end if
if request("replyid")="" then
	replyid=Announceid
elseif not isInteger(request("replyid")) then
	replyid=Announceid
else
	replyid=request("replyid")
end if
if request("star")="" or not isnumeric(request("star")) then
	star=1
else
	star=clng(request("star"))
end if

if request("skin")="" or not isnumeric(request("skin")) then
	skin=Forum_Setting(62)
else
	skin=request("skin")
end if
if skin=1 then
	nskin=0
	skinpic=Forum_boardpic(12)
	skiname="平板"
elseif skin="0" then
	nskin=1
	skinpic=Forum_boardpic(11)
	skiname="树形"
else
	skin=0
	nskin=1
	skinpic=Forum_boardpic(11)
	skiname="树形"
end if

if cint(lockboard)=2 then
	if not founduser then
		Errmsg=Errmsg+"<br>"+"<li>本论坛为认证论坛，请<a href=login.asp>登陆</a>并确认您的用户名已经得到管理员的认证后进入。"
		founderr=true
	else
		if chkboardlogin(boardid,membername)=false then
		Errmsg=Errmsg+"<br>"+"<li>本论坛为认证论坛，请确认您的用户名已经得到管理员的认证后进入。"
		founderr=true
		end if
	end if
end if

if founderr then
	call nav()
	call head_var(1)
	call dvbbs_error()
else
	conn.execute("update topic set hits=hits+1 where topicid="&Announceid)
	sql="select title,istop,isbest,PostUserName,PostUserid,hits,isvote,child,pollid,LockTopic,PostTable from topic where topicID="&Announceid
	set rs=conn.execute(sql)
	if not(rs.bof and rs.eof) then
		if rs("locktopic")=2 then
			ErrMsg=ErrMsg+"<br>"+"<li>该帖子已经被管理员删除！</li>"
			founderr=true
		end if
		topic_1=rs(0)
		istop=rs(1)
		isbest=rs(2)
		PostUserName=rs(3)
		PostUserID=rs(4)
		view=rs(5)
		isVote=rs(6)
		TopicCount=rs(7)+1
		pollid=rs(8)
		Locktopic=rs(9)
		TotalUseTable=rs(10)
		stats=""&boardtype&"浏览："&topic_1&""
		if PostUserName=membername then
			call readRe()
			mycanreply=GroupSetting(4)
		else
			mycanreply=GroupSetting(5)
			if Cint(GroupSetting(2))=0 then
				Errmsg=Errmsg+"<br>"+"<li>您没有浏览在本论坛查看其他人发布的帖子的权限，请<a href=login.asp>登陆</a>或者同管理员联系。"
				founderr=true
			end if
		end if
	else
		ErrMsg=ErrMsg+"<br>"+"<li>您指定的贴子不存在</li>"
		founderr=true
	end if
	set rs=nothing
	call nav()
	call head_var(2)
	if founderr then
		call dvbbs_error()
	else
		call main()
		if founderr then call dvbbs_error()
	end if
end if
call footer()

sub main()
if lockboard<>1 and Cint(mycanreply)=1 and Cint(locktopic)=0 then
	canreply=true
end if

if IsVote=1 then
	Dim vrs,vote,vote_1,votenum,votenum_1,m,g,votetype
	Dim vurs,voteyes
	voteyes=0
	g=0
	sql="select * from vote where VoteID="&PollID
	set vrs=server.createobject("adodb.recordset")
	vrs.open sql,conn,1,1
	if not (vrs.eof and vrs.bof) then
	voteyes=1
	response.write "<table cellspacing=1 cellpadding=3 align=center class=tableborder1 style=""table-layout:fixed;word-break:break-all""><tr><th align=left colspan=2 height=25>[投票]："&Topic_1&"</th></tr><form action=postvote.asp?BoardID="&BoardID&"&voteID="&PollID&"&id="&Announceid&"&action="&vrs("votetype")&" method=POST>"
	vote=split(vrs("vote"),"|")
	votenum=split(vrs("votenum"),"|")
	for i = 0 to ubound(votenum)
		votenum_1=cint(votenum_1)+votenum(i)
	next
	if votenum_1=0 then votenum_1=1
	for m = 0 to ubound(vote)
		g=g+1
		if g=11 then g=1
		if cint(vrs("votetype"))=0 then
		votetype="<input type=radio name=postvote value="""&m&""">"
		else
		votetype="<input type=checkbox name=postvote_"&m&" value="""&m&""">"
		end if
		response.write "<tr><td width=""60%"" height=25 class=tablebody1>"&m+1&".  "&votetype & htmlencode(vote(m))&"</td><td width=""40%"" class=tablebody1><img src="""&Forum_info(7)&"bar"&g&".gif"" width="""&Cint(replace(FormatPercent(votenum(m)/votenum_1),"%",""))*3.3&""" height=8> <b>"&votenum(m)&"票</b></td></tr>"
	next

	if not founduser or datediff("d",vrs("timeout"),Now())>0 then
		response.write "<tr><td class=tablebody2 colspan=2 height=25>您还没有登陆，不能进行投票；或者已经过了投票期限。[<a href=javascript:openScript('viewvoters.asp?id="&pollid&"',300,500)>查看投票用户</a>]</td></tr>"
	else
		set vurs=conn.execute("select userid from voteuser where voteid="&PollID&" and userid="&userid)
		if vurs.eof and vurs.bof then
			response.write "<tr><td colspan=2 height=25 class=tablebody1><input type=submit name=Submit value='投 票'>  [截止时间："&vrs("timeout")&" | <a href=javascript:openScript('viewvoters.asp?id="&pollid&"',300,500)>查看投票用户</a>]</td></tr>"
		else
			response.write "<tr><td class=tablebody2 colspan=2 height=25>您已经投过票了，请看结果吧。[过期时间："&vrs("timeout")&" | <a href=javascript:openScript('viewvoters.asp?id="&pollid&"',300,500)>查看投票用户</a>]</td></tr>"
		end if
		set vurs=nothing
	end if
	response.write "</form></table><BR>"
	end if
	vrs.close
	set vrs=nothing
end if
%>
<table cellpadding=0 cellspacing=0 border=0 width=<%=Forum_body(12)%> align=center>
	<tr>
	<td align=left width="30%" valign=middle>&nbsp; 
	<a href="announce.asp?BoardID=<%=BoardID%>"><img src="<%=Forum_info(7)&Forum_boardpic(1)%>" alt="发表一个新主题" border=0></a>&nbsp;
	<a href="vote.asp?BoardID=<%=BoardID%>"><img src="<%=Forum_info(7)&Forum_boardpic(2)%>" alt=发表一个新投票 border=0></a>&nbsp;
	<a href="reannounce.asp?BoardID=<%=BoardID%>&id=<%=AnnounceID%>&star=<%=star%>"><img src="<%=Forum_info(7)&Forum_boardpic(4)%>" alt=回复主题 border=0></a>
	</td>
	<td align=right width="70%" valign=middle>您是本帖的第 <B><%=view%></B> 个阅读者<a href="go.asp?BoardID=<%=BoardID%>&sid=<%=Announceid%>"><img src="<%=Forum_info(7)&Forum_boardpic(14)%>" border=0 alt=浏览上一篇主题 width=52 height=12></a>&nbsp;
	<a href="javascript:this.location.reload()"><img src="<%=Forum_info(7)&Forum_boardpic(15)%>" border=0 alt=刷新本主题 width=40 height=12></a> &nbsp;
	<a href="?BoardID=<%=BoardID%>&replyID=<%=replyID%>&id=<%=request("id")%>&star=<%=star%>&skin=<%=nskin%>"><img src="<%=Forum_info(7)&skinpic%>" width=40 height=12 border=0 alt=<%=skiname%>显示贴子></a>　<a href="go.asp?BoardID=<%=BoardID%>&sid=<%=Announceid%>&action=next"><img src="<%=Forum_info(7)&Forum_boardpic(13)%>" border=0 alt=浏览下一篇主题 width=52 height=12></a>
	</td>
	</tr>
</table>

<TABLE cellPadding=0 cellSpacing=1 align=center class=tableborder1>
	<tr align=middle> 
	<td align=left valign=middle width="100%" height=25>
		<table width=100% cellPadding=0 cellSpacing=0 border=0>
		<tr>
		<th align=left valign=middle width="73%" height=25>
		&nbsp;* 贴子主题</B>： <%=htmlencode(topic_1)%></th>
		<th width=37% align=right>
		<a href=# onclick="javascript:WebBrowser.ExecWB(4,1)"><img src="<%=Forum_info(7)&Forum_TopicPic(0)%>" border=0 width=16 height=16 alt=保存该页为文件 align=absmiddle></a>&nbsp;<object id="WebBrowser" width=0 height=0 classid="CLSID:8856F961-340A-11D0-A96B-00C04FD705A2"></object>
		<a href="report.asp?BoardID=<%=BoardID%>&id=<%=AnnounceID%>"><img src=<%=Forum_info(7)&Forum_TopicPic(1)%> alt=报告本帖给版主 border=0></a>&nbsp;
		<a href="printpage.asp?BoardID=<%=BoardID%>&id=<%=AnnounceID%>"><img src="<%=Forum_info(7)&Forum_TopicPic(2)%>" alt=显示可打印的版本 border=0></a>&nbsp;
		<a href="pag.asp?BoardID=<%=BoardID%>&id=<%=AnnounceID%>"><img src="<%=Forum_info(7)&Forum_TopicPic(3)%>" border=0 alt=把本贴打包邮递></a>&nbsp;
		<a href="favadd.asp?BoardID=<%=BoardID%>&id=<%=Announceid%>"><IMG SRC="<%=Forum_info(7)&Forum_TopicPic(4)%>" BORDER=0 alt=把本贴加入论坛收藏夹></a>&nbsp;
		<a href="sendpage.asp?BoardID=<%=BoardID%>&id=<%=AnnounceID%>"><img src="<%=Forum_info(7)&Forum_TopicPic(5)%>" border=0 alt=发送本页面给朋友></a>&nbsp;
		<a href=#><span style="CURSOR: hand" onClick="window.external.AddFavorite('<%=Forum_info(1)%>/dispbbs.asp?BoardID=<%=BoardID%>&id=<%=AnnounceID%>', ' <%=Forum_info(0)%> - <%=htmlencode(replace(topic_1,"&#39;",""))%>')"><IMG SRC="<%=Forum_info(7)&Forum_TopicPic(6)%>" BORDER=0 width=15 height=15 alt=把本贴加入IE收藏夹></a>&nbsp;
		</th>
		</tr>
		</table>
	</td>
	</tr>
</table>

<%
if cint(skin)=1 and replyid=Announceid then
	sql="Select top 1  B.AnnounceID,B.BoardID,B.UserName,B.Topic,B.dateandtime,B.body,B.Expression,B.ip,B.RootID,B.signflag,B.isbest,B.PostUserid,B.layer,b.isagree,U.useremail,U.homepage,U.oicq,U.sign,U.userclass,U.title,U.width,U.height,U.article,U.face,U.addDate,U.userWealth,U.userEP,U.userCP,U.birthday,U.sex,u.UserGroup,u.LockUser,u.userPower,U.titlepic,U.UserGroupID,U.LastLogin from "&TotalUseTable&" B inner join [user] U on U.UserID=B.PostUserID where B.BoardID="&BoardID&" and B.rootID="&AnnounceID&" and not B.locktopic=2 order by Announceid"
elseif cint(skin)=1 then
	sql="Select B.AnnounceID,B.BoardID,B.UserName,B.Topic,B.dateandtime,B.body,B.Expression,B.ip,B.RootID,B.signflag,B.isbest,B.PostUserid,B.layer,b.isagree,U.useremail,U.homepage,U.oicq,U.sign,U.userclass,U.title,U.width,U.height,U.article,U.face,U.addDate,U.userWealth,U.userEP,U.userCP,U.birthday,U.sex,u.UserGroup,u.LockUser,u.userPower,U.titlepic,U.UserGroupID,U.LastLogin from "&TotalUseTable&" B inner join [user] U on U.UserID=B.PostUserID where B.BoardID="&BoardID&" and B.AnnounceID="&replyID&" and not B.locktopic=2"
else
	sql="Select B.AnnounceID,B.BoardID,B.UserName,B.Topic,B.dateandtime,B.body,B.Expression,B.ip,B.RootID,B.signflag,B.isbest,B.PostUserid,B.layer,b.isagree,U.useremail,U.homepage,U.oicq,U.sign,U.userclass,U.title,U.width,U.height,U.article,U.face,U.addDate,U.userWealth,U.userEP,U.userCP,U.birthday,U.sex,u.UserGroup,u.LockUser,u.userPower,U.titlepic,U.UserGroupID,U.LastLogin from "&TotalUseTable&" B inner join [user] U on U.UserID=B.PostUserID where B.BoardID="&BoardID&" and B.RootID="&Announceid&" and not B.locktopic=2 order by B.AnnounceID"
end if
set rs=server.createobject("adodb.recordset")
'response.write sql
rs.open sql,conn,1
if not(rs.eof and rs.bof) then
  	if TopicCount mod Cint(Forum_Setting(12))=0 then
     		Pcount= TopicCount \ Cint(Forum_Setting(12))
  	else
     		Pcount= TopicCount \ Cint(Forum_Setting(12))+1
  	end if
	RS.MoveFirst
	if star > Pcount then star = Pcount
   	if star < 1 then star = 1
	RS.Move (star-1) * Forum_Setting(12)
	page_count=0
	Dim useremail,homepage,oicq,sign,userclass,title,width,uheight,article,face
	Dim addtime,userWealth,userEP,userCP,userbirth,sex,UserGroup,usercolor,namestyle
	Dim LockUser,userPower,istopic,PostUserGroup,UserLastLogin,userip,reUserGroupID
	Dim bestTitle,istoptitle
	Dim RootID,signflag,isbest,PostUserID,layer,isagree,bbsisagree
	istopic=false
'AnnounceID=0,BoardID=1,UserName=2,Topic=3,dateandtime=4,body=5,Expression=6,ip=7,RootID=8,signflag=9,isbest=10,PostUserid=11,layer=12,isagree=13,useremail=14,homepage=15,oicq=16,sign=17,userclass=18,title=19,width=20,height=21,article=22,face=23,addDate=24,userWealth=25,userEP=26,userCP=27,birthday=28,sex=29,UserGroup=30,LockUser=31,userPower=32,titlepic=33,UserGroupID=34,LastLogin=35
	do while not rs.eof and page_count < Cint(Forum_Setting(12))
	username=rs(2)
	RootID=rs(8)
	signflag=rs(9)
	isbest=rs(10)
	PostUserID=rs(11)
	layer=rs(12)
	isagree=rs(13)
	if isnull(isagree) or isagree="" then isagree="0|0"
	bbsisagree=split(isagree,"|")
	if star=1 and skin=0 and layer=1 then
		istopic=true
		replyid=rs(0)
		noagree=bbsisagree(1)
		isagree=bbsisagree(0)
	elseif layer=1 and skin=1 then
		istopic=true
		replyid=rs(0)
		noagree=bbsisagree(1)
		isagree=bbsisagree(0)
	end if
	useremail=rs(14)
	homepage=rs(15)
	oicq=rs(16)
	sign=rs(17)
	userclass=rs(18)
	title=rs(19)
	width=rs(20)
	uheight=rs(21)
	article=rs(22)
	face=rs(23)
	addtime=rs(24)
	userWealth=rs(25)
	userEP=rs(26)
	userCP=rs(27)
	userbirth=rs(28)
	sex=rs(29)
	UserGroup=rs(30)
	LockUser=rs(31)
	userPower=rs(32)
	reUserGroupID=rs(34)
	UserLastLogin=rs(35)
	if reUserGroupID=8 then
	namestyle="filter:glow(color="&Forum_body(22)&",strength=2)"
	usercolor=Forum_body(21)
	elseif reUserGroupID=3 then
	namestyle="filter:glow(color="&Forum_body(18)&",strength=2)"
	usercolor=Forum_body(17)
	elseif reUserGroupID=1 then
	namestyle="filter:glow(color="&Forum_body(20)&",strength=2)"
	usercolor=Forum_body(19)
	else
	namestyle="filter:glow(color="&Forum_body(16)&",strength=2)"
	usercolor=Forum_body(15)
	end if
	if bgcolor="tablebody1" then
		bgcolor="tablebody2"
		abgcolor="tablebody1"
	else
		bgcolor="tablebody1"
		abgcolor="tablebody2"
	end if
	if userbirth="" or isnull(userbirth) then
	userbirth=""
	else
	userbirth=astro(userbirth)
	end if

	response.write "<TABLE cellPadding=5 cellSpacing=1 align=center class=tableborder1 style=""table-layout:fixed; word-break:break-all"">"
	response.write "<tr><td class="&bgcolor&" valign=top width=175>"
	response.write "<table width=100% cellpadding=4 cellspacing=0><tr>"
	response.write "<td width=* valign=middle style="""&namestyle&""">&nbsp;<a name="&rs("AnnounceID")&"><font color="&usercolor&"><B>"&htmlencode(UserName)&"</B></font></a>	</td><td width=25 valign=middle>"
	if DateDiff("s",UserLastLogin,Now())>Cint(Forum_Setting(8))*60 then
		if sex=1 then
		response.write "<img src=pic/ofMale.gif alt=帅哥哟，离线，有人找我吗？>"
		else
		response.write "<img src=pic/ofFeMale.gif alt=美女呀，离线，快来找我吧！>"
		end if
	else
		if sex=1 then
		response.write "<img src=pic/Male.gif alt=帅哥哟，在线，有人找我吗？>"
		else
		response.write "<img src=pic/FeMale.gif alt=美女呀，在线，快来找我吧！>"
		end if
	end if
	response.write "</td><td width=16 valign=middle>"&userbirth&"</td>"
	response.write "</tr></table>"
	if Cint(GroupSetting(37))=1 and face<>"" then response.write "&nbsp;&nbsp;<img src="&htmlencode(FilterJS(face))&" width="&width&" height="&uheight&"><br>"
	response.write "&nbsp;&nbsp;<img src="&Forum_info(7)&rs("titlepic")&"><br>"
	if Cint(GroupSetting(36))=1 and title<>"" then response.write "&nbsp;&nbsp;头衔：" & htmlencode(title) & "<br>"
	response.write "&nbsp;&nbsp;等级："& userclass &"<BR>"
	if isnumeric(userPower) and userPower>0 then response.write "&nbsp;&nbsp;威望："&userPower&"<br>"
	response.write "&nbsp;&nbsp;文章："&article&"<br>"
	response.write "&nbsp;&nbsp;积分：" & userEP&"<br>"
	if UserGroup<>"" and not UserGroup="无门无派" then response.write "&nbsp;&nbsp;门派：" & htmlencode(UserGroup) & "<br>"

	response.write "&nbsp;&nbsp;注册："&FormatDateTime(addtime,2)&"<BR>"
	response.write "</td><td class="&bgcolor&" valign=top width=*><table width=100% ><tr><td width=*>"
	response.write "<a href=""javascript:openScript('messanger.asp?action=new&touser="&HTMLEncode(UserName)&"',500,400)""><img src="""&Forum_info(7)&Forum_TopicPic(7)&""" border=0 alt=""给"&HTMLEncode(UserName)&"发送一个短消息""></a>&nbsp;"
	response.write "<a href=""friendlist.asp?action=addF&myFriend="&HTMLEncode(UserName)&""" target=_blank><img src="""&Forum_info(7)&Forum_TopicPic(21)&""" border=0 alt=""把"&HTMLEncode(UserName)&"加入好友""></a>&nbsp;"
	response.write "<a href=""dispuser.asp?id="&postuserid&""" target=_blank><img src="""&Forum_info(7)&Forum_TopicPic(9)&""" border=0 alt=""查看"&HTMLEncode(UserName)&"的个人资料""></a>&nbsp;"
	response.write "<a href=""queryResult.asp?stype=1&nSearch=3&keyword="&HTMLEncode(UserName)&"&BoardID="&cstr(BoardID)&"&SearchDate=ALL"" target=_blank><img src="""&Forum_info(7)&Forum_TopicPic(8)&""" border=0 alt=""搜索"&HTMLEncode(UserName)&"在"&boardtype&"的所有贴子""></a>&nbsp;"

	if useremail<>"" then response.write "<A href=""mailto:"& htmlencode(useremail) &"""><IMG alt=""点击这里发送电邮给"& HTMLEncode(UserName) &""" border=0 src="&Forum_info(7)&Forum_TopicPic(10)&"></A>&nbsp;"

	if oicq<>"" then response.write "<a href=""http://search.tencent.com/cgi-bin/friend/user_show_info?ln="&oicq&""" target=_blank title="""&htmlencode(UserName)&"["&oicq&"]的QQ情况""><img src="&Forum_info(7)&Forum_TopicPic(11)&" width=16 height=16 border=0>OICQ</a>&nbsp;"

	if homepage<>"" then response.write "<A href="""& htmlencode(homepage) &""" target=_blank><IMG alt=""访问"& HTMLEncode(UserName) &"的主页"" border=0 src="&Forum_info(7)&Forum_TopicPic(14)&"></A>&nbsp;"

	response.write "<a href=""reannounce.asp?BoardID="&BoardID&"&replyID="&rs(0)&"&id="&AnnounceID&"&star="&star&"&reply=true""><img src="""&Forum_info(7)&Forum_TopicPic(15)&""" border=0 alt=引用回复这个贴子></a>&nbsp;"
	response.write "<a href=""reannounce.asp?BoardID="&BoardID&"&replyID="&rs(0)&"&id="&AnnounceID&"&star="&star&"""><img src="""&Forum_info(7)&"reply_a.gif"" border=0 alt=回复这个贴子></a></td><td width=50><b>"
	if star=1 and page_count=0 then
	response.write "楼顶"
	else
	response.write "第<font color="&Forum_body(8)&">" &(star-1)*Cint(Forum_Setting(12))+page_count+1 & "</font>楼"
	end if
	response.write "</b></td></tr><tr><td bgcolor="&Forum_body(27)&" height=1 colspan=2></td></tr></table>"
	response.write "<blockquote>"
	if LockUser=2 then
		response.write "================================<br><font color=#00008B>该用户发言已被管理员屏蔽</font><br>================================"
	else
		if rs("isbest")=1 and Cint(GroupSetting(41))=0 then
		response.write "================================<br><font color=#00008B>您没有浏览该精华帖子的权限</font><br>================================"
		else
		response.write "<img src="&Forum_info(8)&rs(6)&" border=0 alt=发贴心情>&nbsp;<font style=""font-size:"&Forum_Setting(59)&"pt;line-height:"&Forum_Setting(60)&"pt""><b>"&htmlencode(rs(3)) &"</b><br>"&dvbcode(rs(5),reUserGroupID)&"</font>"
		end if
	end if
	if signflag=1 and lockuser=0 and Cint(GroupSetting(38))=1 then
		if sign<>"" then response.write "<p>----------------------------------------------<br>"&dvsigncode(sign,reUserGroupID)&"</p>"
	end if
	if Cint(GroupSetting(30))=0 then
		userip="*.*.*.*"
	else
		userip=rs("ip")
	end if
	response.write "</blockquote></td></tr>"
	response.write "<tr><td class="&bgcolor&" valign=middle align=center width=175><a href=look_ip.asp?boardid="&boardid&"&userid="&PostUserID&"&ip="&userip&"&action=lookip target=_blank><img align=absmiddle border=0 width=13 height=15 src="""&Forum_info(7)&Forum_TopicPic(20) & """ alt=""点击查看用户来源及管理<br>发贴IP："&userip&"""></a> "&rs(4)&"</td>"
	response.write "<td class="&bgcolor&" valign=middle width=*>"
	response.write "<table width=100% cellpadding=0 cellspacing=0><tr><td align=left valign=middle> "

	if founduser then
		if Cint(GroupSetting(10))=1 or Cint(GroupSetting(23))=1 then
			response.write "<a href=editannounce.asp?BoardID="&BoardID&"&replyID="&rs(0)&"&id="&AnnounceID&"&star="&star&"><img src="&Forum_info(7)&Forum_TopicPic(16)&" border=0 alt=编辑这个贴子 align=absmiddle></a>"
		end if
	end if
	if istopic then response.write "&nbsp;&nbsp;<a href=""postagree.asp?boardid="&boardid&"&id="&Announceid&"&isagree=1"" title=""同意该帖观点，给他一朵鲜花，将消耗您"&GroupSetting(47)&"点金钱"">鲜花</a>(<font color="&Forum_body(8)&">"&isagree&"</font>)&nbsp;&nbsp;<a href=""postagree.asp?boardid="&boardid&"&id="&Announceid&"&isagree=2"" title=""不同意该帖观点，给他一个鸡蛋，将消耗您"&GroupSetting(47)&"点金钱"">鸡蛋</a>(<font color="&Forum_body(8)&">"&noagree&"</font>)"

	response.write "</td><td align=right nowarp valign=bottom width=110>"

	if Cint(GroupSetting(11))=1 or Cint(GroupSetting(18))=1 then
		if rs("isbest")=0 then
			besttitle="精华"
		else
			bestTitle="解除精华"
		end if
		if not istopic then
			response.write "<a href=""admin_postings.asp?action=删除跟帖&BoardID="&BoardID&"&ID="&AnnounceID&"&replyID="&rs(0)&"&userid="&postuserid&""" title=注意：本操作将删除单个贴子><img src="&Forum_info(7)&Forum_TopicPic(17)&" border=0></a>&nbsp;"
		end if
	response.write "<a href=""admin_postings.asp?action=复制&BoardID="&BoardID&"&ID="&AnnounceID&"&replyID="&rs(0)&""" title=""复制单个贴子到别的版面""><img src="""&Forum_info(7)&Forum_TopicPic(18)&""" border=0></a>&nbsp;"
	response.write "<a href=""admin_postings.asp?action="&bestTitle&"&BoardID="&BoardID&"&ID="&AnnounceID&"&replyID="&rs(0)&""" title="""&bestTitle&"""><img src="""&Forum_info(7)&Forum_TopicPic(19)&""" border=0></a>"
	end if
	response.write "</td>"
	response.write "<td align=right valign=bottom width=4> </td>"
	response.write "</tr></table>"
	response.write "</td></tr></table>"
page_count = page_count + 1
isagree=""
istopic=false
rs.movenext
loop
else
	ErrMsg=ErrMsg+"<br>"+"<li>您指定的贴子不存在</li>"
	exit sub
end if

response.write "<BR><table cellpadding=0 cellspacing=3 border=0 width="&Forum_body(12)&" align=center><tr><td valign=middle nowrap>本主题贴数<b>"&TopicCount&"</b>，分页："

if star > 4 then
	response.write "<a href=""?BoardID="&BoardID&"&id="&request("ID")&"&replyID="&replyID&"&star=1&skin="&request("skin")&""">[1]</a> ..."
end if
if Pcount>star+3 then
	endpage=star+3
else
	endpage=Pcount
end if
for i=star-3 to endpage
	if not i<1 then
		if i = clng(star) then
        response.write " <font color="&Forum_body(8)&">["&i&"]</font>"
		else
        response.write " <a href=""?BoardID="&BoardID&"&id="&request("ID")&"&replyID="&replyID&"&star="&i&"&skin="&request("skin")&""">["&i&"]</a>"
		end if
	end if
next
if star+3 < Pcount then 
	response.write "... <a href=""?BoardID="&BoardID&"&id="&request("ID")&"&replyID="&replyID&"&star="&Pcount&"&skin="&request("skin")&""">["&Pcount&"]</a>"
end if
rs.close
set rs=nothing

response.write "</td><td valign=middle nowrap align=right><select onchange=""if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}""><option selected>跳转论坛至...</option>"

dim rs1,sql1
sql="select id,class from class order by orders"
set rs=conn.execute(sql)
if not(rs.eof and rs.bof) then
do while not rs.eof
	response.write "<option>╋ "& rs(1) &"</option>"

	sql1="select boardid,boardtype from board where "&guestlist&" class="& rs(0)&" order by orders"
	set rs1=conn.execute(sql1)
	if rs1.eof and rs1.bof then
		response.write "<option>没有论坛</option>"
	else
		do while not rs1.eof
			response.write "<option value=""list.asp?boardid="&rs1(0)&""">　├" & rs1(1) & "</option>"
		rs1.movenext
		loop
	end if
rs.movenext
loop
end if
set rs=nothing
set rs1=nothing
response.write "</select></td></tr></table>"

if cint(skin)=1 then
%>
<br>
<table cellpadding=3 cellspacing=1 class=tableborder1 align=center>
<tr><th align=left width=90% valign=middle> &nbsp;*树形目录</th>
<th width=10% align=right valign=middle height=24 id=tabletitlelink> <a href=#top><img src=pic/gotop.gif border=0>顶端</a>&nbsp;</th></tr>
<%
dim outtext,bytestr,fontcolor
set rs=server.createobject("adodb.recordset")
sql="select AnnounceID,parentID,BoardID,UserName,PostUserid,Topic,DateAndTime,length,RootID,layer,orders,Expression,body from "&TotalUseTable&" where BoardID="&cstr(BoardID)&" and RootID="&Announceid&" and not Locktopic=2 order by RootID desc,orders"
	rs.open sql,conn,1,1
	if not(rs.eof and rs.bof) then
	do while not rs.eof
		if bgcolor="tablebody1" then
			bgcolor="tablebody2"
		else
			bgcolor="tablebody1"
		end if         
		if clng(replyid)=rs(0) then
			fontcolor="<font color="&Forum_body(8)&">"
		else
			fontcolor=""
		end if
		bytestr="("+cstr(rs("length"))
		if rs("Length")-1=1 then
			bytestr=bytestr+"字)"
		else
			bytestr=bytestr+"字)"
		end if
		if rs("length")=0 then
			bytestr="(空)"
		end if
		if rs("layer")>1 then
			for i=2 to rs("layer")
			outtext=outtext & " &nbsp; &nbsp; "
			next
			outtext=outtext & "回复："
		else
			outtext=outtext & "主题："
		end if
		outtext=outtext & "&nbsp;<img src="&Forum_info(8)&rs("Expression")&" height=16 width16>  "
		outtext=outtext &  "<a href='dispbbs.asp?BoardID="&BoardID&"&ID="&cstr(rs("RootID"))&"&replyID="&Cstr(rs("AnnounceID"))&"&skin=1'>" & fontcolor
		if rs("topic")="" or isnull(rs("topic")) then
			if len(rs("body"))>35 then
			outtext=outtext & mid(htmlencode(replace(rs("body"),chr(10),"")),1,49)+".."
			else
			outtext=outtext & htmlencode(replace(rs("body"),chr(10),""))
			end if
		else
			if len(rs("Topic"))>35 then
			outtext=outtext & mid(htmlencode(rs("Topic")),1,49)+".."
			else
			outtext=outtext & htmlencode(rs("Topic"))
			end if
		end if
		response.write "<TR><td class="&bgcolor&" width=""100%"" height=22 colspan=2>"&outtext&"</a><I><font color=gray>"&bytestr&" － <a href=dispuser.asp?name="&htmlencode(rs("UserName"))&" target=_blank title='作者资料'><font color=gray>"&HTMLEncode(rs("UserName"))&"</font></a>，" & Formatdatetime(rs("DateAndTime"),1) & "</font></I></td></tr>"
		rs.movenext
		outtext=""
		loop
	end if
	rs.close
	set rs=nothing
response.write "</table><br>"
end if

if canreply then
%>
<table cellpadding=1 cellspacing=1 class=tableborder1 align=center>
<tr>
<th align=left colspan=2 width=100% valign=middle height=25> &nbsp;*快速回复：<%=htmlencode(topic_1)%></th></tr>
<form action="SaveReAnnounce.asp?method=fastreply&BoardID=<%=BoardID%>" method=POST  name=frmAnnounce onSubmit=submitonce(this)>
<input type=hidden name="followup" value="<%=replyid%>">
<input type=hidden name="RootID" value="<%=AnnounceID%>">
<input type=hidden name="star" value="<%=star%>">
<input type=hidden name="TotalUseTable" value="<%=TotalUseTable%>">
<TR><TD noWrap width=175 class=tablebody2><b>你的用户名:</b></TD>
<TD height=30 class=tablebody2><INPUT maxLength=25 size=23 value="<%=membername%>" name=UserName>
&nbsp;&nbsp; <A href="reg.asp">还没注册?</A>  &nbsp;&nbsp;  <b>密码:</b>
<INPUT type=password maxLength=20 size=23 value="<%=memberword%>" name=passwd>
&nbsp;&nbsp; <A href="lostpass.asp">忘记密码?</A></TD></TR>
<TR> <TD vAlign=top noWrap class=tablebody1><b>内容</b><br>
<li>HTML标签： <%=iif(Forum_Setting(35),"可用","不可用")%>
<li>UBB标签： <%=iif(Forum_Setting(34),"可用","不可用")%>
<li>贴图标签： <%=iif(Forum_Setting(37),"可用","不可用")%>
<li>Flash标签：<%=iif(Forum_Setting(38),"可用","不可用")%>
<li>表情字符转换：<%=iif(Forum_Setting(36),"可用","不可用")%>
<li>上传图片：<%=iif(Forum_Setting(3),"可用","不可用")%>
<li>最多<%=Forum_Setting(13)\1024%>KB 
</TD>
<TD class=tablebody1> 
<TEXTAREA name=Content cols=100 rows=7 wrap=VIRTUAL title=可以使用Ctrl+Enter直接提交贴子  onkeydown=ctlent()></TEXTAREA>
</TD></TR>
<TR bgColor=<%=Forum_body(4)%>>
<TD noWrap class=tablebody2 height=30>
<INPUT type=checkbox value=yes name=Forum_Setting(2)>
邮件回复 <INPUT type=checkbox CHECKED value=yes name=signflag>
显示签名 
</TD>
<TD width="100%" class=tablebody2> 
<input type=Submit value="OK!发表我的回应帖子" name=Submit>
&nbsp;<input type=button value="预 览" name=Button onclick=gopreview()>&nbsp;<input type=reset name=Clear value=清空内容！>[Ctrl+Enter直接提交贴子] 
</TD>
</TR></FORM>
</TABLE>
<%
end if
if istopic or skin=0 then
	if istop=0 then istoptitle="固顶" else istoptitle="解固"
%><BR>
<TABLE cellSpacing=0 cellPadding=0 width="<%=Forum_body(12)%>" border=0 align=center>
<tr valign=center> <td width =100% align=right> 
<a href="admin_postings.asp?action=锁定&BoardID=<%=BoardID%>&ID=<%=Announceid%>" title=锁定本主题>锁定</a> 
  | <a href="admin_postings.asp?action=解锁&BoardID=<%=BoardID%>&ID=<%=Announceid%>" title=将本主题解开锁定>解锁</a>
  | <a href="admin_postings.asp?action=删除主题&BoardID=<%=BoardID%>&ID=<%=Announceid%>" title=注意：本操作将删除本主题所有贴子，不能恢复>删除</a>
  | <a href="admin_postings.asp?action=移动&BoardID=<%=BoardID%>&ID=<%=Announceid%>&replyID=<%=replyID%>" title=移动主题>移动</a>  |  <a href="admin_postings.asp?action=<%=istopTitle%>&BoardID=<%=BoardID%>&ID=<%=Announceid%>" title=将本主题<%=istopTitle%>><%=istopTitle%></a>
  | <a href="admin_postings.asp?action=奖励&BoardID=<%=BoardID%>&ID=<%=Announceid%>" title=对发起主题用户奖励>奖励</a>
  | <a href="admin_postings.asp?action=惩罚&BoardID=<%=BoardID%>&ID=<%=Announceid%>" title=对发起主题用户惩罚>惩罚</a>
  | <a href="announcements.asp?BoardID=<%=BoardID%>&action=AddAnn">发布公告</a>
</td></tr></table>
<%end if%>
<form name=preview action=preview.asp method=post target=preview_page>
<input type=hidden name=title value=""><input type=hidden name=body value="">
</form>
<script>
function gopreview()
{
<%if voteyes=0 then%>
document.forms[1].body.value=document.forms[0].Content.value;
var popupWin = window.open('preview.asp', 'preview_page', 'scrollbars=yes,width=750,height=450');
document.forms[1].submit()
<%else%>
document.forms[2].body.value=document.forms[1].Content.value;
var popupWin = window.open('preview.asp', 'preview_page', 'scrollbars=yes,width=750,height=450');
document.forms[2].submit()
<%end if%>
}
</script>
<script language=javascript>
ie = (document.all)? true:false
if (ie){
function ctlent(eventobject){if(event.ctrlKey && window.event.keyCode==13){this.document.frmAnnounce.submit();}}
}

function submitonce(theform){
//if IE 4+ or NS 6+
if (document.all||document.getElementById){
//screen thru every element in the form, and hunt down "submit" and "reset"
for (i=0;i<theform.length;i++){
var tempobj=theform.elements[i]
if(tempobj.type.toLowerCase()=="submit"||tempobj.type.toLowerCase()=="reset")
//disable em
tempobj.disabled=true
}
}
}
</script>
<%
call activeonline()
end sub

sub readRe()
	dim rs1,ID
	set rs1=conn.execute("select reAnn from [user] where UserID="&UserID&" and reAnn is not null")
	if not (rs1.eof and rs1.bof) then
		ID=split(rs1("reAnn"),"|")(1)
		if clng(ID)=clng(AnnounceID) then
			conn.execute ("update [user] set reAnn=null where UserID="&UserID)
		end if
	end if
	rs1.close
	set rs1=nothing
end sub

function isOnline(Userid,sex)
	dim ors
	set ors=conn.execute("select username from online where userid="&userid&" "&userhiddensql&"")
	if ors.eof and ors.bof then
		if sex=1 then
		isonline="<img src=pic/ofMale.gif alt=帅哥哟，离线，有人找我吗？>"
		else
		isonline="<img src=pic/ofFeMale.gif alt=美女呀，离线，快来找我吧！>"
		end if
	else
		if sex=1 then
		isonline="<img src=pic/Male.gif alt=帅哥哟，在线，有人找我吗？>"
		else
		isonline="<img src=pic/FeMale.gif alt=美女呀，在线，快来找我吧！>"
		end if
	end if
	ors.close
	set ors=nothing
end function
%>