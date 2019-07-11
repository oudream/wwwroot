<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/char_board.asp" -->
<!-- #include file="inc/char_login.asp" -->
<!-- #include file="inc/chkinput.asp" -->
<!--#include file="md5.asp"-->
<%
	dim announceid
	dim UserName
	dim userPassword
	dim useremail
	dim Topic
	dim body
	dim dateTimeStr
	dim ip
	dim Expression
	dim signflag
	dim mailflag
	dim boardstat
	dim usercookies
	dim votetype,vote,votenum
	dim vote_1,votelen,votenumlen,j
	dim votetimeout
	dim rootid
	stats="发表投票"

	if BoardID="" or not isInteger(BoardID) or BoardID=0 then
		Errmsg=Errmsg+"<br>"+"<li>错误的版面参数！请确认您是从有效的连接进入。"
		founderr=true
	else
		BoardID=clng(BoardID)
	end if
	IP=Request.ServerVariables("REMOTE_ADDR") 
	Expression=Checkstr(Request.Form("Expression")&".gif")
   	Topic=Checkstr(trim(request.Form("subject")))
   	Body=Checkstr(trim(request.Form("Content")))
   	UserName=Checkstr(trim(request.Form("username")))
	signflag=Checkstr(trim(request.Form("signflag")))
	mailflag=Checkstr(trim(request.Form("Forum_Setting(2)")))
	if memberword=trim(request.Form("passwd")) then
	UserPassWord=Checkstr(trim(request.Form("passwd")))
	else
   	UserPassWord=md5(Checkstr(trim(request.Form("passwd"))))
	end if
	votetype=Checkstr(request.Form("votetype"))
	vote=Checkstr(trim(replace(request.Form("vote"),"|","")))

	rem -----检查user输入数据的合法性------	
	dim num1,rndnum,k
	if instr(Expression,"face")=0 then
	Expression="face0.gif"
	end if
	if signflag="yes" then
		signflag=1
	else
		signflag=0
	end if
	if mailflag="yes" then
		mailflag=1
	else
		mailflag=0
	end if
	if cint(Forum_Setting(42))=1 then
	if not (isnull(session("lastpost")) or boardmaster or master) then
		if DateDiff("s",session("lastpost"),Now())<cint(Forum_Setting(43)) then
   		ErrMsg=ErrMsg+"<Br>"+"<li>本论坛限制发贴距离时间为10秒，请稍后再发。"
   		FoundErr=True
		end if
	end if
	end if
	if chkpost=false then
   		ErrMsg=ErrMsg+"<Br>"+"<li>您提交的数据不合法，请不要从外部提交发言。"
   		FoundErr=True
	end if
	if UserName="" or UserPassWord="" then
		username=membername
		UserPassWord=memberword
	end if
	if UserName="" then
   		ErrMsg=ErrMsg+"<Br>"+"<li>请输入姓名(长度不能大于20)"
   		FoundErr=True
	end if
	if Topic="" then
   		FoundErr=True
		ErrMsg=ErrMsg+"<Br>"+"<li>主题不应为空。"
	elseif strLength(topic)>50 then
   		FoundErr=True
		ErrMsg=ErrMsg+"<Br>"+"<li>主题长度不能超过50"
	end if
	if strLength(body)>Forum_Setting(13) then
   		ErrMsg=ErrMsg+"<Br>"+"<li>发言内容不得大于" & CSTR(Forum_Setting(13)) & "bytes"
   		FoundErr=true
	end if
   	if body="" then
		ErrMsg=ErrMsg+"<Br>"+"<li>没有填写内容。"
   		FoundErr=true
	end if
	if vote="" then
		ErrMsg=ErrMsg+"<Br>"+"<li>请输入投票内容"
   		FoundErr=true
	else
		vote=split(vote,chr(13)&chr(10))
		j=0
		for i = 0 to ubound(vote)
			if not (vote(i)="" or vote(i)=" ") then
				vote_1=""&vote_1&""&vote(i)&"|"
				j=j+1
			end if
			if i>cint(Forum_Setting(39))-2 then exit for
		next
		for k = 1 to j
			votenum=""&votenum&"0|"
		next
		votelen=len(vote_1)
		votenumlen=len(votenum)
		votenum=left(votenum,votenumlen-1)
		vote=left(vote_1,votelen-1)
	end if
	if not isnumeric(request("votetimeout")) then
   		ErrMsg=ErrMsg+"<Br>"+"<li>发言内容不得大于" & CSTR(Forum_Setting(13)) & "bytes"
   		FoundErr=true
	else
		if request("votetimeout")="0" then
		votetimeout=dateadd("d",9999,Now())
		else
		votetimeout=dateadd("d",request("votetimeout"),Now())
		end if
		votetimeout=replace(replace(CSTR(votetimeout+Forum_Setting(0)/24),"上午",""),"下午","")
	end if
	session("lastpost")=Now()

	if founderr then
		call nav()
		call head_var(1)
		call dvbbs_error()
	else
		call nav()
		call head_var(2)
		call main()
		if founderr then call dvbbs_error()
	end if
	call footer()

sub main()
	if Cint(GroupSetting(8))=0 then
		Errmsg=Errmsg+"<br>"+"<li>您没有在本论坛发表投票的权限，请<a href=login.asp>登陆</a>或者同管理员联系。"
		founderr=true
		exit sub
	end if
	if cint(lockboard)=2 then
		if not founduser then
			Errmsg=Errmsg+"<br>"+"<li>本论坛为认证论坛，请<a href=login.asp>登陆</a>并确认您的用户名已经得到管理员的认证后进入。"
			founderr=true
			exit sub
		else
			if chkboardlogin(boardid,membername)=false then
			Errmsg=Errmsg+"<br>"+"<li>本论坛为认证论坛，请确认您的用户名已经得到管理员的认证后进入。"
			founderr=true
			exit sub
			end if
		end if
	end if

	usercookies=request.Cookies("aspsky")("usercookies")
	if isnull(usercookies) or usercookies="" then usercookies=3

	if chkuserlogin(username,userpassword,usercookies,2)=false then
		errmsg=errmsg+"<br>"+"<li>您的用户名并不存在，或者您的密码错误，或者您的帐号已被管理员锁定。"
		founderr=true
		exit sub
	end if

	if lockboard=3 then
		if not master then
			Errmsg=ErrMsg+"<Br>"+"<li>您没有权限在本版面发布贴子！"
			founderr=true
			exit sub
		end if
	end if

	dim LastPost,LastPost_1,uploadpic_n,Forumupload,u
	dim LastPostTimes,voteid
	DateTimeStr=replace(replace(CSTR(NOW()+Forum_Setting(0)/24),"上午",""),"下午","")

	'插入投票记录
	conn.execute("insert into vote (vote,votenum,votetype,timeout) values ('"&vote&"','"&votenum&"',"&votetype&",'"&votetimeout&"')")
	set rs=conn.execute("select top 1 voteid from vote order by voteid desc")
	voteid=rs(0)
	'插入主题表
	sql="insert into topic (Title,Boardid,PostUsername,PostUserid,DateAndTime,Expression,LastPost,LastPostTime,isvote,PollID,voteTotal,PostTable) values ('"&topic&"',"&boardid&",'"&username&"',"&userid&",'"&DateTimeStr&"','"&Expression&"','$$"&DateTimeStr&"$$$$','"&DateTimeStr&"',1,"&voteid&",0,'"&NowUseBbs&"')"
	conn.execute(sql)

	set rs=conn.execute("select top 1 topicid from topic order by topicid desc")
	rootid=rs(0)
	Forumupload=split(Forum_upload,",")
	for u=0 to ubound(Forumupload)
	if instr(body,"[upload="&Forumupload(u)&"]") and instr(body,"[/upload]")>0 then
		uploadpic_n=Forumupload(u)
		exit for
	end if
	next
	'插入回复表
	Sql="insert into "&NowUseBbs&"(Boardid,ParentID,username,topic,body,DateAndTime,length,rootid,layer,orders,ip,Expression,locktopic,signflag,emailflag,isbest,PostUserID) values ("&boardid&",0,'"&username&"','"&topic&"','"&body&"','"&DateTimeStr&"','"&strlength(body)&"',"&rootid&",1,0,'"&ip&"','"&Expression&"',0,"&signflag&","&mailflag&",0,"&userid&")"
	conn.execute(sql)

	set rs=conn.execute("select top 1 Announceid from "&NowUseBbs&" order by Announceid desc")
	Announceid=rs(0)
	LastPost=replace(username,"$","") & "$" & Announceid & "$" & DateTimeStr & "$" & replace(cutStr(body,20),"$","") & "$" & uploadpic_n & "$" & UserID & "$" & rootid

	conn.execute("update topic set LastPost='"&LastPost&"' where topicid="&rootid)

	LastPost_1=replace(username,"$","") & "$" & Announceid & "$" & DateTimeStr & "$" & replace(cutStr(Topic,20),"$","") & "$" & uploadpic_n & "$" & UserID & "$" & rootid

	if datediff("d",LastPostTime,Now())=0 then
		sql="update board set lastbbsnum=lastbbsnum+1,lasttopicnum=lasttopicnum+1,todaynum=todaynum+1,LastPost='"&LastPost_1&"' where boardid="&cstr(boardID)
	else
		sql="update board set lastbbsnum=lastbbsnum+1,lasttopicnum=lasttopicnum+1,todaynum=1,LastPost='"&LastPost_1&"' where boardid="&cstr(boardID)
	end if
	conn.execute(sql)

	Dim updateinfo
	set rs=conn.execute("select LastPost,TodayNum,MaxPostNum from config where active=1")
	LastPostTimes=split(rs(0),"$")
	LastPostTime=LastPostTimes(2)
	if not isdate(LastPostTime) then LastPostTime=Now()
	if datediff("d",LastPostTime,Now())=0 then
		if rs(1)+1>rs(2) then updateinfo=",MaxPostNum=todaynum+1"
		sql="update config set topicnum=topicnum+1,bbsnum=bbsnum+1,todayNum=todayNum+1,LastPost='"&LastPost&"' "&updateinfo&" where active=1"
	else
		sql="update config set topicnum=topicnum+1,bbsnum=bbsnum+1,yesterdaynum="&rs(1)&",todayNum=1,LastPost='"&LastPost&"' where active=1"
	end if
	conn.execute(sql)

	dim PostRetrunName,PostRetrun
	select case Forum_Setting(69)
	case 1
	response.write "<meta http-equiv=refresh content=""3;URL=index.asp"">"
	PostRetrunName="首页"
	case 2
	response.write "<meta http-equiv=refresh content=""3;URL=list.asp?boardid="&boardid&""">"
	PostRetrunName="您所发布的论坛"
	case 3
	response.write "<meta http-equiv=refresh content=""3;URL=dispbbs.asp?boardid="&boardid&"&id="&rootid&""">"
	PostRetrunName="您所发表的帖子"
	end select
	set rs=nothing
%>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
<tr align=center><th width="100%">状态：<%=stats%></td>
</tr><tr><td width="100%" class=tablebody1>
本页面将在3秒后自动返回<%=PostRetrunName%>，<b>您可以选择以下操作：</b><br><ul>
<li><a href="index.asp">返回首页</a></li>
<li><a href="list.asp?boardid=<%=boardid%>"><%=boardtype%></a></li>
<li><a href="dispbbs.asp?boardid=<%=boardid%>&id=<%=rootid%>&star=<%=request("star")%>#<%=announceid%>"><%=PostRetrunName%></a></li>
</ul></td></tr></table>
<%
end sub
%>