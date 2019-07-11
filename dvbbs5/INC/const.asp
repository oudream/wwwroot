<!--#include file=char.asp-->
<%
'论坛Cookies名称，一般不需要修改，当您的网站启用两个以上的论坛时必须将这些论坛系统的Cookies名称修改为不同
Dim Forum_sn
Forum_sn="aspsky"
if Request.Cookies("iscookies")="" then
	Response.Cookies("iscookies")="0"
	Response.Cookies("iscookies").Expires=date+3650
	response.write "<META http-equiv=Content-Type content=text/html; charset=gb2312><meta HTTP-EQUIV=REFRESH CONTENT=3>正在登陆论坛……<br><br>本系统要求使用COOKIES，假如您的浏览器禁用COOKIES，您将不能登录本系统……"
	response.end
end if

dim UserAgent,Stats,ScriptName
UserAgent=Trim(lcase(Request.Servervariables("HTTP_USER_AGENT")))
ScriptName=lcase(request.ServerVariables("PATH_INFO"))
If Instr(UserAgent,"teleport")>0 or Instr(UserAgent,"webzip")>0 or Instr(UserAgent,"flashget")>0 or Instr(UserAgent,"offline")>0 Then
	response.redirect "error.htm"
	response.end
end if

Dim copyright,Version,Cookiepath,Badwords,Splitwords,StopReadme,Maxonline,MaxonlineDate
Dim LastTopicNum,LastPostTime,LastPostInfo,lastbbsnum
Dim guestlist
Dim boardtype,lockboard,master_2,todaynum,master_1,boardmasterlist
Dim skinid,skin,tid,tempid
Dim templateused
Dim Forum_info,Forum_setting,Forum_ubb,Forum_body,Forum_ads,Forum_user
Dim Forum_pic,Forum_boardpic,Forum_TopicPic,Forum_statePic,Forum_upload
Dim MaxOldID,NowUseBbs,AllPostTable,AllPostTableName,BoardReadme
Dim sqlstr
dim mastermsg
dim Forum_getre,Forum_smsmsg
dim backpage,navinfo
dim Forum_header
dim membername,memberword,memberclass,userhidden
dim sql,rs
dim BoardID,FoundBoard,i
Dim Founderr,Errmsg,sucmsg
Dim Founduser,userid
Dim FoundStyle
Dim UserPermission,FoundUserPer
FoundUserPer=false
FoundBoard=false
Founduser=false
Founderr=false
FoundStyle=false
if request("BoardID")="" or (not isInteger(request("BoardID"))) or request("boardid")="0" or instr(scriptname,"index.asp")>0 then
	BoardID=0
	FoundBoard=false
else
	BoardID=Clng(Request("BoardID"))
	FoundBoard=true
end if
membername=checkStr(request.cookies("aspsky")("username"))
memberword=checkStr(request.cookies("aspsky")("password"))
memberclass=checkStr(request.cookies("aspsky")("userclass"))
userhidden=request.cookies("aspsky")("userhidden")
if request("tempid")<>"" and isnumeric(request("tempid")) then
	tempid=request("tempid")
	foundstyle=true
elseif request.cookies("aspsky")("tempid")<>"" and isnumeric(request.cookies("aspsky")("tempid")) then
	tempid=request.cookies("aspsky")("tempid")
	foundstyle=true
else
	tempid=0
	foundstyle=false
end if
userid=request.cookies("aspsky")("userid")
if not isnumeric(userhidden) or userhidden="" then userhidden=2
'论坛基本变量调用
if not FoundBoard then
	if not isnumeric(request("skinid")) or request("skinid")="" then
		sqlstr=" active=1"
	else
		sqlstr=" id="&request("skinid")&" "
	end if
	sql = "select * from config where "&sqlstr&""
else
	sql = "select b.boardtype,b.lockboard,b.boardmaster,b.todaynum,b.lasttopicnum,b.lastbbsnum,b.LastPost,b.readme,c.* from board b inner join config c on b.sid=c.id where b.boardid="&boardid
end if
set rs=conn.execute(sql)
if rs.eof and rs.bof then
	response.write "错误的系统参数，很可能您指定了错误的论坛名称，请从有效连接进入。"
	response.end
else
	Forum_info=split(rs("Forum_info"),",")
	Forum_Setting=split(rs("Forum_setting"),",")
	Forum_ads=split(rs("Forum_ads"),"$")
	Forum_body=split(rs("Forum_body"),"|||")
	Forum_user=split(rs("Forum_user"),",")

	Copyright=rs("Forum_copyright")
	Version="   Powered by：<a href=http://www.dvbbs.net/download.asp>DVBBS Ver5.0 Final</a>"
	badwords=rs("badwords")
	Splitwords=rs("Splitwords")
	StopReadme=rs("StopReadme")
	Forum_pic=split(rs("Forum_pic"),",")
	Forum_boardpic=split(rs("Forum_boardpic"),",")
	Forum_TopicPic=split(rs("Forum_Topicpic"),",")
	Forum_statePic=split(rs("Forum_statepic"),",")
	Forum_upload=rs("Forum_upload")
	Forum_ubb=split(rs("Forum_ubb"),",")
	Forum_sn=rs("Forum_sn")
	MaxOldID=rs("Maxoldid")
	NowUseBbs=rs("NowUseBbs")
	AllPostTable=split(rs("AllPostTable"),"|")
	AllPostTableName=split(rs("AllPostTableName"),"|")

	if not FoundBoard then
		LastPostInfo=split(rs("LastPost"),"$")
		if not isdate(LastPostInfo(2)) then LastPostInfo(2)=Now()
		LastPostTime=LastPostInfo(2)
		cookiepath=rs("cookiepath")
		Maxonline=rs("Maxonline")
		MaxonlineDate=rs("MaxonlineDate")
		skinid=rs("id")
		tid=rs("tid")
		stats="论坛信息"
		if datediff("d",LastPostTime,Now())>0 then
			conn.execute("update config set yesterdaynum=todaynum,LastPost='"&LastPostInfo(0)&"$"&LastPostInfo(1)&"$"&Now()&"$"&LastPostInfo(3)&"$"&LastPostInfo(4)&"$"&LastPostInfo(5)&"'")
			conn.execute("update config set todaynum=0")
		end if
	else
		LastPostInfo=split(rs(6),"$")
		if not isdate(LastPostInfo(2)) then LastPostInfo(2)=Now()
		LastPostTime=LastPostInfo(2)
		boardtype=rs("boardtype")
		lockboard=rs("lockboard")
		boardmasterlist=rs("boardmaster")
		todaynum=rs(3)
		LastTopicNum=rs("LastTopicNum")
		LastBbsNum=rs("LastBbsNum")
		tid=rs("tid")
		stats=boardtype
		BoardReadme=rs("Readme")
		if datediff("d",LastPostTime,Now())>0 then
			conn.execute("update board set todaynum=0 where boardid="&boardid)
		end if
	end if
end if
rs.close
set rs=nothing

if (instr(scriptname,"list.asp")>0 and instr(scriptname,"uploadlist")=0 and instr(scriptname,"friendlist")=0 and instr(scriptname,"favlist")=0) or instr(scriptname,"dispbbs")>0 then
if cint(Forum_Setting(19))=1 then
	if not isnull(session("ReflashTime")) and cint(Forum_Setting(20))>0 then
		if DateDiff("s",session("ReflashTime"),Now())<cint(Forum_Setting(20)) then
   		response.write "<META http-equiv=Content-Type content=text/html; charset=gb2312><meta HTTP-EQUIV=REFRESH CONTENT=3>正在打开页面，请稍后……"
		response.end
		else
		session("ReflashTime")=Now()
		end if
	end if
end if
end if

if (instr(scriptname,"admin")=0 and instr(scriptname,"login")=0 and instr(scriptname,"chklogin")=0) or master then
	if cint(Forum_Setting(21))=1 then
	Response.write StopReadme
	response.end
	end if

	if LockIP(Request.ServerVariables("REMOTE_ADDR")) then
	response.write "您的IP已经被限制不能访问本论坛，请和管理员联系"&Forum_info(5)&"。"
	response.end
	end if
end if

Server.ScriptTimeOut=Forum_Setting(1)

Rem 用户信息
dim vipuser,boardmaster,master,superboardmaster
dim reboardid,reid
dim LastLogin,myuserep,mysex,myusercp,mymoney,mypower,myArticle,myClass
dim UserGroupID,GroupSetting
vipuser=false
boardmaster=false
superboardmaster=false
master=false
if userid<>"" and isInteger(userid) then
	sql="select u.userclass,u.userid,u.userpassword,u.lastlogin,u.showRe,u.reAnn,u.userGroupID,g.GroupSetting,u.username,u.userWealth,u.Userep,u.Usercp,u.sex,u.UserPower,u.Article from [user] u inner join UserGroups G on u.UserGroupID=g.UserGroupID where u.userid="&userid
	set rs=conn.execute(sql)
	if rs.eof and rs.bof then
		founduser=false
	else
		if trim(rs(2))=trim(memberword) and lcase(trim(membername))=lcase(trim(rs(8))) then
			founduser=true
			select case rs(6)
			case 8
				vipuser=true
			case 3
				boardmaster=true
			case 2
				superboardmaster=true
			case 1
				master=true
			end select
			userid=rs(1)
			lastlogin=rs(3)
			if rs(4)=1 and rs(5)<>"" then
			reBoardID=split(rs(5),"|")(0)
			reID=split(rs(5),"|")(1)
			end if
			UserGroupID=rs(6)
			myuserep=rs(10)
			mysex=rs(12)
			myusercp=rs(11)
			mymoney=rs(9)
			mypower=rs(13)
			myArticle=rs(14)
			myClass=rs(0)
			GroupSetting=split(rs(7),",")
			if userhidden=2 and DateDiff("s",rs(3),Now())>Cint(Forum_Setting(8))*60 then
			conn.execute("update [user] set UserLastIp='"&Request.ServerVariables("REMOTE_ADDR")&"',LastLogin=Now() where userid="&userid)
			end if
		else
			founduser=false
		end if
	end if
	rs.close
	set rs=nothing
end if
if not founduser then
	founduser=false
	userid=0
	set rs=conn.execute("select GroupSetting from UserGroups where UserGroupID=7")
	GroupSetting=split(rs(0),",")
	UserGroupID=7
	rs.close
	set rs=nothing
end if
if boardid>=0 then GetBoardPermission
if FoundBoard and FoundUser then
	if master then
		boardmaster=true
	else
		if instr("|"&boardmasterlist&"|","|"&membername&"|")>0 then
		boardmaster=true
		else
		boardmaster=false
		end if
	end if
end if
'endtime=timer()
'response.write "页面执行时间："&FormatNumber((endtime-startime)*1000,3)&"毫秒"
'response.end
%>