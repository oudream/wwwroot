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
	stats="����ͶƱ"

	if BoardID="" or not isInteger(BoardID) or BoardID=0 then
		Errmsg=Errmsg+"<br>"+"<li>����İ����������ȷ�����Ǵ���Ч�����ӽ��롣"
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

	rem -----���user�������ݵĺϷ���------	
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
   		ErrMsg=ErrMsg+"<Br>"+"<li>����̳���Ʒ�������ʱ��Ϊ10�룬���Ժ��ٷ���"
   		FoundErr=True
		end if
	end if
	end if
	if chkpost=false then
   		ErrMsg=ErrMsg+"<Br>"+"<li>���ύ�����ݲ��Ϸ����벻Ҫ���ⲿ�ύ���ԡ�"
   		FoundErr=True
	end if
	if UserName="" or UserPassWord="" then
		username=membername
		UserPassWord=memberword
	end if
	if UserName="" then
   		ErrMsg=ErrMsg+"<Br>"+"<li>����������(���Ȳ��ܴ���20)"
   		FoundErr=True
	end if
	if Topic="" then
   		FoundErr=True
		ErrMsg=ErrMsg+"<Br>"+"<li>���ⲻӦΪ�ա�"
	elseif strLength(topic)>50 then
   		FoundErr=True
		ErrMsg=ErrMsg+"<Br>"+"<li>���ⳤ�Ȳ��ܳ���50"
	end if
	if strLength(body)>Forum_Setting(13) then
   		ErrMsg=ErrMsg+"<Br>"+"<li>�������ݲ��ô���" & CSTR(Forum_Setting(13)) & "bytes"
   		FoundErr=true
	end if
   	if body="" then
		ErrMsg=ErrMsg+"<Br>"+"<li>û����д���ݡ�"
   		FoundErr=true
	end if
	if vote="" then
		ErrMsg=ErrMsg+"<Br>"+"<li>������ͶƱ����"
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
   		ErrMsg=ErrMsg+"<Br>"+"<li>�������ݲ��ô���" & CSTR(Forum_Setting(13)) & "bytes"
   		FoundErr=true
	else
		if request("votetimeout")="0" then
		votetimeout=dateadd("d",9999,Now())
		else
		votetimeout=dateadd("d",request("votetimeout"),Now())
		end if
		votetimeout=replace(replace(CSTR(votetimeout+Forum_Setting(0)/24),"����",""),"����","")
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
		Errmsg=Errmsg+"<br>"+"<li>��û���ڱ���̳����ͶƱ��Ȩ�ޣ���<a href=login.asp>��½</a>����ͬ����Ա��ϵ��"
		founderr=true
		exit sub
	end if
	if cint(lockboard)=2 then
		if not founduser then
			Errmsg=Errmsg+"<br>"+"<li>����̳Ϊ��֤��̳����<a href=login.asp>��½</a>��ȷ�������û����Ѿ��õ�����Ա����֤����롣"
			founderr=true
			exit sub
		else
			if chkboardlogin(boardid,membername)=false then
			Errmsg=Errmsg+"<br>"+"<li>����̳Ϊ��֤��̳����ȷ�������û����Ѿ��õ�����Ա����֤����롣"
			founderr=true
			exit sub
			end if
		end if
	end if

	usercookies=request.Cookies("aspsky")("usercookies")
	if isnull(usercookies) or usercookies="" then usercookies=3

	if chkuserlogin(username,userpassword,usercookies,2)=false then
		errmsg=errmsg+"<br>"+"<li>�����û����������ڣ���������������󣬻��������ʺ��ѱ�����Ա������"
		founderr=true
		exit sub
	end if

	if lockboard=3 then
		if not master then
			Errmsg=ErrMsg+"<Br>"+"<li>��û��Ȩ���ڱ����淢�����ӣ�"
			founderr=true
			exit sub
		end if
	end if

	dim LastPost,LastPost_1,uploadpic_n,Forumupload,u
	dim LastPostTimes,voteid
	DateTimeStr=replace(replace(CSTR(NOW()+Forum_Setting(0)/24),"����",""),"����","")

	'����ͶƱ��¼
	conn.execute("insert into vote (vote,votenum,votetype,timeout) values ('"&vote&"','"&votenum&"',"&votetype&",'"&votetimeout&"')")
	set rs=conn.execute("select top 1 voteid from vote order by voteid desc")
	voteid=rs(0)
	'���������
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
	'����ظ���
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
	PostRetrunName="��ҳ"
	case 2
	response.write "<meta http-equiv=refresh content=""3;URL=list.asp?boardid="&boardid&""">"
	PostRetrunName="������������̳"
	case 3
	response.write "<meta http-equiv=refresh content=""3;URL=dispbbs.asp?boardid="&boardid&"&id="&rootid&""">"
	PostRetrunName="�������������"
	end select
	set rs=nothing
%>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
<tr align=center><th width="100%">״̬��<%=stats%></td>
</tr><tr><td width="100%" class=tablebody1>
��ҳ�潫��3����Զ�����<%=PostRetrunName%>��<b>������ѡ�����²�����</b><br><ul>
<li><a href="index.asp">������ҳ</a></li>
<li><a href="list.asp?boardid=<%=boardid%>"><%=boardtype%></a></li>
<li><a href="dispbbs.asp?boardid=<%=boardid%>&id=<%=rootid%>&star=<%=request("star")%>#<%=announceid%>"><%=PostRetrunName%></a></li>
</ul></td></tr></table>
<%
end sub
%>