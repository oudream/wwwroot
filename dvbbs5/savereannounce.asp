<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/char_board.asp" -->
<!-- #include file="inc/char_login.asp" -->
<!-- #include file="inc/chkinput.asp" -->
<!-- #include file="inc/email.asp" -->
<!--#include file="md5.asp"-->
<%
	dim UserName
	dim userPassword
	dim useremail
	dim Topic,PostTopic
	dim body
	dim somerr
	dim dateTimeStr
	dim ParentID
	dim RootID
	dim iLayer
	dim iOrders
	dim ip
	dim announceid
	dim Expression
	dim signflag
	dim mailflag
	dim Email,mailbody,SendMail
	dim boardstat
	dim usercookies
	dim LastPost,LastPost_1,UpLoadPic_n
	dim rsLayer
	dim LastPostTimes
	dim toptopic
	dim TotalUseTable
	dim FoundTable
	FoundTable=false

	if BoardID="" or not isInteger(BoardID) or BoardID=0 then
		Errmsg=Errmsg+"<br>"+"<li>����İ����������ȷ�����Ǵ���Ч�����ӽ��롣"
		founderr=true
	else
		BoardID=clng(BoardID)
	end if
	if request("followup")="" then
		foundErr=true
		Errmsg=Errmsg+"<br>"+"<li>�Ƿ������Ӳ�����"
	elseif not isInteger(request("followup")) then
		foundErr=true
		Errmsg=Errmsg+"<br>"+"<li>�Ƿ������Ӳ�����"
	else
		announceid=request("followup")
		ParentID=request("followup")
	end if
	if request("RootID")="" then
		foundErr=true
		Errmsg=Errmsg+"<br>"+"<li>�Ƿ������Ӳ�����"
	elseif not isInteger(request("RootID")) then
		foundErr=true
		Errmsg=Errmsg+"<br>"+"<li>�Ƿ������Ӳ�����"
	else
		rootID=request("RootID")
	end if

   	UserName=Checkstr(trim(request.Form("username")))
	if memberword=trim(request.Form("passwd")) then
	UserPassWord=Checkstr(trim(request.Form("passwd")))
	else
   	UserPassWord=md5(Checkstr(trim(request.Form("passwd"))))
	end if
	IP=Request.ServerVariables("REMOTE_ADDR") 
	Expression=Checkstr(Request.Form("Expression")&".gif")
   	PostTopic=Checkstr(trim(request.Form("subject")))
   	Body=Checkstr(trim(request.Form("Content")))
	signflag=Checkstr(trim(request.Form("signflag")))
	mailflag=Checkstr(trim(request.Form("emailflag")))
	TotalUseTable=Checkstr(request.Form("TotalUseTable"))

	For i=0 to ubound(AllPostTable)
		if AllPostTable(i)=TotalUseTable then
			FoundTable=true
			Exit For
		end if
	Next
	if Not FoundTable then
   		ErrMsg=ErrMsg+"<Br>"+"<li>��ָ���˷Ƿ������ݱ�������ȷ�����Ǵ���Ч�ı��ύ��"
   		FoundErr=True
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
   		ErrMsg=ErrMsg+"<Br>"+"<li>����������"
   		foundErr=True
	end if
	if strLength(PostTopic)>50 then
   		foundErr=True
      	ErrMsg=ErrMsg+"<Br>"+"<li>���ⳤ�Ȳ��ܳ���50"
	end if
	if request("method")="Topic" then
		if PostTopic="" then
			if body="" then
   				ErrMsg=ErrMsg+"<Br>"+"<li>��������ݱ�����д��һ��"
   				foundErr=True
			end if
		end if
	end if
	if request("method")="fastreply" then
		if body="" then
   		ErrMsg=ErrMsg+"<Br>"+"<li>���ٻظ�����д�������ݡ�"
   		foundErr=True
		end if
		Expression="face0.gif"
	end if
	if strLength(body)>Clng(Forum_Setting(13)) then
   		ErrMsg=ErrMsg+"<Br>"+"<li>�������ݲ��ô���" & CSTR(Forum_Setting(13)) & "bytes"
   		foundErr=true
	end if
	if body="" then
		ErrMsg=ErrMsg+"<Br>"+"<li>û����д���ݡ�"
		foundErr=true
	end if
	session("lastpost")=Now()
	if founderr then
		stats="������Ϣ"
		call nav()
		call head_var(1)
		call dvbbs_error()
	else
		stats=boardtype & "�ظ����ӳɹ�"
		call nav()
		call head_var(2)
		call main()
		if founderr then call dvbbs_error()
	end if
	call footer()

sub main()
	if Cint(GroupSetting(5))=0 then
		Errmsg=Errmsg+"<br>"+"<li>��û���ڱ���̳�ظ����������ӵ�Ȩ�ޣ���<a href=login.asp>��½</a>����ͬ����Ա��ϵ��"
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

	sql="select locktopic,LastPost,title from topic where topicid="&cstr(rootid)
	set rs=conn.execute(sql)
	if not rs.eof and not rs.bof then
		toptopic=rs(2)
		if not master and rs("locktopic")=1 then
			Errmsg=ErrMsg+"<Br>"+"<li>�������Ѿ����������ܷ���ظ���"
			founderr=true
			exit sub
		elseif rs("locktopic")=2 then
			Errmsg=ErrMsg+"<Br>"+"<li>�������Ѿ�ɾ�������ܷ���ظ���"
			founderr=true
			exit sub
		end if
		if not isnull(rs(1)) then
			LastPost=split(rs(1),"$")
			if ubound(LastPost)=6 then
			UpLoadPic_n=LastPost(4)
			else
			UpLoadPic_n=""
			end if
		end if
	end if
	set rs=nothing

	usercookies=request.Cookies("aspsky")("usercookies")
	if isnull(usercookies) or usercookies="" then usercookies=3

	if chkuserlogin(username,userpassword,usercookies,3)=false then
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

	set rsLayer=conn.execute("select b.layer,b.orders,b.EmailFlag,b.username,u.userEmail from "&TotalUseTable&" b inner join [user] u on b.postuserid=u.userid where b.Announceid="&ParentID)
	if not(rsLayer.eof and rsLayer.bof) then
		if isnull(rsLayer(0)) then
			iLayer=1
		else
			iLayer=rslayer(0)
		end if
		if isNUll(rslayer(1)) then
			iOrders=0
		else
			iOrders=rsLayer(1) 
		end if
		if rsLayer(3)=membername then
			if Cint(GroupSetting(4))=0 then
			Errmsg=Errmsg+"<br>"+"<li>��û���ڱ���̳�ظ��Լ������Ȩ�ޡ�"
			founderr=true
			exit sub
			end if
		end if
		if rsLayer(2)=1 then
			topic="����"&Forum_info(0)&"������������˻ظ���"
			email=rsLayer(4)
			mailbody=mailbody & ""&rsLayer(3)&"�����ã�<br>"
			mailbody=mailbody & "����"&Forum_info(0)&"������������˻ظ���<br>"
			mailbody=mailbody & "�뵽���µ�ַ������������ݡ�<br>"
			mailbody=mailbody & "<a href=http://"&request.servervariables("server_name")&replace(lcase(request.servervariables("script_name")),"savereannounce.asp","")&"dispbbs.asp?boardid="&boardid&"&id="&rootid&"&star="&request("star")&"#"&parentid&" target=_blank>�鿴��������</a>"
			if Forum_Setting(2)=0 then
                               
			elseif Forum_Setting(2)=1 then
				call jmail(email,topic,mailbody)
			elseif Forum_Setting(2)=2 then
				call Cdonts(email,topic,mailbody)
			elseif Forum_Setting(2)=3 then
				call aspemail(email,topic,mailbody)
			end if
		end if
	else
		iLayer=1
		iOrders=0
	end if
	rsLayer.close
	set rsLayer=nothing
	if rootid<>0 then 
		iLayer=ilayer+1
		conn.execute "update "&TotalUseTable&" set orders=orders+1 where rootid="&cstr(RootID)&" and orders>"&cstr(iOrders)
		iOrders=iOrders+1
	end if

	DateTimeStr=replace(replace(CSTR(NOW()+Forum_Setting(0)/24),"����",""),"����","")
	'����ظ���
	Sql="insert into "&TotalUseTable&"(Boardid,ParentID,username,topic,body,DateAndTime,length,rootid,layer,orders,ip,Expression,locktopic,signflag,emailflag,isbest,PostUserID) values ("&boardid&","&ParentID&",'"&username&"','"&PostTopic&"','"&body&"','"&DateTimeStr&"','"&strlength(body)&"',"&RootID&","&ilayer&","&iorders&",'"&ip&"','"&Expression&"',0,"&signflag&","&mailflag&",0,"&userid&")"
	conn.execute(sql)
	set rs=conn.execute("select top 1 announceid from "&TotalUseTable&" order by announceid desc")
	announceid=rs(0)
	rs.close
	set rs=nothing
	if PostTopic="" then
		PostTopic=replace(cutStr(body,14),chr(10),"")
	else
		PostTopic=replace(cutStr(PostTopic,14),chr(10),"")
	end if
	LastPost=replace(username,"$","") & "$" & Announceid & "$" & DateTimeStr & "$" & replace(PostTopic,"$","") & "$" & UpLoadPic_n & "$" & UserID & "$" & rootid
	'����������¼
	sql="update topic set child=child+1,LastPostTime=Now(),LastPost='"&LastPost&"' where TopicID="&cstr(rootID)
	conn.execute(sql)

	toptopic=replace(replace(cutStr(toptopic,14),chr(10),""),"'","")
	LastPost_1=replace(username,"$","") & "$" & Announceid & "$" & DateTimeStr & "$" & replace(topTopic,"$","") & "$" & UpLoadPic_n & "$" & UserID & "$" & rootid
	if datediff("d",LastPostTime,Now())=0 then
		sql="update board set lastbbsnum=lastbbsnum+1,todaynum=todaynum+1,LastPost='"&LastPost_1&"' where boardid="&cstr(boardID)
	else
		sql="update board set lastbbsnum=lastbbsnum+1,todaynum=1,LastPost='"&LastPost_1&"' where boardid="&cstr(boardID)
	end if
	conn.execute(sql)

	Dim updateinfo
	set rs=conn.execute("select LastPost,TodayNum,MaxPostNum from config where active=1")
	LastPostTimes=split(rs(0),"$")
	LastPostTime=LastPostTimes(2)
	if not isdate(LastPostTime) then LastPostTime=Now()
	if datediff("d",LastPostTime,Now())=0 then
		if rs(1)+1>rs(2) then updateinfo=",MaxPostNum=todaynum+1"
		sql="update config set bbsnum=bbsnum+1,todayNum=todayNum+1,LastPost='"&LastPost_1&"' "&updateinfo&" where active=1"
	else
		sql="update config set bbsnum=bbsnum+1,yesterdaynum="&rs(1)&",todayNum=1,LastPost='"&LastPost_1&"' where active=1"
	end if
	conn.execute(sql)
		
	rem �����û��Ļظ����ӣ����Ƿ����
	call haveRe()
	dim PostRetrunName
	select case Forum_Setting(69)
	case 1
	response.write "<meta http-equiv=refresh content=""3;URL=index.asp"">"
	PostRetrunName="��ҳ"
	case 2
	response.write "<meta http-equiv=refresh content=""3;URL=list.asp?boardid="&boardid&""">"
	PostRetrunName="������������̳"
	case 3
	response.write "<meta http-equiv=refresh content=""3;URL=dispbbs.asp?boardid="&boardid&"&id="&rootid&"&star="&request("star")&"#"&Announceid&""">"
	PostRetrunName="�������������"
	end select
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

sub haveRe()
	dim userid1,rs1,child
	sql="select postuserid,child from topic where topicID="&rootID
	rs1=conn.execute (sql)
	userid1=rs1(0)
	child=rs1(1)
	set rs1=nothing
	
	if userid<>userid1 and child<2 then
		sql="update [user] set reAnn='"&boardID&"|"& rootID &"' where userid="&userid1
		conn.execute sql
	end if
end sub
%>