<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/char_board.asp" -->
<!-- #include file="inc/char_login.asp" -->
<!-- #include file="inc/chkinput.asp" -->
<% 
dim announceid
dim UserName
dim userPassword
dim useremail
dim Topic
dim body
dim ip
dim Expression
dim rootid
dim signflag
dim mailflag
dim char_changed
dim addbody
dim totalusetable
Dim FoundTable
FoundTable=false
addbody=request("Content")
dim re
Set re=new RegExp
re.IgnoreCase =true
re.Global=True

re.Pattern="\[align=right\]\[color=#000066\](.*)\[\/color\]\[\/align\]"
addbody=re.Replace(addbody,"")
set re=Nothing

if membername<>request("username") then 
	char_changed = "[align=right][color=#000066][�������Ѿ���"&membername&"��"&Now()&"�༭��][/color][/align]"
else
	char_changed = "[align=right][color=#000066][�������Ѿ���������"&Now()&"�༭��][/color][/align]"
end if

UserName=trim(checkStr(request("username")))
UserPassWord=trim(request("passwd"))
IP=Request.ServerVariables("REMOTE_ADDR") 
Expression=Request.Form("Expression")&".gif"
BoardID=Request("boardID")
rootID=Cstr(Request("ID"))
AnnounceID=request("replyID")
Topic=checkStr(trim(request("subject")))
Body=checkstr(addbody)+chr(13)+chr(10)+char_changed+chr(13)
signflag=trim(request("signflag"))
mailflag=trim(request("Forum_Setting(2)"))
foundErr=false
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
if not founduser then
	ErrMsg=ErrMsg+"<Br>"+"<li>���½������޸ġ�"
	foundErr=True
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
if instr(Expression,"face")=0 then
	Expression="face1.gif"
end if
if chkpost=false then
	ErrMsg=ErrMsg+"<Br>"+"<li>���ύ�����ݲ��Ϸ����벻Ҫ���ⲿ�ύ���ԡ�"
   	FoundErr=True
end if
if AnnounceID="" then
	founderr=true
	Errmsg=Errmsg+"<br>"+"<li>��ָ��������ӡ�"
elseif not isInteger(AnnounceID) then
	founderr=true
	Errmsg=Errmsg+"<br>"+"<li>�Ƿ������Ӳ�����"
end if
if BoardID="" then
	founderr=true
	Errmsg=Errmsg+"<br>"+"<li>��ָ��������ӡ�"
elseif not isInteger(BoardID) then
	founderr=true
	Errmsg=Errmsg+"<br>"+"<li>�Ƿ������Ӳ�����"
end if
if rootID="" then
	founderr=true
	Errmsg=Errmsg+"<br>"+"<li>��ָ��������ӡ�"
elseif not isInteger(rootID) then
	founderr=true
	Errmsg=Errmsg+"<br>"+"<li>�Ƿ������Ӳ�����"
end if
if UserName="" then
	ErrMsg=ErrMsg+"<Br>"+"<li>����������(���Ȳ��ܴ���20)"
   	foundErr=True
elseif Trim(UserPassWord)="" then
	ErrMsg=ErrMsg+"<Br>"+"<li>����������(���Ȳ��ܴ���16)"
   	foundErr=True
end if
if rootid=announceid then
	if Topic="" then
   		foundErr=True
		ErrMsg=ErrMsg+"<Br>"+"<li>���ⲻӦΪ��"
	elseif strLength(topic)>50 then
   		foundErr=True
      	ErrMsg=ErrMsg+"<Br>"+"<li>���ⳤ�Ȳ��ܳ���50"
	end if
end if
if request("content")="" then
	ErrMsg=ErrMsg+"<Br>"+"<li>���ݲ���Ϊ��"
   	foundErr=True
end if
if strLength(body)>Forum_Setting(13) then
	ErrMsg=ErrMsg+"<Br>"+"<li>�������ݲ��ô���" & CSTR(Forum_Setting(13)) & "bytes"
   	foundErr=true
end if
stats=boardtype & "�༭���ӳɹ�"

if founderr then
	call nav()
	call head_var(1)
	call dvbbs_error()
else
	call nav()
	call head_var(2)
	call main()
end if
call activeonline()
call footer()

sub main()
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
dim LastBoard
dim LastTopic
dim LastPost
dim caneditpost
caneditpost=false
sql="select b.username,u.usergroupID from "&TotalUseTable&" b,[user] u where b.postuserid=u.userid and b.AnnounceID="&AnnounceID
set rs=conn.execute(sql)
if rs.eof and rs.bof then
	Errmsg=Errmsg+"<br>"+"<li>û���ҵ���Ӧ�����ӡ�"
	Founderr=true
	exit sub
else
	if rs("username")=membername then
		if Cint(GroupSetting(10))=0 then
			Errmsg=Errmsg+"<br>"+"<li>��û���ڱ���̳�༭�Լ����ӵ�Ȩ�ޣ���<a href=login.asp>��½</a>����ͬ����Ա��ϵ��"
			founderr=true
			CanEditPost=False
		else
			CanEditPost=True
		end if
	else
		if (master or superboardmaster or boardmaster) and Cint(GroupSetting(23))=1 then
			CanEditPost=True
		else
			CanEditPost=False
		end if
		if UserGroupID>3 and Cint(GroupSetting(23))=1 then
			CanEditPost=true
		end if
		if Cint(GroupSetting(23))=1 and FoundUserPer then
			CanEditPost=True
		elseif Cint(GroupSetting(23))=0 and FoundUserPer then
			CanEditPost=False
		end if
		if UserGroupID<4 and UserGroupID=rs("UserGroupID") then
			Errmsg=Errmsg+"<br>"+"<li>ͬ�ȼ��û������޸ġ�"
			Founderr=true
			exit sub
		elseif UserGroupID<4 and UserGroupID>rs("UserGroupID") then
			Errmsg=Errmsg+"<br>"+"<li>�����޸ĵȼ������ߵ��û������ӡ�"
			Founderr=true
			exit sub
		end if
		if not CanEditPost then
			Errmsg=Errmsg+"<br>"+"<li>��û���㹻��Ȩ�ޱ༭�����ӣ���͹���Ա��ϵ��"
			Founderr=true
			exit sub
		end if
	end if
end if
'ȡ����ǰ�������ظ�id,�������Ϊ���ظ��������Ӧ����
sql="select LastPost from board where boardid="&boardid
set rs=conn.execute(sql)
if not (rs.eof and rs.bof) then
	if not isnull(rs(0)) and rs(0)<>"" then
		LastBoard=split(rs(0),"$")
		if ubound(LastBoard)=6 then
			if Clng(LastBoard(6))=Clng(AnnounceID) then
			LastPost=LastBoard(0) & "$" & LastBoard(1) & "$" & Now() & "$" & replace(cutStr(Topic,20),"$","") & "$" & LastBoard(4) & "$" & LastBoard(5) & "$" & LastBoard(6)
			conn.execute("update board set LastPost='"&LastPost&"' where boardid="&boardid)
			end if
		end if
	end if
end if

'ȡ�õ�ǰ�������ظ�id,�������Ϊ���ظ��������Ӧ����
sql="select LastPost from topic where topicid="&rootid
set rs=conn.execute(sql)
if not (rs.eof and rs.bof) then
	if not isnull(rs(0)) and rs(0)<>"" then
		LastTopic=split(rs(0),"$")
		if ubound(LastTopic)=6 then
			if Clng(LastTopic(1))=Clng(Announceid) then
			LastPost=LastTopic(0) & "$" & LastTopic(1) & "$" & Now() & "$" & replace(cutStr(body,20),"$","") & "$" & LastTopic(4) & "$" & LastTopic(5) & "$" & LastTopic(6)
			conn.execute("update topic set LastPost='"&LastPost&"' where topicid="&rootid)
			end if
		end if
	end if
end if

Set rs = Server.CreateObject("ADODB.Recordset")
sql="SELECT * FROM "&TotalUseTable&" where username='"&trim(username)&"' and AnnounceID="&Announceid
rs.Open sql,conn,1,3
if rs.eof and rs.bof then
	foundErr=True
	ErrMsg=ErrMsg+"<Br>"+"<li>�����Ǳ����ӵ����ߣ���Ȩ�޸ģ�"
	rs.close:set rs=nothing
	exit sub
elseif not master and rs("locktopic")=1 then
	Errmsg=ErrMsg+"<Br>"+"<li>�������Ѿ����������ܱ༭��"
	foundErr=true
	rs.close:set rs=nothing
	exit sub
else
	if rs("parentid")=0 then
		conn.execute("update topic set title='"&topic&"',LastPostTime=Now() where topicid="&rootid)
	end if
	rs("Topic") =Topic
	rs("Body") =Body
	rs("length")=strlength(body)
	rs("ip")=ip
	rs("Expression")=Expression
	rs("signflag")=signflag
	rs("emailflag")=mailflag
	rs.Update
	dim PostRetrunName
	select case Forum_Setting(69)
	case 1
		response.write "<meta http-equiv=refresh content=""3;URL=index.asp"">"
		PostRetrunName="��ҳ"
	case 2
		response.write "<meta http-equiv=refresh content=""3;URL=list.asp?boardid="&boardid&""">"
		PostRetrunName="������������̳"
	case 3
		response.write "<meta http-equiv=refresh content=""3;URL=dispbbs.asp?boardid="&boardid&"&id="&rootid&"&star="&request("star")&"#"&rootid&""">"
		PostRetrunName="�����޸ĵ�����"
	end select
	rs.close:set rs=nothing
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
end if
end sub
%>