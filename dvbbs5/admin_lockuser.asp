<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<%
dim username
dim locktype
dim ip
Dim TotalUseTable
Dim AdminUserPer
AdminUserPer=false
if (master or boardmaster or superboardmaster) and Cint(GroupSetting(42))=1 then
	AdminUserPer=true
else
	AdminUserPer=false
end if
if UserGroupID>3 and Cint(GroupSetting(42))=1 then
	AdminUserPer=true
end if
if FoundUserPer and Cint(GroupSetting(42))=1 then
	AdminUserPer=true
elseif FoundUserPer and Cint(GroupSetting(42))=0 then
	AdminUserPer=false
end if

ip=Request.ServerVariables("REMOTE_ADDR")
stats="�����û�"
if request("name")="" then
	Errmsg=Errmsg+"<br>"+"<li>��ָ�����������û���"
	Founderr=true
else
	username=CheckStr(request("name"))
end if

call nav()
call head_var(1)
if founderr then
	call dvbbs_error()
else
	if request("action")="power" then
		call Poweruser()
	elseif request("action")="DelTopic" then
		call DelTopic()
	elseif request("action")="getpermission" then
		call boardlist()
	elseif request("action")="userBoardPermission" then
		call GetUserPermission()
	elseif request("action")="saveuserpermission" then
		call saveuserpermission()
	else
		call lockuser()
	end if
	if founderr then call dvbbs_error()
end if
call footer()

sub lockuser()
dim canlockuser
canlockuser=false
if (master or boardmaster or superboardmaster) and Cint(GroupSetting(28))=1 then
	canlockuser=true
else
	canlockuser=false
end if
if UserGroupID>3 and Cint(GroupSetting(28))=1 then
	canlockuser=true
end if
if FoundUserPer and Cint(GroupSetting(28))=1 then
	canlockuser=true
elseif FoundUserPer and Cint(GroupSetting(28))=0 then
	canlockuser=false
end if

if not canlockuser then
	Errmsg=Errmsg+"<br>"+"<li>��û��Ȩ��ִ�д˲�����"
	founderr=true
	exit sub
end if
if request("action")="lock_1" then
	conn.execute("update [user] set LockUser=1 where username='"&username&"' and UserGroupID>1")
	locktype="����"
elseif request("action")="lock_2" then
	conn.execute("update [user] set LockUser=2 where username='"&username&"' and UserGroupID>1")
	locktype="����"
elseif request("action")="lock_3" then
	conn.execute("update [user] set LockUser=0 where username='"&username&"' and UserGroupID>1")
	locktype="����"
else
	Errmsg=Errmsg+"<br>"+"<li>��ָ����ȷ�Ĳ�����"
	Founderr=true
	exit sub
end if

sucmsg="<li>��ѡ����û��Ѿ�"&locktype&"��"
call dvbbs_suc()
end sub

sub Poweruser()
dim title,content
dim canlockuser
canlockuser=false
if (master or boardmaster or superboardmaster) and Cint(GroupSetting(43))=1 then
	canlockuser=true
else
	canlockuser=false
end if
if UserGroupID>3 and Cint(GroupSetting(43))=1 then
	canlockuser=true
end if
if FoundUserPer and Cint(GroupSetting(43))=1 then
	canlockuser=true
elseif FoundUserPer and Cint(GroupSetting(43))=0 then
	canlockuser=false
end if

if not canlockuser then
	Errmsg=Errmsg+"<br>"+"<li>��û��Ȩ��ִ�д˲�����"
	founderr=true
	exit sub
end if
if request("checked")="yes" then
	dim doWealth,douserEP,douserCP,douserPower
	dim doWealthMsg,douserEPMsg,douserCPMsg,douserPowerMsg,allMsg
	if not isnumeric(request("doWealth")) or request("doWealth")="0" or request("doWealth")="" then
		doWealth=0
		doWealthMsg=""
	else
		doWealth=request("doWealth")
		doWealthMsg="��Ǯ" & request("doWealth") & "��"
	end if

	if not isnumeric(request("douserEP")) or request("douserEP")="0" or request("douserEP")="" then
		douserEP=0
		douserEPMsg=""
	else
		douserEP=request("douserEP")
		douserEPMsg="����" & request("douserEP") & "��"
	end if

	if not isnumeric(request("douserCP")) or request("douserCP")="0" or request("douserCP")="" then
		douserCP=0
		douserCPMsg=""
	else
		douserCP=request("douserCP")
		douserCPMsg="����" & request("douserCP") & "��"
	end if

	if not isnumeric(request("douserPower")) or request("douserPower")="0" or request("douserPower")="" then
		douserPower=0
		douserPowerMsg=""
	else
		douserPower=request("douserPower")
		douserPowerMsg="����" & request("douserPower")
	end if

	if doWealthMsg="" and douserEPMsg="" and douserCPMsg="" and douserPowerMsg="" then
		allmsg="û�ж��û����з�ֵ����"
	else
		allmsg="�û�������" & doWealthMsg & douserEPMsg & douserCPMsg & douserPowerMsg
	end if
	'response.write allmsg
	'response.end

	title=request.form("title")
	content=request.form("content")
	content="ԭ��" & title & content
	if request.form("title")="" and request.form("content")="" then
		Errmsg=Errmsg+"<br><li>��д������ԭ��"
		founderr=true
		exit sub
	end if

	sql="insert into log (l_touser,l_username,l_content,l_ip) values ('"&username&"','"&membername&"','�û�������"&content& "��"&allmsg&"','"&ip&"')"
	conn.execute(sql)
	if allmsg<>"" then
		conn.execute("update [user] set userWealth=userWealth+"&doWealth&",userCP=userCP+"&douserCP&",userEP=userEP+"&douserEP&",userPower=userPower+"&douserPower&" where username='"&username&"'")
	end if
	locktype="�û�����"
	sucmsg="<li>��ѡ����û��Ѿ�"&locktype&"��"
	call dvbbs_suc()
else
%>
<FORM METHOD=POST ACTION="admin_lockuser.asp?action=power">
<table style="width:70%" cellspacing="1" cellpadding="3" align="center" class=tableborder1>
  <tr> 
    <th height=24>��̳�������ģ�����Ҫ���еĲ����ǽ����û�</th>
  </tr>   
  <tr> 
    <td class=tablebody1 height=24><b>
      ��������</b>��  
	  <select name="title" size=1>
<option value="">�Զ���</option>
<option value="��η��������">��η��������</option>
<option value="�����������й���">�����������й���</option>
<option value="��η����ˮ����">��η����ˮ����</option>
<option value="��η���������">��η���������</option>
	  </select>
	  <input type="text" name="content" size=50>  *</td>
  </tr>   
  <tr> 
    <td class=tablebody1 height=24><b>
      �û�����</b>��  ��Ǯ
	<select name="doWealth" size=1>

<%for i=-50 to 50%>
<option value="<%=i%>" <%if cint(i)=cint(0) then%>selected<%end if%>><%=i%></option>
<%next%>
	</select>&nbsp;����
	<select name="douserCP" size=1>

<%for i=-50 to 50%>
<option value="<%=i%>" <%if cint(i)=cint(0) then%>selected<%end if%>><%=i%></option>
<%next%>
	</select>&nbsp;����
	<select name="douserEP" size=1>

<%for i=-50 to 50%>
<option value="<%=i%>" <%if cint(i)=cint(0) then%>selected<%end if%>><%=i%></option>
<%next%>
	</select>&nbsp;����
	<select name="douserPower" size=1>

<%for i=-5 to 5%>
<option value="<%=i%>" <%if cint(i)=cint(0) then%>selected<%end if%>><%=i%></option>
<%next%>
	</select>
  *</td>
  </tr> 
<input type=hidden value="yes" name="checked">
<input type=hidden value="<%=username%>" name="name">
  <tr> 
    <td class=tablebody2 height=24>
      ������ʹ�ù���Ա�Ĺ���ְ�ܣ�����Ա���в���������¼&nbsp;<input type="submit" name=submit value="ȷ�ϲ���"></td>
  </tr>   
</table>
</FORM>
<%
end if
end sub

sub deltopic()
Dim DelDate,suserid
Dim todaynum
DelDate=checkstr(request.form("delTopicDate"))
suserid=checkstr(request.form("SetUserID"))
dim canlockuser
canlockuser=false
if (master or boardmaster or superboardmaster) and Cint(GroupSetting(29))=1 then
	canlockuser=true
else
	canlockuser=false
end if
if UserGroupID>3 and Cint(GroupSetting(29))=1 then
	canlockuser=true
end if
if FoundUserPer and Cint(GroupSetting(29))=1 then
	canlockuser=true
elseif FoundUserPer and Cint(GroupSetting(29))=0 then
	canlockuser=false
end if

if not canlockuser then
	Errmsg=Errmsg+"<br>"+"<li>��û��Ȩ��ִ�д˲�����"
	founderr=true
	exit sub
end if
if DelDate="" or not isnumeric(DelDate) then
	Errmsg=Errmsg+"<br>"+"<li>�Ƿ��Ĳ�����"
	Founderr=true
	exit sub
end if
if suserid="" or not isnumeric(suserid) then
	Errmsg=Errmsg+"<br>"+"<li>�Ƿ��Ĳ�����"
	Founderr=true
	exit sub
end if
'ɾ�����ӱ�������
'�����û����������û�ɾ��������Ϊ�������лظ������ظ��˵�ɾ��������������
i=0
Set rs=conn.execute("select TopicID,Child,BoardID from topic where PostUserID="&suserid&" and datediff('d',DateAndTime,Now())<="&DelDate&"")
do while not rs.eof
	i=i+1
	if clng(rs(0))>clng(MaxOldID) then
		TotalUseTable=NowUseBbs
	else
		TotalUseTable="bbs1"
	end if
	'���·����û�����
	conn.execute("update [user] set article=article-1,UserDel=UserDel-"&rs(1)&" where userid="&suserid)
	'��������״̬
	conn.execute("update topic set locktopic=2 where topicid="&rs(0))
	'��������״̬
	conn.execute("update "&TotalUseTable&" set locktopic=2 where rootid="&rs(0))
	if Cint(DelDate)=1 then
		todaynum=rs(1)
	else
		todaynum=0
	end if
	'���°�������
	conn.execute("update board set lastbbsnum=lastbbsnum-"&rs(1)&",lasttopicNum=lasttopicNum-1,todayNum=todayNum-"&todayNum&" where boardid="&rs(2))
	'����������
	conn.execute("update config set TodayNum=todayNum-"&todaynum&",BbsNum=bbsNum-"&rs(1)&",TopicNum=topicNum-1")
	'�������ظ�����
	LastCount(rs(2))
rs.movenext
loop
set rs=nothing
sucmsg="<li>ɾ�����û�����"&i&"����<li>���û�ɾ��������Ϊ��ɾ�����лظ������ظ��˵�ɾ��������������"
call dvbbs_suc()
end sub

sub boardlist()
dim trs
dim dispboard
if not AdminUserPer then
	Errmsg=Errmsg+"<br>"+"<li>��û��Ȩ��ִ�д˲�����"
	founderr=true
	exit sub
else
	dispboard=true
end if
response.write "<table style=""width:70%"" cellspacing=1 cellpadding=3 align=center class=tableborder1>"
response.write "<tr><th colspan=6 height=23 align=left>�༭"&request("username")&"��̳Ȩ�ޣ���ɫ��ʾ���û��ڸð������Զ���Ȩ�ޣ�</th></tr>"
if not isnumeric(request("userid")) then
	response.write "<tr><td colspan=6 class=tablebody1>������û�������</td></tr>"
	founderr=true
end if
if not founderr then
response.write "<tr><td colspan=6 class=tablebody1 height=25><a href=?action=userBoardPermission&boardid=0&userid="&request("userid")&">�༭���û�������ҳ���Ȩ��</a>����Ҫ��Զ��Ų������ã�</td></tr>"
'----------------------boardinfo--------------------
response.write "<tr><td colspan=6 class=tablebody1>"
dim rs_1,rs_2
dim sql_1,sql_2
	sql_2 = "select * from class order by orders"
	set rs_2=conn.execute(sql_2)
	do while not rs_2.Eof
%>
<table width="95%" border="0" cellspacing="3" cellpadding="0" align=center>
<tr><td height="21"><B><%=rs_2("class")%></b></td></tr>
</table>
<%
		sql_1 = "select boardid,boardtype,boardmaster from board where class="&rs_2("id")&" order by orders"
		set rs_1=conn.execute(sql_1)
		do while not rs_1.EOF
		if boardmaster and instr("|"&rs_1(2)&"|","|"&membername&"|")=0 then
			dispboard=false
		end if
		if master or superboardmaster then dispboard=true
		if dispboard then
%>
<table width="95%" border="0" cellspacing="3" cellpadding="0" align=center>
<tr> 
<td height="18" width=50% ><li><%=rs_1("boardtype")%></td>
<td width=50% >
<a href="?action=userBoardPermission&boardid=<%=rs_1("boardid")%>&userid=<%=request("userid")%>&name=<%=request("name")%>">
<%
set trs=conn.execute("select uc_userid from UserAccess where uc_Boardid="&rs_1("boardid")&" and uc_userid="&request("userid"))
if not (trs.eof and trs.bof) then
	response.write "<font color=red>"
end if
%>
�༭�û��ڸ���̳Ȩ��</a></td>
</tr>
</table>
<%
		end if
		rs_1.MoveNext
		loop
%>
<table width="95%" border="0" cellspacing="3" cellpadding="0" align=center>
<tr>
<td height="21" align=left><hr color=black height=1 width=70% size=1></td>
</tr>
</table>
<%
	rs_2.MoveNext 
	Loop
	set trs=nothing
set rs_1=nothing
set rs_2=nothing
response.write "</td></tr></table>"
'-------------------end-------------------
	end if
end sub

sub GetUserPermission()
if not AdminUserPer then
	Errmsg=Errmsg+"<br>"+"<li>��û��Ȩ��ִ�д˲�����"
	founderr=true
	exit sub
end if
dim usertitle
if not isnumeric(request("userid")) then
	errmsg=errmsg+"<br><li>������û�������"
	founderr=true
	exit sub
end if
if not isnumeric(request("boardid")) then
	errmsg=errmsg+"<br><li>����İ��������"
	founderr=true
	exit sub
end if
if not founderr then
	response.write "<table class=tableborder1 cellspacing=1 cellpadding=3  align=center>"
	set rs=conn.execute("select u.UserGroupID,ug.title,u.username from [user] u inner join UserGroups UG on u.userGroupID=ug.userGroupID where u.userid="&request("userid"))
	UserGroupID=rs(0)
	usertitle=rs(1)
	membername=rs(2)
	set rs=conn.execute("select boardtype from board where boardid="&request("boardid"))
	if rs.eof and rs.bof then
	boardtype="��̳����ҳ��"
	else
	boardtype=rs(0)
	end if
	response.write "<tr><th colspan=6 height=23 align=left>�༭ "&membername&" �� "&boardtype&" Ȩ��</th></tr>"
	response.write "<tr><td colspan=6 height=25 class=tablebody1>ע�⣺���û����� <B>"&usertitle&"</B> �û����У�����������������Զ���Ȩ�ޣ�����û�Ȩ�޽����Զ���Ȩ��Ϊ��</td></tr>"

Dim reGroupSetting
Dim FoundGroup,FoundUserPermission,FoundGroupPermission
FoundGroup=false
FoundUserPermission=false
FoundGroupPermission=false

set rs=conn.execute("select * from UserAccess where uc_boardid="&request("boardid")&" and uc_userid="&request("userid"))
if not (rs.eof and rs.bof) then
	reGroupSetting=split(rs("uc_Setting"),",")
	FoundGroup=true
	FoundUserPermission=true
end if

if not foundgroup then
set rs=conn.execute("select * from BoardPermission where boardid="&request("boardid")&" and groupid="&usergroupid)
if not(rs.eof and rs.bof) then
	reGroupSetting=split(rs("PSetting"),",")
	FoundGroup=true
	FoundGroupPermission=true
end if
end if

if not foundgroup then
set rs=conn.execute("select * from usergroups where usergroupid="&usergroupid)
if rs.eof and rs.bof then
	response.write "δ�ҵ����û��飡"
	response.end
else
	FoundGroup=true
	FoundGroupPermission=true
	reGroupSetting=split(rs("GroupSetting"),",")
end if
end if
%>
<FORM METHOD=POST ACTION="?action=saveuserpermission">
<input type=hidden name="userid" value="<%=request("userid")%>">
<input type=hidden name="BoardID" value="<%=request("boardid")%>">
<input type=hidden name="username" value="<%=membername%>">
<input type=hidden name="name" value="<%=membername%>">
<tr> 
<td height="23" colspan="2" class=tablebody1><input type=radio name="isdefault" value="1" <%if FoundGroupPermission then%>checked<%end if%>><B>ʹ���û���Ĭ��ֵ</B> (ע��: �⽫ɾ���κ�֮ǰ�������Զ�������)</td>
</tr>
<tr> 
<td height="23" colspan="2"  class=tablebody1><input type=radio name="isdefault" value="0" <%if FoundUserPermission then%>checked<%end if%>><B>ʹ���Զ�������</B> </td>
</tr>
<tr> 
<th height="23" colspan="2" align=left>�����鿴Ȩ��</th>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>���������̳</td>
<td height="23" width="40%" class=tablebody1>��<input name="canview" type=radio value="1" <%if reGroupSetting(0)=1 then%>checked<%end if%>>&nbsp;��<input name="canview" type=radio value="0" <%if reGroupSetting(0)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>���Բ鿴��Ա��Ϣ(����������Ա�����Ϻͻ�Ա�б�)
</td>
<td height="23" width="40%" class=tablebody1>��<input name="canviewuserinfo" type=radio value="1" <%if reGroupSetting(1)=1 then%>checked<%end if%>>&nbsp;��<input name="canviewuserinfo" type=radio value="0" <%if reGroupSetting(1)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>���Բ鿴�����˷���������
</td>
<td height="23" width="40%" class=tablebody1>��<input name="canviewpost" type=radio value="1" <%if reGroupSetting(2)=1 then%>checked<%end if%>>&nbsp;��<input name="canviewpost" type=radio value="0" <%if reGroupSetting(2)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>���������������
</td>
<td height="23" width="40%" class=tablebody1>��<input name="canviewbest" type=radio value="1" <%if reGroupSetting(41)=1 then%>checked<%end if%>>&nbsp;��<input name="canviewbest" type=radio value="0" <%if reGroupSetting(41)=0 then%>checked<%end if%>></td>
</tr>
<tr> 
<th height="23" colspan="2" align=left>����<b>����Ȩ��</b></th>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>���Է���������</td>
<td height="23" width="40%" class=tablebody1>��<input name="cannewpost" type=radio value="1" <%if reGroupSetting(3)=1 then%>checked<%end if%>>&nbsp;��<input name="cannewpost" type=radio value="0" <%if reGroupSetting(3)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>���Իظ��Լ�������
</td>
<td height="23" width="40%" class=tablebody1>��<input name="canreplymytopic" type=radio value="1" <%if reGroupSetting(4)=1 then%>checked<%end if%>>&nbsp;��<input name="canreplymytopic" type=radio value="0" <%if reGroupSetting(4)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>���Իظ������˵�����
</td>
<td height="23" width="40%" class=tablebody1>��<input name="canreplytopic" type=radio value="1" <%if reGroupSetting(5)=1 then%>checked<%end if%>>&nbsp;��<input name="canreplytopic" type=radio value="0" <%if reGroupSetting(5)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>��������̳�������ֵ�ʱ���������(�ʻ��ͼ���)?
</td>
<td height="23" width="40%" class=tablebody1>��<input name="canpostagree" type=radio value="1" <%if reGroupSetting(6)=1 then%>checked<%end if%>>&nbsp;��<input name="canpostagree" type=radio value="0" <%if reGroupSetting(6)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>�������������Ǯ
</td>
<td height="23" width="40%" class=tablebody1><input name="postagreemoney" type=text size=4 value="<%=reGroupSetting(47)%>"></td>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>�����ϴ�����
</td>
<td height="23" width="40%" class=tablebody1>��<input name="canupload" type=radio value="1" <%if reGroupSetting(7)=1 then%>checked<%end if%>>&nbsp;��<input name="canupload" type=radio value="0" <%if reGroupSetting(7)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>����ϴ��ļ�����
</td>
<td height="23" width="40%" class=tablebody1><input name="canuploadnum" type=text size=4 value="<%=reGroupSetting(40)%>"></td>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>�ϴ��ļ���С����
</td>
<td height="23" width="40%" class=tablebody1><input name="MaxUploadSize" type=text size=4 value="<%=reGroupSetting(44)%>"> KB</td>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>���Է�����ͶƱ</td>
<td height="23" width="40%" class=tablebody1>��<input name="canpostvote" type=radio value="1" <%if reGroupSetting(8)=1 then%>checked<%end if%>>&nbsp;��<input name="canpostvote" type=radio value="0" <%if reGroupSetting(8)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>���Բ���ͶƱ</td>
<td height="23" width="40%" class=tablebody1>��<input name="canvote" type=radio value="1" <%if reGroupSetting(9)=1 then%>checked<%end if%>>&nbsp;��<input name="canvote" type=radio value="0" <%if reGroupSetting(9)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>���Է���С�ֱ�</td>
<td height="23" width="40%" class=tablebody1>��<input name="cansmallpaper" type=radio value="1"  <%if reGroupSetting(17)=1 then%>checked<%end if%>>&nbsp;��<input name="cansmallpaper" type=radio value="0" <%if reGroupSetting(17)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>����С�ֱ������Ǯ</td>
<td height="23" width="40%" class=tablebody1><input name="smallpapermoney" type=text value="<%=reGroupSetting(46)%>" size=4></td>
</tr>
<tr> 
<th height="23" colspan="2" align=left>����<b>����/����༭Ȩ��</b></th>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>���Ա༭�Լ�������
</td>
<td height="23" width="40%" class=tablebody1>��<input name="caneditmytopic" type=radio value="1" <%if reGroupSetting(10)=1 then%>checked<%end if%>>&nbsp;��<input name="caneditmytopic" type=radio value="0" <%if reGroupSetting(10)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>����ɾ���Լ�������
</td>
<td height="23" width="40%" class=tablebody1>��<input name="candelmytopic" type=radio value="1" <%if reGroupSetting(11)=1 then%>checked<%end if%>>&nbsp;��<input name="candelmytopic" type=radio value="0" <%if reGroupSetting(11)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>�����ƶ��Լ������ӵ�������̳
</td>
<td height="23" width="40%" class=tablebody1>��<input name="canmovemytopic" type=radio value="1" <%if reGroupSetting(12)=1 then%>checked<%end if%>>&nbsp;��<input name="canmovemytopic" type=radio value="0" <%if reGroupSetting(12)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>���Դ�/�ر��Լ�����������
</td>
<td height="23" width="40%" class=tablebody1>��<input name="canclosemytopic" type=radio value="1" <%if reGroupSetting(13)=1 then%>checked<%end if%>>&nbsp;��<input name="canclosemytopic" type=radio value="0" <%if reGroupSetting(13)=0 then%>checked<%end if%>></td>
</tr>
<tr> 
<th height="23" colspan="2" align=left>����<b>����Ȩ��</b></th>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>����������̳
</td>
<td height="23" width="40%" class=tablebody1>��<input name="cansearch" type=radio value="1" <%if reGroupSetting(14)=1 then%>checked<%end if%>>&nbsp;��<input name="cansearch" type=radio value="0" <%if reGroupSetting(14)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>����ʹ��'���ͱ�ҳ������'����
</td>
<td height="23" width="40%" class=tablebody1>��<input name="canmailtopic" type=radio value="1" <%if reGroupSetting(15)=1 then%>checked<%end if%>>&nbsp;��<input name="canmailtopic" type=radio value="0" <%if reGroupSetting(15)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>�����޸ĸ�������
</td>
<td height="23" width="40%" class=tablebody1>��<input name="canmodify" type=radio value="1" <%if reGroupSetting(16)=1 then%>checked<%end if%>>&nbsp;��<input name="canmodify" type=radio value="0" <%if reGroupSetting(16)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>�����п��Բ鿴������ͷ��
</td>
<td height="23" width="40%" class=tablebody1>��<input name="canusetitle" type=radio value="1" <%if reGroupSetting(36)=1 then%>checked<%end if%>>&nbsp;��<input name="canusetitle" type=radio value="0" <%if reGroupSetting(36)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>�����п��Բ鿴������ͷ��
</td>
<td height="23" width="40%" class=tablebody1>��<input name="canuseface" type=radio value="1"  <%if reGroupSetting(37)=1 then%>checked<%end if%>>&nbsp;��<input name="canuseface" type=radio value="0" <%if reGroupSetting(37)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>�����п��Կ��Բ鿴������ǩ��
</td>
<td height="23" width="40%" class=tablebody1>��<input name="canusesign" type=radio value="1"  <%if reGroupSetting(38)=1 then%>checked<%end if%>>&nbsp;��<input name="canusesign" type=radio value="0" <%if reGroupSetting(38)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>���������̳�¼�
</td>
<td height="23" width="40%" class=tablebody1>��<input name="canvieweven" type=radio value="1"  <%if reGroupSetting(39)=1 then%>checked<%end if%>>&nbsp;��<input name="canvieweven" type=radio value="0" <%if reGroupSetting(39)=0 then%>checked<%end if%>></td>
</tr>
<tr> 
<th height="23" colspan="2" align=left>��������Ȩ��</th>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>����ɾ������������
</td>
<td height="23" width="40%" class=tablebody1>��<input name="candeltopic" type=radio value="1" <%if reGroupSetting(18)=1 then%>checked<%end if%>>&nbsp;��<input name="candeltopic" type=radio value="0"  <%if reGroupSetting(18)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>�����ƶ�����������
</td>
<td height="23" width="40%" class=tablebody1>��<input name="canmovetopic" type=radio value="1" <%if reGroupSetting(19)=1 then%>checked<%end if%>>&nbsp;��<input name="canmovetopic" type=radio value="0"  <%if reGroupSetting(19)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>���Դ�/�ر�����������
</td>
<td height="23" width="40%" class=tablebody1>��<input name="canclosetopic" type=radio value="1" <%if reGroupSetting(20)=1 then%>checked<%end if%>>&nbsp;��<input name="canclosetopic" type=radio value="0"  <%if reGroupSetting(20)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>���Թ̶�/����̶�����
</td>
<td height="23" width="40%" class=tablebody1>��<input name="cantoptopic" type=radio value="1" <%if reGroupSetting(21)=1 then%>checked<%end if%>>&nbsp;��<input name="cantoptopic" type=radio value="0"  <%if reGroupSetting(21)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>���Խ���/�ͷ������û�
</td>
<td height="23" width="40%" class=tablebody1>��<input name="canawardtopic" type=radio value="1" <%if reGroupSetting(22)=1 then%>checked<%end if%>>&nbsp;��<input name="canawardtopic" type=radio value="0"  <%if reGroupSetting(22)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>���Խ���/�ͷ��û�
</td>
<td height="23" width="40%" class=tablebody1>��<input name="canaward" type=radio value="1" <%if reGroupSetting(43)=1 then%>checked<%end if%>>&nbsp;��<input name="canaward" type=radio value="0"  <%if reGroupSetting(43)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>���Ա༭����������
</td>
<td height="23" width="40%" class=tablebody1>��<input name="canmodifytopic" type=radio value="1" <%if reGroupSetting(23)=1 then%>checked<%end if%>>&nbsp;��<input name="canmodifytopic" type=radio value="0" <%if reGroupSetting(23)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>���Լ���/�����������
</td>
<td height="23" width="40%" class=tablebody1>��<input name="canbesttopic" type=radio value="1" <%if reGroupSetting(24)=1 then%>checked<%end if%>>&nbsp;��<input name="canbesttopic" type=radio value="0"  <%if reGroupSetting(24)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>���Է�������
</td>
<td height="23" width="40%" class=tablebody1>��<input name="canAnnounce" type=radio value="1" <%if reGroupSetting(25)=1 then%>checked<%end if%>>&nbsp;��<input name="canAnnounce" type=radio value="0"  <%if reGroupSetting(25)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>���Թ�����
</td>
<td height="23" width="40%" class=tablebody1>��<input name="canAdminAnnounce" type=radio value="1" <%if reGroupSetting(26)=1 then%>checked<%end if%>>&nbsp;��<input name="canAdminAnnounce" type=radio value="0"  <%if reGroupSetting(26)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>���Թ���С�ֱ�
</td>
<td height="23" width="40%" class=tablebody1>��<input name="canAdminPaper" type=radio value="1" <%if reGroupSetting(27)=1 then%>checked<%end if%>>&nbsp;��<input name="canAdminPaper" type=radio value="0"  <%if reGroupSetting(27)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>��������/����/��������û�
</td>
<td height="23" width="40%" class=tablebody1>��<input name="canAdminUser" type=radio value="1" <%if reGroupSetting(28)=1 then%>checked<%end if%>>&nbsp;��<input name="canAdminUser" type=radio value="0"  <%if reGroupSetting(28)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>����ɾ���û�1��10������������
</td>
<td height="23" width="40%" class=tablebody1>��<input name="canDelUserTopic" type=radio value="1" <%if reGroupSetting(29)=1 then%>checked<%end if%>>&nbsp;��<input name="canDelUserTopic" type=radio value="0"  <%if reGroupSetting(29)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>���Բ鿴����IP����Դ
</td>
<td height="23" width="40%" class=tablebody1>��<input name="canviewip" type=radio value="1" <%if reGroupSetting(30)=1 then%>checked<%end if%>>&nbsp;��<input name="canviewip" type=radio value="0"  <%if reGroupSetting(30)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>�����޶�IP����
</td>
<td height="23" width="40%" class=tablebody1>��<input name="canadminip" type=radio value="1" <%if reGroupSetting(31)=1 then%>checked<%end if%>>&nbsp;��<input name="canadminip" type=radio value="0"  <%if reGroupSetting(31)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>���Թ����û�Ȩ��
</td>
<td height="23" width="40%" class=tablebody1>��<input name="adminpermission" type=radio value="1" <%if reGroupSetting(42)=1 then%>checked<%end if%>>&nbsp;��<input name="adminpermission" type=radio value="0"  <%if reGroupSetting(42)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>��������ɾ�����ӣ�ǰ̨��
</td>
<td height="23" width="40%" class=tablebody1>��<input name="canbatchtopic" type=radio value="1" <%if reGroupSetting(45)=1 then%>checked<%end if%>>&nbsp;��<input name="canbatchtopic" type=radio value="0"  <%if reGroupSetting(45)=0 then%>checked<%end if%>></td>
</tr>
<tr> 
<th height="23" colspan="2" align=left>��������Ȩ��</th>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>���Է��Ͷ���
</td>
<td height="23" width="40%" class=tablebody1>��<input name="cansendsms" type=radio value="1"  <%if reGroupSetting(32)=1 then%>checked<%end if%>>&nbsp;��<input name="cansendsms" type=radio value="0" <%if reGroupSetting(32)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>��෢���û�
</td>
<td height="23" width="40%" class=tablebody1><input name="Maxsendsms" size=5 type=text value="<%=reGroupSetting(33)%>"></td>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>�������ݴ�С����
</td>
<td height="23" width="40%" class=tablebody1><input name="Maxsmsbody" size=5 type=text value="<%=reGroupSetting(34)%>"> byte</td>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>�����С����
</td>
<td height="23" width="40%" class=tablebody1><input name="Maxsmsbox" size=5 type=text value="<%=reGroupSetting(35)%>"> KB</td>
</tr>
<tr> 
<td height="23" width="60%" class=tablebody2>
</td>
<td height="23" width="40%" class=tablebody2><input type="submit" name="submit" value="�� ��"></td>
</tr>
<input type=hidden value="yes" name="groupaction">
</FORM>
</table>
<%
end if
end sub

sub SaveUserPermission()
if not AdminUserPer then
	Errmsg=Errmsg+"<br>"+"<li>��û��Ȩ��ִ�д˲�����"
	founderr=true
	exit sub
end if
if not isnumeric(request("userid")) then
	errmsg=errmsg+"<br><li>������û�������"
	founderr=true
	exit sub
end if
if not isnumeric(request("boardid")) then
	errmsg=errmsg+"<br><li>����İ��������"
	founderr=true
	exit sub
end if
'���һ�ν�����֤����ҪΪ��֤�����Ƿ��������İ������
if boardmaster and Not FoundUserPer then
	set rs=conn.execute("select boardmaster from board where boardid="&request("boardid"))
	if not(rs.eof and rs.bof) then
		if instr("|"&rs_1(2)&"|","|"&membername&"|")=0 then
			errmsg=errmsg+"<br><li>�����Ǹð���İ�������Ȩ�����û�Ȩ�޲�����"
			founderr=true
			exit sub
		end if
	else
		errmsg=errmsg+"<br><li>�����Ǹð���İ�������Ȩ�����û�Ȩ�޲�����"
		founderr=true
		exit sub
	end if
end if

dim myGroupSetting
myGroupSetting=Request.Form("canview") & "," & Request.Form("canviewuserinfo") & "," & Request.Form("canviewpost") & "," & Request.Form("cannewpost") & "," & Request.Form("canreplymytopic") & "," & Request.Form("canreplytopic") & "," & Request.Form("canpostagree") & "," & Request.Form("canupload") & "," & Request.Form("canpostvote") & "," & Request.Form("canvote") & "," & Request.Form("caneditmytopic") & "," & Request.Form("candelmytopic") & "," & Request.Form("canmovemytopic") & "," & Request.Form("canclosemytopic") & "," & Request.Form("cansearch") & "," & Request.Form("canmailtopic") & "," & Request.Form("canmodify") & "," & Request.Form("cansmallpaper") & "," & Request.Form("candeltopic") & "," & Request.Form("canmovetopic") & "," & Request.Form("canclosetopic") & "," & Request.Form("cantoptopic") & "," & Request.Form("canawardtopic") & "," & Request.Form("canmodifytopic") & "," & Request.Form("canbesttopic") & "," & Request.Form("canAnnounce") & "," & Request.Form("canAdminAnnounce") & "," & Request.Form("canAdminPaper") & "," & Request.Form("canAdminUser") & "," & Request.Form("canDelUserTopic") & "," & Request.Form("canviewip") & "," & Request.Form("canadminip") & "," & Request.Form("cansendsms") & "," & Request.Form("Maxsendsms") & "," & Request.Form("Maxsmsbody") & "," & Request.Form("Maxsmsbox") & "," & Request.Form("canusetitle") & "," & Request.Form("canuseface") & "," & Request.Form("canusesign") & "," & Request.Form("canvieweven") & "," & Request.Form("canuploadnum") & "," & Request.Form("canviewbest") & "," & Request.Form("adminpermission") & "," & request.form("canaward") & "," & request.form("MaxUploadSize") & "," & request.form("canbatchtopic") & "," & request.form("smallpapermoney") & "," & request.form("postagreemoney")
if request("isdefault")=1 then
	conn.execute("delete from UserAccess where uc_boardid="&request("boardid")&" and uc_userid="&request("userid"))
else
	set rs=conn.execute("select * from UserAccess where uc_boardid="&request("boardid")&" and uc_userid="&request("userid"))
	if rs.eof and rs.bof then
		conn.execute("insert into UserAccess (uc_userid,uc_boardid,uc_setting) values ("&request("userid")&","&request("boardid")&",'"&myGroupSetting&"')")
	else
		conn.execute("update UserAccess set uc_setting='"&myGroupSetting&"' where uc_boardid="&request("boardid")&" and uc_userid="&request("userid"))
	end if
end if
sucmsg="<li>�����û�Ȩ�޳ɹ���"
call dvbbs_suc()
end sub

'����ָ����̳��Ϣ
function LastCount(boardid)
dim LastTopic,body,LastRootid,LastPostTime,LastPostUser
dim LastPost,uploadpic_n,Lastpostuserid,Lastid
Dim trs
set trs=conn.execute("select top 1 topic,body,Announceid,dateandtime,username,postuserid,rootid from "&TotalUseTable&" where not locktopic=2 and boardid="&boardid&" and not locktopic=2 order by announceid desc")
if not(trs.eof and trs.bof) then
	Lasttopic=trs(0)
	body=trs(1)
	LastRootid=trs(2)
	LastPostTime=trs(3)
	LastPostUser=trs(4)
	LastPostUserid=trs(5)
	Lastid=trs(6)
else
	LastTopic="��"
	LastRootid=0
	LastPostTime=now()
	LastPostUser="��"
	LastPostUserid=0
	Lastid=0
end if
if Lasttopic="" or isnull(Lasttopic) then LastTopic=left(body,20)
trs.close
set trs=nothing
LastPost=LastPostUser & "$" & LastRootid & "$" & LastPostTime & "$" & replace(left(LastTopic,20),"$","") & "$" & uploadpic_n & "$" & LastPostUserID & "$" & LastID
sql="update board set LastPost='"&LastPost&"' where boardid="&boardid
conn.execute(sql)
end function
%>
