<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<title><%=Forum_info(0)%>--����ҳ��</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta NAME=GENERATOR Content="Microsoft FrontPage 3.0" CHARSET=GB2312>
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0" marginheight="0" marginwidth="0" bgcolor="#DDEEFF">
<%
	Server.ScriptTimeout=9999999
	if not master or session("flag")="" then
		Errmsg=Errmsg+"<br>"+"<li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��½</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
		call dvbbs_error()
	else
		dim tmprs
		dim body
		call main()
		conn.close
		set conn=nothing
	end if

	sub main()
%>
<table cellpadding="2" cellspacing="1" border="0" width="95%" class="tableBorder" align=center>
<tr>
<th align=left colspan=2 height=23>��̳���ݴ���</th>
</tr>
<tr>
<td width="20%" class="forumrow">ע������</td>
<td width="80%" class="forumrow">�����еĲ������ܽ��ǳ����ķ�������Դ�����Ҹ���ʱ��ܳ�������ϸȷ��ÿһ��������ִ�С�</td>
</tr>
<%
	if request("action")="updat" then
		if request("submit")="������̳����" then
		call updateboard()
		elseif request("submit")="�� ��" then
		call fixtopic()
		else
		call updateall()
		end if
		if founderr then
		response.write errmsg
		else
		response.write body
		end if
	elseif request("action")="delboard" then
		if isnumeric(request("boardid")) then
		conn.execute("update topic set locktopic=2 where boardid="&request("boardid"))
		for i=0 to ubound(AllPostTable)
		conn.execute("update "&AllPostTable(i)&" set locktopic=2 where boardid="&request("boardid"))
		next
		end if
		response.write "<tr><td align=left colspan=2 height=23 class=forumrow>�����̳���ݳɹ����뷵�ظ����������ݣ�</td></tr>"
	elseif request("action")="updateuser" then
%>
<FORM METHOD=POST ACTION="?action=updateuserinfo">
<tr> 
<th align=left colspan=2 height=23>�����û�����</th>
</tr>
<tr>
<td width="20%" class="forumrow">���¼����û�����</td>
<td width="80%" class="forumrow">ִ�б�����������<font color=red>��ǰ��̳���ݿ�</font>�������¼��������û���������������</td>
</tr>
<tr>
<td width="20%" class="forumrow">��ʼ�û�ID</td>
<td width="80%" class="forumrow"><input type=text name="beginID" value="1" size=10>&nbsp;�û�ID��������д�������һ��ID�ſ�ʼ�����޸�</td>
</tr>
<tr>
<td width="20%" class="forumrow">�����û�ID</td>
<td width="80%" class="forumrow"><input type=text name="endID" value="100" size=10>&nbsp;�����¿�ʼ������ID֮����û����ݣ�֮�����ֵ��ò�Ҫѡ�����</td>
</tr>
<tr>
<td width="20%" class="forumrow"></td>
<td width="80%" class="forumrow"><input type="submit" name="Submit" value="���¼����û�����"></td>
</tr>
</form>

<FORM METHOD=POST ACTION="?action=updateuserinfo">
<tr>
<td width="20%" class="forumrow" valign=top>�����û��ȼ�</td>
<td width="80%" class="forumrow">ִ�б�����������<font color=red>��ǰ��̳���ݿ�</font>�û�������������̳�ĵȼ��������¼����û��ȼ�����������Ӱ��ȼ�Ϊ������������ܰ��������ݡ�</td>
</tr>
<tr>
<td width="20%" class="forumrow">��ʼ�û�ID</td>
<td width="80%" class="forumrow"><input type=text name="beginID" value="1" size=10>&nbsp;�û�ID��������д�������һ��ID�ſ�ʼ�����޸�</td>
</tr>
<tr>
<td width="20%" class="forumrow">�����û�ID</td>
<td width="80%" class="forumrow"><input type=text name="endID" value="100" size=10>&nbsp;�����¿�ʼ������ID֮����û����ݣ�֮�����ֵ��ò�Ҫѡ�����</td>
</tr>
<tr>
<td width="20%" class="forumrow"></td>
<td width="80%" class="forumrow"><input type="submit" name="Submit" value="�����û��ȼ�"></td>
</tr>
</form>

<FORM METHOD=POST ACTION="?action=updateuserinfo">
<tr>
<td width="20%" class="forumrow" valign=top>�����û���Ǯ/����/����</td>
<td width="80%" class="forumrow">ִ�б�����������<font color=red>��ǰ��̳���ݿ�</font>�û��ķ�����������̳������������¼����û��Ľ�Ǯ/����/������������Ҳ�����¼��������������ܰ���������<BR>ע�⣺���Ƽ��û����б������������������ݺܶ��ʱ���뾡����Ҫʹ�ã����ұ������Ը�������ɾ�����ӵ�������Ӧ��ֵ�������㣬ֻ�ǰ��շ������ܵ���̳��ֵ���ý������㣬�������ز�����<font color=red>���ұ�������������û���Ϊ�������ͷ���ԭ�����Ա���û���ֵ���޸ġ�</font></td>
</tr>
<tr>
<td width="20%" class="forumrow">��ʼ�û�ID</td>
<td width="80%" class="forumrow"><input type=text name="beginID" value="1" size=10>&nbsp;�û�ID��������д�������һ��ID�ſ�ʼ�����޸�</td>
</tr>
<tr>
<td width="20%" class="forumrow">�����û�ID</td>
<td width="80%" class="forumrow"><input type=text name="endID" value="100" size=10>&nbsp;�����¿�ʼ������ID֮����û����ݣ�֮�����ֵ��ò�Ҫѡ�����</td>
</tr>
<tr>
<td width="20%" class="forumrow"></td>
<td width="80%" class="forumrow"><input type="submit" name="Submit" value="�����û���Ǯ/����/����"></td>
</tr>
</FORM>
<%
	elseif request("action")="updateuserinfo" then
		if request("submit")="���¼����û�����" then
		call updateTopic()
		elseif request("submit")="�����û��ȼ�" then
		call updategrade()
		else
		call updatemoney()
		end if
		if founderr then
		response.write errmsg
		else
		response.write body
		end if
	else
%>
<tr> 
<th align=left colspan=2 height=23>������̳����</th>
</tr>

<form action="admin_update.asp?action=updat" method=post>
<tr>
<td width="20%" class="forumrow">���·���̳����</td>
<td width="80%" class="forumrow"><input type="submit" name="Submit" value="������̳����"><BR><BR>���ｫ���¼���ÿ����̳����������ͻظ������������ӣ����ظ���Ϣ�ȣ�����ÿ��һ��ʱ������һ�Ρ�</td>
</tr>
<tr>
<td width="20%" class="forumrow">��������̳����</td>
<td width="80%" class="forumrow"><input type="submit" name="Submit" value="������̳������"><BR><BR>���ｫ���¼���������̳����������ͻظ������������ӣ��������û��ȣ�����ÿ��һ��ʱ������һ�Ρ�</td>
</tr>
<tr> 
<th align=left colspan=2 height=23>�޸�����(�޸�ָ����Χ�����ӵ����ظ�����)</th>
</tr>
<tr>
<td width="20%" class="forumrow">��ʼ��ID��</td>
<td width="80%" class="forumrow"><input type=text name="beginID" value="1" size=10>&nbsp;��������ID��������д�������һ��ID�ſ�ʼ�����޸�</td>
</tr>
<tr>
<td width="20%" class="forumrow">������ID��</td>
<td width="80%" class="forumrow"><input type=text name="EndID" value="1000" size=10>&nbsp;�����¿�ʼ������ID֮����������ݣ�֮�����ֵ��ò�Ҫѡ�����</td>
</tr>
<tr>
<td width="20%" class="forumrow"></td>
<td width="80%" class="forumrow"><input type="submit" name="Submit" value="�� ��"></td>
</tr>
</form>
<%
	end if
%>
</table><BR><BR>
<%
	end sub

sub updateboard()
dim allarticle
dim alltopic
dim Ers,Esql
dim Maxid
dim lastpost
dim username,dateandtime,rootid,topic,announceid,postuserid
set rs=server.createobject("adodb.recordset")
if isnumeric(request("boardid")) and request("boardid")<>"" then
	sql="select boardid,boardtype from board where boardid="&request("boardid")
else
	sql="select boardid,boardtype from board"
end if
set rs=conn.execute(sql)
if rs.eof and rs.bof then
	Errmsg=Errmsg+"<tr><td align=left colspan=2 height=23 class=forumrow>��û����̳���棬�뵽��̳������ӣ�</td></tr>"
	founderr=true
	exit sub
else
	do while not rs.eof
    allarticle=BoardAnnounceNum(rs("BoardID"))
    alltopic=BoardTopicNum(rs("BoardID"))
	sql="select top 1 username,dateandtime,topic,Announceid,PostUserID,rootid,body from "&NowUseBBS&" where boardid="&rs("boardid")&" order by Announceid desc"
	set Ers=conn.execute(sql)
	if ers.eof and ers.bof then
		username="��"
		dateandtime=now()
		rootid=0
		topic="��"
		Announceid=0
		postuserid=0
	else
		username=Ers("username")
		dateandtime=Ers("dateandtime")
		rootid=Ers("rootid")
		if Ers("topic")="" then
			topic=left(Ers("body"),20)
		else
			topic=left(Ers("topic"),20)
		end if
		Announceid=ers("Announceid")
		postuserid=ers("postuserid")
	end if
	LastPost=username & "$" & Announceid & "$" & dateandtime & "$" & replace(topic,"$","") & "$$" & postuserid & "$" & rootid 
	Esql="update board set lastbbsnum="&allarticle&",lasttopicnum="&alltopic&",TodayNum="&todays(rs("boardid"))&",LastPost='"&LastPost&"' where boardid="&rs("boardid")&""
	conn.execute(Esql)
	body=body & "<tr><td colspan=2 class=forumrow>������̳���ݳɹ���"&rs("boardtype")&"����"&allarticle&"ƪ���ӣ�"&alltopic&"ƪ���⣬������"&todays(rs("boardid"))&"ƪ���ӡ�</td></tr>"
	rs.movenext
	loop
end if
set rs=nothing
set ers=nothing
end sub

sub updateall()
sql="update config set TopicNum="&topicnum()&",BbsNum="&announcenum()&",TodayNum="&alltodays()&",UserNum="&allusers()&",lastUser='"&newuser()&"'"
conn.execute(sql)
body="<tr><td colspan=2 class=forumrow>��������̳���ݳɹ���ȫ����̳����"&announcenum()&"ƪ���ӣ�"&topicnum()&"ƪ���⣬������"&alltodays()&"ƪ���ӣ���"&allusers()&"�û������¼���Ϊ"&newuser()&"��</td></tr>"
end sub

sub fixtopic()
if not isnumeric(request.form("beginid")) then
	body="<tr><td colspan=2 class=forumrow>����Ŀ�ʼ������</td></tr>"
	exit sub
end if
if not isnumeric(request.form("endid")) then
	body="<tr><td colspan=2 class=forumrow>����Ľ���������</td></tr>"
	exit sub
end if
if clng(request.form("beginid"))>clng(request.form("endid")) then
	body="<tr><td colspan=2 class=forumrow>��ʼIDӦ�ñȽ���IDС��</td></tr>"
	exit sub
end if
dim TotalUseTable,Ers
dim username,dateandtime,rootid,announceid,postuserid,lastpost,topic
set rs=server.createobject("adodb.recordset")
sql="select topicid,PostTable from topic where topicid>="&request.form("beginid")&" and topicid<="&request.form("endid")
set rs=conn.execute(sql)
if rs.eof and rs.bof then
	body="<tr><td colspan=2 class=forumrow>�Ѿ�����¼����β�ˣ���������£�</td></tr>"
	exit sub
end if
do while not rs.eof
	sql="select top 1 username,dateandtime,topic,Announceid,PostUserID,rootid,body from "&rs(1)&" where rootid="&rs(0)&" order by Announceid desc"
	set ers=conn.execute(sql)
	if not (ers.eof and ers.bof) then
		username=Ers("username")
		dateandtime=Ers("dateandtime")
		rootid=Ers("rootid")
		topic=left(Ers("body"),20)
		Announceid=ers("Announceid")
		postuserid=ers("postuserid")
		LastPost=username & "$" & Announceid & "$" & dateandtime & "$" & replace(topic,"$","") & "$$" & postuserid & "$" & rootid
		conn.execute("update topic set LastPost='"&LastPost&"' where topicid="&rs(0))
	end if
rs.movenext
loop
set ers=nothing
set rs=nothing
%>
<form action="admin_update.asp?action=updat" method=post>
<tr> 
<th align=left colspan=2 height=23>�����޸�����(�޸�ָ����Χ�����ӵ����ظ�����)</th>
</tr>
<tr>
<td width="20%" class="forumrow">��ʼ��ID��</td>
<td width="80%" class="forumrow"><input type=text name="beginID" value="<%=request.form("endid")+1%>" size=5>&nbsp;��������ID��������д�������һ��ID�ſ�ʼ�����޸�</td>
</tr>
<tr>
<td width="20%" class="forumrow">������ID��</td>
<td width="80%" class="forumrow"><input type=text name="EndID" value="<%=request.form("endid")+(request.form("endid")-request.form("beginid"))+1%>" size=5>&nbsp;�����¿�ʼ������ID֮����������ݣ�֮�����ֵ��ò�Ҫѡ�����</td>
</tr>
<tr>
<td width="20%" class="forumrow"></td>
<td width="80%" class="forumrow"><input type="submit" name="Submit" value="�� ��"></td>
</tr>
</form>
<%
end sub

'����̳��������
function todays(boardid)
tmprs=conn.execute("Select count(announceid) from "&NowUseBBS&" Where boardid="&boardid&" and not locktopic=2 and datediff('d',dateandtime,Now())=0")
todays=tmprs(0)
set tmprs=nothing
if isnull(todays) then todays=0
end function

'ȫ����̳��������
function alltodays()
tmprs=conn.execute("Select count(announceid) from "&NowUseBBS&" Where not locktopic=2 and datediff('d',dateandtime,Now())=0")
alltodays=tmprs(0)
set tmprs=nothing
if isnull(alltodays) then alltodays=0
end function

'����ע���û�����
function allusers() 
tmprs=conn.execute("Select count(userid) from [user]") 
allusers=tmprs(0) 
set tmprs=nothing 
if isnull(allusers) then allusers=0 
end function
'����ע���û�
function newuser()
sql="Select top 1 username from [user] order by userid desc"
set tmprs=conn.execute(sql)
if tmprs.eof and tmprs.bof then
	newuser="û�л�Ա"
else
   	newuser=tmprs("username")
end if
set tmprs=nothing
end function 

'������̳����
function AnnounceNum()
dim AnnNum
AnnNum=0
AnnounceNum=0
For i=0 to ubound(AllPostTable)
	tmprs=conn.execute("Select Count(announceID) from "&AllPostTable(i)&" where not locktopic=2") 
	AnnNum=tmprs(0)
	set tmprs=nothing 
	if isnull(AnnNum) then AnnNum=0
	AnnounceNum=AnnounceNum + AnnNum
next
set tmprs=nothing
end function
'����̳����
function BoardAnnounceNum(boardid)
dim BoardAnnNum
BoardAnnNum=0
BoardAnnounceNum=0
For i=0 to ubound(AllPostTable)
	tmprs=conn.execute("Select Count(announceID) from "&AllPostTable(i)&" where boardid="&boardid&" and not locktopic=2") 
	BoardAnnNum=tmprs(0) 
	set tmprs=nothing 
	if isnull(BoardAnnNum) then BoardAnnNum=0
	BoardAnnounceNum=BoardAnnounceNum + BoardAnnNum
next
set tmprs=nothing
end function

'������̳����
function TopicNum() 
tmprs=conn.execute("Select Count(topicid) from topic where not locktopic=2") 
TopicNum=tmprs(0) 
set tmprs=nothing 
if isnull(TopicNum) then TopicNum=0 
end function 
'����̳����
function BoardTopicNum(boardid) 
tmprs=conn.execute("Select Count(topicid) from topic where boardid="&boardid&" and not locktopic=2")
BoardTopicNum=tmprs(0) 
set tmprs=nothing 
if isnull(BoardTopicNum) then BoardTopicNum=0 
end function

'�����û�������
sub updateTopic()
if not isnumeric(request.form("beginid")) then
	body="<tr><td colspan=2 class=forumrow>����Ŀ�ʼ������</td></tr>"
	exit sub
end if
if not isnumeric(request.form("endid")) then
	body="<tr><td colspan=2 class=forumrow>����Ľ���������</td></tr>"
	exit sub
end if
if clng(request.form("beginid"))>clng(request.form("endid")) then
	body="<tr><td colspan=2 class=forumrow>��ʼIDӦ�ñȽ���IDС��</td></tr>"
	exit sub
end if
dim userTopic
sql="select userid from [user] where userid>="&request.form("beginid")&" and userid<="&request.form("endid")
set rs=conn.execute(sql)
if rs.eof and rs.bof then
	body="<tr><td colspan=2 class=forumrow>�Ѿ�����¼����β�ˣ���������£�</td></tr>"
	exit sub
end if
do while not rs.eof
	userTopic=Userallnum(rs(0))
	conn.execute("update [user] set article="&userTopic&" where userid="&rs(0))
rs.movenext
loop
set rs=nothing
%>
<FORM METHOD=POST ACTION="?action=updateuserinfo">
<tr> 
<th align=left colspan=2 height=23>���������û�����</th>
</tr>
<tr>
<td width="20%" class="forumrow">���¼����û�����</td>
<td width="80%" class="forumrow">ִ�б�����������<font color=red>��ǰ��̳���ݿ�</font>�������¼��������û���������������</td>
</tr>
<tr>
<td width="20%" class="forumrow">��ʼ�û�ID</td>
<td width="80%" class="forumrow"><input type=text name="beginID" value="<%=request.form("endid")+1%>" size=10>&nbsp;�û�ID��������д�������һ��ID�ſ�ʼ�����޸�</td>
</tr>
<tr>
<td width="20%" class="forumrow">�����û�ID</td>
<td width="80%" class="forumrow"><input type=text name="endID" value="<%=request.form("endid")+(request.form("endid")-request.form("beginid"))+1%>" size=10>&nbsp;�����¿�ʼ������ID֮����û����ݣ�֮�����ֵ��ò�Ҫѡ�����</td>
</tr>
<tr>
<td width="20%" class="forumrow"></td>
<td width="80%" class="forumrow"><input type="submit" name="Submit" value="���¼����û�����"></td>
</tr>
</form>
<%
end sub

'�����û���Ǯ/����/����
sub updatemoney()
if not isnumeric(request.form("beginid")) then
	body="<tr><td colspan=2 class=forumrow>����Ŀ�ʼ������</td></tr>"
	exit sub
end if
if not isnumeric(request.form("endid")) then
	body="<tr><td colspan=2 class=forumrow>����Ľ���������</td></tr>"
	exit sub
end if
if clng(request.form("beginid"))>clng(request.form("endid")) then
	body="<tr><td colspan=2 class=forumrow>��ʼIDӦ�ñȽ���IDС��</td></tr>"
	exit sub
end if
dim userTopic,userReply,userWealth
dim userEP,userCP

sql="select logins,userid from [user] where userid>="&request.form("beginid")&" and userid<="&request.form("endid")
set rs=conn.execute(sql)
do while not rs.eof
	userTopic=UserTopicNum(rs(1))
	userreply=UserReplyNum(rs(1))
	userwealth=rs(0)*Forum_user(4) + userTopic*Forum_user(1) + userreply*Forum_user(2)
	userEP=rs(0)*Forum_user(9) + userTopic*Forum_user(6) + userreply*Forum_user(7)
	userCP=rs(0)*Forum_user(14) + userTopic*Forum_user(11) + userreply*Forum_user(12)
	conn.execute("update [user] set userWealth="&userWealth&",userep="&userep&",usercp="&usercp&" where userid="&rs(1))
rs.movenext
loop
set rs=nothing
%>
<FORM METHOD=POST ACTION="?action=updateuserinfo">
<tr> 
<th align=left colspan=2 height=23>���������û�����</th>
</tr>
<tr>
<td width="20%" class="forumrow" valign=top>�����û���Ǯ/����/����</td>
<td width="80%" class="forumrow">ִ�б�����������<font color=red>��ǰ��̳���ݿ�</font>�û��ķ�����������̳������������¼����û��Ľ�Ǯ/����/������������Ҳ�����¼��������������ܰ���������<BR>ע�⣺���Ƽ��û����б������������������ݺܶ��ʱ���뾡����Ҫʹ�ã����ұ������Ը�������ɾ�����ӵ�������Ӧ��ֵ�������㣬ֻ�ǰ��շ������ܵ���̳��ֵ���ý������㣬�������ز�����<font color=red>���ұ�������������û���Ϊ�������ͷ���ԭ�����Ա���û���ֵ���޸ġ�</font></td>
</tr>
<tr>
<td width="20%" class="forumrow">��ʼ�û�ID</td>
<td width="80%" class="forumrow"><input type=text name="beginID" value="<%=request.form("endid")+1%>" size=10>&nbsp;�û�ID��������д�������һ��ID�ſ�ʼ�����޸�</td>
</tr>
<tr>
<td width="20%" class="forumrow">�����û�ID</td>
<td width="80%" class="forumrow"><input type=text name="endID" value="<%=request.form("endid")+(request.form("endid")-request.form("beginid"))+1%>" size=10>&nbsp;�����¿�ʼ������ID֮����û����ݣ�֮�����ֵ��ò�Ҫѡ�����</td>
</tr>
<tr>
<td width="20%" class="forumrow"></td>
<td width="80%" class="forumrow"><input type="submit" name="Submit" value="�����û���Ǯ/����/����"></td>
</tr>
</form>
<%
end sub

'�����û��ȼ�
sub updategrade()
if not isnumeric(request.form("beginid")) then
	body="<tr><td colspan=2 class=forumrow>����Ŀ�ʼ������</td></tr>"
	exit sub
end if
if not isnumeric(request.form("endid")) then
	body="<tr><td colspan=2 class=forumrow>����Ľ���������</td></tr>"
	exit sub
end if
if clng(request.form("beginid"))>clng(request.form("endid")) then
	body="<tr><td colspan=2 class=forumrow>��ʼIDӦ�ñȽ���IDС��</td></tr>"
	exit sub
end if
dim oldMinArticle
oldMinArticle=0
set rs=conn.execute("select * from usertitle order by MinArticle desc")
do while not rs.eof
	conn.execute("update [user] set userclass='"&rs("usertitle")&"',titlepic='"&rs("titlepic")&"' where (userid>="&request.form("beginid")&" and userid<="&request.form("endid")&") and (article<"&oldMinArticle&" and article>="&rs("MinArticle")&" ) and usergroupid=4")
	oldMinArticle=rs("MinArticle")
rs.movenext
loop
rs.close
set rs=nothing
%>
<FORM METHOD=POST ACTION="?action=updateuserinfo">
<tr> 
<th align=left colspan=2 height=23>���������û�����</th>
</tr>
<tr>
<td width="20%" class="forumrow" valign=top>�����û��ȼ�</td>
<td width="80%" class="forumrow">ִ�б�����������<font color=red>��ǰ��̳���ݿ�</font>�û�������������̳�ĵȼ��������¼����û��ȼ�����������Ӱ��ȼ�Ϊ������������ܰ��������ݡ�</td>
</tr>
<tr>
<td width="20%" class="forumrow">��ʼ�û�ID</td>
<td width="80%" class="forumrow"><input type=text name="beginID" value="<%=request.form("endid")+1%>" size=10>&nbsp;�û�ID��������д�������һ��ID�ſ�ʼ�����޸�</td>
</tr>
<tr>
<td width="20%" class="forumrow">�����û�ID</td>
<td width="80%" class="forumrow"><input type=text name="endID" value="<%=request.form("endid")+(request.form("endid")-request.form("beginid"))+1%>" size=10>&nbsp;�����¿�ʼ������ID֮����û����ݣ�֮�����ֵ��ò�Ҫѡ�����</td>
</tr>
<tr>
<td width="20%" class="forumrow"></td>
<td width="80%" class="forumrow"><input type="submit" name="Submit" value="�����û��ȼ�"></td>
</tr>
</form>
<%
end sub

'�û�����������
function UserTopicNum(userid)
dim topicnum
topicnum=0
usertopicnum=0
set tmprs=conn.execute("select count(*) from topic where PostUserID="&userid&" and not locktopic=2")
TopicNum=tmprs(0)
if isnull(TopicNum) then TopicNum=0
UserTopicNum=UserTopicNum + TopicNum
set tmprs=nothing
end function
'�û����лظ���
function UserReplyNum(userid)
dim replynum
replynum=0
userreplynum=0
For i=0 to ubound(AllPostTable)
	set tmprs=conn.execute("select count(announceid) from "&AllPostTable(i)&" where not ParentID=0 and not locktopic=2 and PostUserID="&userid)
	replyNum=tmprs(0)
	if isnull(replyNum) then replyNum=0
	UserReplyNum=UserReplyNum + replynum
next
set tmprs=nothing
end function
'�û���������
function Userallnum(userid)
dim allnum
allnum=0
userallnum=0
For i=0 to ubound(AllPostTable)
	set tmprs=conn.execute("select count(announceid) from "&AllPostTable(i)&" where not locktopic=2 and  PostUserID="&userid)
	allnum=tmprs(0)
	if isnull(allnum) then allnum=0
	userallnum=userallnum+allnum
next
set tmprs=nothing
end function
%>