<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<%
Server.ScriptTimeOut=999999
dim topicid
dim tablename
stats="���ӹ���"
topicid=request("topicid")
if request("action")<>"��ջ���վ" then
	if topicid="" or isnull(topicid) then
		Errmsg=Errmsg+"<li>"+"��ѡ��������Ӻ���в�����"
		Founderr=true
	end if
end if
if request("tablename")="topic" then
	tablename="topic"
elseif instr(request("tablename"),"bbs")>0 then
	tablename=request("tablename")
else
	Errmsg=Errmsg+"<li>"+"�����ϵͳ������"
	Founderr=true
end if
if not master then
	Errmsg=Errmsg+"<li>"+"������ϵͳ����Ա��������û�е�½��"
	Founderr=true
end if

call nav()
call head_var(1)
if founderr=true then
	call dvbbs_error()
else
	if request("action")="ɾ��" then
		call delete()
	elseif request("action")="��ԭ" then
		call redel()
	elseif request("action")="��ջ���վ" then
		call Alldel()
	else
		Errmsg=Errmsg+"<li>"+"��ָ�����������"
		Founderr=true
		call dvbbs_error()
	end if
end if
call footer()

'ɾ������վ����
sub delete()
if instr(tablename,"bbs")>0 then
	conn.execute("delete from "&tablename&" where locktopic=2 and Announceid in ("&TopicID&")")
elseif tablename="topic" then
	for i=0 to ubound(allposttable)
	conn.execute("delete from "&allposttable(i)&" where locktopic=2 and rootid in ("&TopicID&")")
	next
	conn.execute("delete from topic where locktopic=2 and topicid in ("&TopicID&")")
end if

sucmsg="<li>���Ӳ����ɹ���<li>���Ĳ�����Ϣ�Ѿ���¼�ڰ���"
call dvbbs_suc()
end sub

'��ԭ����վ����
sub redel()
dim tempnum
if instr(tablename,"bbs")>0 then
	sql="update "&tablename&" set locktopic=0 where Announceid in ("&TopicID&")"
	conn.execute(sql)
	'����ûظ�����Ӧ������ɾ������ͬʱ��ԭ��������
	set rs=conn.execute("select topicid,posttable from topic where locktopic=0 and topicid in (select distinct rootid from "&tablename&" where Announceid in ("&TopicID&"))")
	do while not rs.eof
		conn.execute("update "&rs(1)&" set locktopic=0 where parentid=0 and rootid="&rs(0))
		set trs=conn.execute("select count(*) from "&rs(1)&" where locktopic=0 and not parentid=0 and rootid="&rs(0))
		conn.execute("update topic set child="&trs(0)&",locktopic=0 where topicid="&rs(0))
	rs.movenext
	loop
	set rs=conn.execute("select username from "&tablename&" where Announceid in ("&TopicID&")")
	do while not rs.eof
	sql="update [user] set article=article+1,userWealth=userWealth+"&Forum_user(3)&",userEP=userEP+"&Forum_user(8)&",userdel=userdel-1 where username='"&rs(0)&"'"
	conn.execute(sql)
	rs.movenext
	loop
	set rs=nothing
elseif tablename="topic" then
	dim TotalUseTable
	sql="update topic set locktopic=0 where topicid in ("&TopicID&")"
	conn.execute(sql)
	dim trs
	set rs=conn.execute("select topicid,posttable from topic where topicid in ("&topicid&")")
	do while not rs.eof
		conn.execute("update "&rs(1)&" set locktopic=0 where rootid="&rs(0))
		set trs=conn.execute("select postuserid from "&rs(1)&" where rootid="&rs(0))
		do while not trs.eof
		sql="update [user] set article=article+1,userWealth=userWealth+"&Forum_user(3)&",userEP=userEP+"&Forum_user(8)&",userdel=userdel-1 where userid="&trs(0)
		conn.execute(sql)
		trs.movenext
		loop
	rs.movenext
	loop
	set rs=nothing
	set trs=nothing
end if
sucmsg="<li>���Ӳ����ɹ���<li>���Ĳ�����Ϣ�Ѿ���¼�ڰ���"
call dvbbs_suc()
end sub

'ȫ��ɾ������վ����
sub AllDel()
for i=0 to ubound(allposttable)
	conn.execute("delete from "&allposttable(i)&" where locktopic=2")
next
conn.execute("delete from topic where locktopic=2")
sucmsg="<li>���Ӳ����ɹ���<li>���Ĳ�����Ϣ�Ѿ���¼�ڰ���"
call dvbbs_suc()
end sub
%>