<%
Rem ==========��̳��ҳ����=========
Rem �ж��û�ϵͳ��Ϣ
function usersysinfo(info,getinfo)
if instr(info,";")>0 then
	dim usersys
	usersys=split(info,";")
	usersys(1)=replace(usersys(1),"MSIE","Internet Explorer")
	usersys(2)=replace(usersys(2),")","")
	usersys(2)=replace(usersys(2),"NT 5.1","XP")
	usersys(2)=replace(usersys(2),"NT 5.0","2000")
	usersys(2)=replace(usersys(2),"9x","Me")
	usersys(1)="� �� ����" & Trim(usersys(1))
	usersys(2)="����ϵͳ��" & Trim(usersys(2))
	if getinfo=1 then
		usersysinfo=usersys(1)
	else
		usersysinfo=usersys(2)
	end if
else
	if getinfo=1 then
		usersysinfo="δ֪"
	else
		usersysinfo="δ֪"
	end if
end if
end function

function online(boardid)
dim tmprs
if boardid=0 then
	sql="Select count(*) from online where userid>0"
else
	sql="Select count(*) from online where userid>0 and boardid="&boardid
end if
set tmprs=conn.execute(sql) 
online=tmprs(0) 
set tmprs=nothing 
if isnull(online) then online=0
end function 

function guest(boardid)
dim tmprs
if boardid=0 then
	sql="Select count(*) from online where userid=0"
else
	sql="Select count(*) from online where userid=0 and boardid="&boardid
end if
set tmprs=conn.execute(sql) 
guest=tmprs(0) 
set tmprs=nothing
if isnull(guest) then guest=0
end function
function allonline()
dim tmprs
tmprs=conn.execute("Select count(*) from online") 
allonline=tmprs(0) 
set tmprs=nothing 
if isnull(allonline) then allonline=0
end function

sub onlineuser(online_u,online_g,boardid)
if boardid>0 and (cint(online_u)=1 or cint(online_g)=1) then
response.write "<tr><td colspan=2 class=tablebody1><table cellpadding=2 cellspacing=1 border=0 width=""100%"" style=""word-break:break-all;""><tr>"
end if
dim online_face
dim sip,acip,userip
if cint(online_u)=1 then
	i=0
	'�û���Ϣ
	if boardid=0 then
	sql="select username,ip,stats,UserGroupID,userhidden,userid,startime,lastimebk,actforip from online where userid>0"
	else
	sql="select username,ip,stats,UserGroupID,userhidden,userid,startime,lastimebk,actforip from online where userid>0 and boardid="&boardid
	end if
	set rs=conn.execute(sql)
	do while not rs.eof
	sip=rs(1)
	acip=rs(8)
	if Cint(GroupSetting(30))=0 then
		userip="�����ñ���"
	else
		if acip <> "" then
        		userip=acip
        	else
			userip=sip
	        end if
	end if
	select case rs(3)
	case 1
	online_face="<img src="&Forum_info(7)&Forum_pic(0)&" alt="&rs(5)&" width=12 height=11>"
	case 2
	online_face="<img src="&Forum_info(7)&Forum_pic(1)&" alt="&rs(5)&" width=12 height=11>"
	case 3
	online_face="<img src="&Forum_info(7)&Forum_pic(1)&" alt="&rs(5)&" width=12 height=11>"
	case 8
	online_face="<img src="&Forum_info(7)&Forum_pic(2)&" alt="&rs(5)&" width=12 height=11>"
	case else
	online_face="<img src="&Forum_info(7)&Forum_pic(3)&" width=12 height=11>&nbsp;"
	end select

	if membername=rs(0) then
		response.write "<td width=""14%"">" & online_face&"&nbsp;<a href=dispuser.asp?id="&rs(5)&" target=_blank title=""Ŀǰλ�ã�"&htmlencode(rs(2))&"&#13;&#10;����ʱ�䣺"&rs(6)&"&#13;&#10;�ʱ�䣺"&rs(7)&"&#13;&#10;��ʵ�ɣУ�"&userip&"""><font color="&Forum_pic(5)&">"&htmlencode(rs(0))&"</font></a></td>"
	else
		if rs(4)=1 then
			if boardmaster or master then
			response.write "<td width=""14%"">" & online_face&"&nbsp;<a href=dispuser.asp?id="&rs(5)&" target=_blank title=""Ŀǰλ�ã�"&htmlencode(rs(2))&"<br>����ʱ�䣺"&rs(6)&"<br>�ʱ�䣺"&rs(7)&"<br>��ʵ�ɣУ�"&userip&""">"&htmlencode(rs(0))&"</a></td>"
			else
			response.write "<td width=""14%""><img src="&Forum_info(7)&Forum_pic(4)&" width=12 height=11>&nbsp;<a href=# target=_blank title=""Ŀǰλ�ã�"&htmlencode(rs(2))&"<br>����ʱ�䣺"&rs(6)&"<br>�ʱ�䣺"&rs(7)&"<br>��ʵ�ɣУ�"&userip&""">�����Ա</a></td>"
			end if
		else
			response.write "<td width=""14%"">" & online_face&"&nbsp;<a href=dispuser.asp?id="&rs(5)&" target=_blank title=""Ŀǰλ�ã�"&htmlencode(rs(2))&"<br>����ʱ�䣺"&rs(6)&"<br>�ʱ�䣺"&rs(7)&"<br>��ʵ�ɣУ�"&userip&""">"&htmlencode(rs(0))&"</a></td>"
		end if
	end if
	if i=6 then response.write "</tr><tr>"
	if i>6 then 
		i=1
	else
		i=i+1
	end if
	rs.movenext
	loop
end if
if cint(online_g)=1 then
	dim onlineusername
	i=0
	if boardid=0 then
	sql="select username,ip,stats,UserGroupID,userhidden,userid,startime,lastimebk,actforip,id from online where userid=0"
	else
	sql="select username,ip,stats,UserGroupID,userhidden,userid,startime,lastimebk,actforip,id from online where userid=0 and boardid="&boardid
	end if
	set rs=conn.execute(sql)
  
	online_face="<img src="&Forum_info(7)&Forum_pic(4)&" width=12 height=11>"
	if not (rs.eof and rs.eof) then
		response.write "</tr><tr>"
	end if
	do while not rs.eof
	sip=rs(1)
	acip=rs(8)
	if Cint(GroupSetting(30))=0 then
		userip="�����ñ���"
	else
		if acip <> "" then
        		userip=acip
        	else
			userip=sip
	        end if
	end if
	if trim(session("userid"))<>"" and isnumeric(session("userid")) then
		if int(session("userid"))=int(rs(9)) then
			onlineusername="<font color="&Forum_pic(5)&">����</font>"
		else
			onlineusername="����"
		end if
	else
		onlineusername="����"
	end if
	response.write "<td width=""14%"">" & online_face&"&nbsp;<a href=# target=_blank title=""Ŀǰλ�ã�"&htmlencode(rs(2))&"<br>����ʱ�䣺"&rs(6)&"<br>�ʱ�䣺"&rs(7)&"<br>��ʵ�ɣУ�"&userip&""">"&onlineusername&"</a></td>"

	if i=6 then response.write "</tr><tr>"
	if i>6 then 
		i=1
	else
		i=i+1
	end if
	rs.movenext
	loop
end if
if boardid>0 and (cint(online_u)=1 or cint(online_g)=1) then
	response.write "</tr></TABLE>"
end if
set rs=nothing
end sub
%>