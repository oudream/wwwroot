<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<%
'=========================================================
' File: index.asp
' Version:5.0
' Date: 2002-9-21
' Script Written by satan
'=========================================================
' Copyright (C) 2001,2002 AspSky.Net. All rights reserved.
' Web: http://www.aspsky.net,http://www.dvbbs.net
' Email: info@aspsky.net,eway@aspsky.net
'=========================================================

dim boardstat
dim onlinenum,guestnum
onlinenum=online(0)
guestnum=guest(0)
Stats="��̳��ҳ"
call activeonline()
Rem ��ҳ������Ϣ
call nav()
if isnull(lastlogin) or lastlogin="" then lastlogin=now()
if Cint(allonline())>Cint(Maxonline) then
	conn.execute("update config set Maxonline="&allonline()&",MaxonlineDate=Now()")
end if
%>
<br>
<TABLE cellSpacing=0 cellPadding=0 width=<%=Forum_body(12)%> border=0 align=center>
<tr>
<td align=center width=100% valign=middle>
<SCRIPT LANGUAGE='JavaScript' SRC='js/fader.js' TYPE='text/javascript'></script>
<SCRIPT LANGUAGE='JavaScript' TYPE='text/javascript'>
prefix="";
arNews = ["<%
sql="select top 1 title,addtime from bbsnews where boardid=0 order by id desc"
set rs=conn.execute(sql)
if rs.bof and rs.eof then
	response.write "<b><a href='announcements.asp?boardid="& boardid &"' target=_blank><ACRONYM TITLE='��ǰû�й���'>��ǰû�й���</ACRONYM></a></b> ("& now() &")"
else
	response.write "<b><a href='announcements.asp?boardid="& boardid &"' target=_blank>"& rs("title") &"</a></b> ("& rs("addtime") &")"
end if
set rs=nothing
%>","","��ӭ����<%=Forum_info(0)%> <%=boardtype%>","",
"�Ͻ�����ʹ�ô��Ի��Υ�߾�Ȱ����Ч��������ID��",""
]
</SCRIPT>
<div id="elFader" style="position:relative;visibility:hidden; height:16" ></div>
</td>
</tr>
</table>
<TABLE border=0 width="<%=Forum_body(12)%>" align=center>
<%if not founduser then%>
<tr><td colspan=2><BR>
��ӭ���� <B><%=Forum_info(0)%></B>.<BR>����������һ�η��ʱ�վ, ������̳�Ϸ��������Ķ�<a href=boardhelp.asp>��̳����</a>. �������Ҫ<a href=reg.asp>ע��</a>���ܷ�����Ϣ.������Ѿ�ע������<a href=login.asp>��½</a>.
<%else%>
<tr><td width=65% ><BR>
>> ��ӭ <B><%=membername%></B>  [ <a href="JavaScript:openScript('messanger.asp?action=new',500,400)">������</a> :: <a href="topicwithme.asp?s=2">���������</a> :: <a href="topicwithme.asp?s=1">���������</a> :: <a href="dispuser.asp?id=<%=userid%>&boardid=<%=boardid%>&action=permission">������ʲô</a> ]
</td><td width=35% align=right><BR>
<%if Cint(newincept())>Cint(0) then%>
<bgsound src="pic/mail.wav" border=0>
<img src=<%=Forum_info(7)%>msg_new_bar.gif> <a href="usersms.asp?action=inbox">�ҵ��ռ���</a> (<a href="javascript:openScript('messanger.asp?action=read&id=<%=inceptid(1)%>&sender=<%=inceptid(2)%>',500,400)"><font color="<%=Forum_body(8)%>"><%=newincept()%> ��</font></a>)
<%else%>
<img src=<%=Forum_info(7)%>msg_no_new_bar.gif> <a href="usersms.asp?action=inbox">�ҵ��ռ���</a> (<font color=gray>0 ��</font>)
<%end if%>
<%end if%>
</td></tr>
<tr><td colspan=2 height=1 bgcolor="<%=Forum_body(27)%>">

</td></tr>
<tr><td colspan=2 height=5>

</td></tr>
<%
sql="select top 1 TopicNum,BbsNum,TodayNum,UserNum,lastUser,yesterdaynum,maxpostnum,maxpostdate from config where active=1"
set rs=conn.execute(sql)
if rs(2)>rs(6) then conn.execute("update config set maxpostnum=todaynum,maxpostdate=Now()")
%>
<TR><TD style="line-height: 20px;">
��ӭ�¼����Ա <a href=dispuser.asp?name=<%= htmlencode(rs(4)) %> target=_blank><b><%= htmlencode(rs(4)) %></b></a>&nbsp;[<a href="toplist.asp?orders=2">�½�����</a>]<BR>��̳���� <B><%= rs(3)%></B> λע���Ա , ����������<b><%= rs(0) %></b> , ����������<b><%= rs(1) %></b><BR>������̳��������<FONT COLOR="<%=Forum_body(8)%>"><B><%=rs(2)%></B></FONT> , ���շ�����<B><%=rs(5)%></B> , ����շ�����<B><%=rs(6)%></B></td><TD valign=bottom align=right style="line-height: 20px;"><a href=queryresult.asp?stype=3>�鿴����</a> , ���Ż��� , <a href=toplist.asp?orders=1>��������</a> , <a href=toplist.asp?orders=7>�û��б�</a><BR>�����һ�η������ڣ�<%=Now()%>
</TD></TR>
</TABLE>
<%
if Cint(Forum_setting(4))=1 then
call main_2()
else
call main_1()
end if
%>
<%sub main_1()%>
<table cellspacing=1 cellpadding=3 align=center class=tableBorder1>
<%if not founduser then%>
<form action=login.asp?action=chk method=post>
<tr>
<th align=left id=tabletitlelink height=25 style="font-weight:normal">
<b>-=> ���ٵ�¼���</b>
[<a href=reg.asp>ע���û�</a>]��[<a href=lostpass.asp style="CURSOR: help">��������</a>]
</th>
</tr>
<tr>
<td class=tablebody1 height=40 width="100%">
&nbsp;�û�����<input maxLength=16 name=username size=12>�������룺<input maxLength=20 name=password size=12 type=password>����<select name=CookieDate><option selected value=0>������</option><option value=1>����һ��</option><option value=2>����һ��</option><option value=3>����һ��</option></select>����<input type=submit name=submit value="�� ½">
</td>
</tr>
</form>
</table>
<%
end if
dim sql1,rs1
dim boardview
sql="select id,class from class order by orders"
set rs=conn.execute(sql)
do while not rs.eof
	response.write "<table cellspacing=1 cellpadding=0 align=center class=tableBorder1>"
	response.write "<TR><Th colSpan=2 height=25 align=left>&nbsp;"&rs(1)&"</Th></TR>"
	if boardmaster or master then
	sql1="select boardid,boardtype,class,readme,lastbbsnum,boardmaster,lockboard,lasttopicnum,indexIMG,LastPost,todaynum from board "
	sql1=sql1&" where class="&rs(0)&" order by orders"
	else
	sql1="select boardid,boardtype,class,readme,lastbbsnum,boardmaster,lockboard,lasttopicnum,indexIMG,LastPost,todaynum from board "
	sql1=sql1&" where class="&rs(0)&" and lockboard<>1 order by orders"
	end if
	set rs1=conn.execute(sql1)
	do while not rs1.eof
	LastPostInfo=split(rs1(9),"$")
	if not isdate(LastPostInfo(2)) then LastPostInfo(2)=Now()
	if rs1(5)<>"" then
		master_1=split(rs1(5), "|")
		for i = 0 to ubound(master_1)
		if i>6 then
			master_2=master_2
		else
			master_2=master_2 & "<a href=dispuser.asp?name="&master_1(i)&" target=_blank>"&master_1(i)&"</a>&nbsp;"
		end if
		next
		if i>7 then master_2=master_2 & "<font color=gray>More...</font>"
	else
		master_2="����"
	end if
'=========================BoardInfo============================
%>
<TR><TD align=middle width=46 class=tablebody1>
<%
	if rs1(6)=2 then
		if datediff("h",lastlogin,LastPostInfo(2))=0 then
		boardstat="<img src="&Forum_info(7)&Forum_pic(15)&" alt=��������>"
		else
		boardstat="<img src="&Forum_info(7)&Forum_pic(14)&" alt=��������>"
		end if
	else
		if datediff("h",lastlogin,LastPostInfo(2))=0 then
		boardstat="<img src="&Forum_info(7)&Forum_pic(7)&" alt=��������>"
		else
		boardstat="<img src="&Forum_info(7)&Forum_pic(6)&" alt=��������>"
		end if
	end if
	response.write boardstat
%>
</TD>
<TD vAlign=top width=* class=tablebody1>
<TABLE cellSpacing=0 cellPadding=2 width=100% border=0>
<tr><td class=tablebody1 width=*><a href="list.asp?boardid=<%=rs1(0)%>">
<font color=<%=Forum_body(14)%>><%=rs1(1)%></font></a>
</td>
<%
response.write "<td width=40 rowspan=2 align=center class=tablebody1>"
if rs1(8)<>"" then
	response.write "<table align=left><tr><td><a href=list.asp?boardid="&rs1(0)&"><img src="&rs1(8)&" align=top border=0></a></td><td width=20></td></tr></table>"
end if
response.write "</td><td width=200 rowspan=2 class=tablebody1>���⣺<a href=""dispbbs.asp?Boardid="&rs1(0)&"&ID="&LastPostInfo(6)&"&replyID="&LastPostInfo(1)&"&skin=1"">"&htmlencode(left(LastPostInfo(3),10))&"..</a><BR>���ߣ�<a href=""dispuser.asp?id="&htmlencode(LastPostInfo(5))&""" target=_blank>"&htmlencode(LastPostInfo(0))&"</a><BR>���ڣ�" & FormatDateTime(LastPostInfo(2),1) & "��" & FormatDateTime(LastPostInfo(2),4)&"&nbsp;<a href=""dispbbs.asp?Boardid="&rs1(0)&"&ID="&LastPostInfo(6)&"&replyID="&LastPostInfo(1)&"&skin=1""><IMG border=0 src="""&Forum_info(7)&"lastpost.gif"" title=""ת����"&htmlencode(LastPostInfo(3))&"""></a></TD></TR><TR><TD width=*><FONT face=Arial><img src="&Forum_info(7)&"Forum_readme.gif align=middle> "
response.write rs1(3)
response.write"</FONT></TD>"

response.write "</TR><TR><TD class=tablebody2 height=20 width=*>������"&master_2&""

response.write "</TD><td width=40 align=center class=tablebody2>&nbsp;</td><TD vAlign=middle class=tablebody2 width=200><table width=100% border=0><tr>"

response.write "<td width=25% vAlign=middle><img src="&Forum_info(7)&"Forum_today.gif alt=������ align=absmiddle>&nbsp;<font color="&Forum_body(8)&">"&rs1(10)&"</font></td><td width=30% vAlign=middle><img src="&Forum_info(7)&"forum_topic.gif alt=���� border=0  align=absmiddle>&nbsp;"&rs1(7)&"</td><td width=45% vAlign=middle><img src="&Forum_info(7)&"Forum_post.gif alt=���� border=0 align=absmiddle>&nbsp;"&rs1(4)&"</td></tr>"
response.write "</table></TD></TR>"&_
               "</TBODY></TABLE>"&_
               "</td></tr>"
'============================End===============================
	master_2=""
	rs1.movenext
	loop
	response.write "</table><BR>"
rs.movenext
loop
set rs=nothing
set rs1=nothing
end sub

sub main_2()
%>
<table cellspacing=1 cellpadding=3 align=center class=tableBorder1>
<%if not founduser then%>
<form action=login.asp?action=chk method=post>
<tr>
<th align=left id=tabletitlelink colSpan=6 height=25 style="font-weight:normal">
<b>-=> ���ٵ�¼���</b>
[<a href=reg.asp>ע���û�</a>]��[<a href=lostpass.asp style="CURSOR: help">��������</a>]
</th>
</tr>
<tr>
<td align=middle class=tablebody2 width=26>
<img src="<%=Forum_info(7)&Forum_statePic(6)%>" width=16 height=16>
</td>
<td class=tablebody1 colSpan=5 width=*>
&nbsp;�û�����<input maxLength=16 name=username size=12>�������룺<input maxLength=20 name=password size=12 type=password>����<select name=CookieDate><option selected value=0>������</option><option value=1>����һ��</option><option value=2>����һ��</option><option value=3>����һ��</option></select>����<input type=submit name=submit value="�� ½">
</td>
</tr>
</form>
<%end if%>
<TR>
<Th  width=30 height=25>״̬</Th>
<Th vAlign=center width=*>��̳����</Th>
<Th vAlign=center align=middle width=80>����</Th>
<Th vAlign=center noWrap align=middle width=38>���� </Th>
<Th vAlign=center noWrap align=middle width=38>���� </Th>
<Th vAlign=center noWrap align=middle width=168>��󷢱� </Th>
</TR>
<%
dim sql1,rs1
dim boardview
sql="select id,class from class order by orders"
set rs=conn.execute(sql)
do while not rs.eof
	response.write "<TR><TD colSpan=6 class=tabletitle2 height=25><B>"&rs(1)&"</B></TD></TR>"
	if boardmaster or master then
	sql1="select boardid,boardtype,class,readme,lastbbsnum,boardmaster,lockboard,lasttopicnum,indexIMG,LastPost from board "
	sql1=sql1&" where class="&rs(0)&" order by orders"
	else
	sql1="select boardid,boardtype,class,readme,lastbbsnum,boardmaster,lockboard,lasttopicnum,indexIMG,LastPost from board "
	sql1=sql1&" where class="&rs(0)&" and lockboard<>1 order by orders"
	end if
	set rs1=conn.execute(sql1)
	do while not rs1.eof
	LastPostInfo=split(rs1(9),"$")
	if not isdate(LastPostInfo(2)) then LastPostInfo(2)=Now()
	if rs1(5)<>"" then
		master_1=split(rs1(5), "|")
		for i = 0 to ubound(master_1)
		if i>2 then
			master_2=master_2
		else
			master_2=master_2 & "<a href=dispuser.asp?name="&master_1(i)&" target=_blank>"&master_1(i)&"</a><br>"
		end if
		next
		if i>3 then master_2=master_2 & "<font color=gray>More...</font>"
	else
		master_2="����"
	end if
'=========================BoardInfo============================
%>
<TR><TD width=26 class=tablebody2 valign=middle>
<%
	if rs1(6)=2 then
		if datediff("h",lastlogin,LastPostInfo(2))=0 then
		boardstat="<img src="&Forum_info(7)&Forum_pic(15)&" alt=��������>"
		else
		boardstat="<img src="&Forum_info(7)&Forum_pic(14)&" alt=��������>"
		end if
	else
		if datediff("h",lastlogin,LastPostInfo(2))=0 then
		boardstat="<img src="&Forum_info(7)&Forum_pic(7)&" alt=��������>"
		else
		boardstat="<img src="&Forum_info(7)&Forum_pic(6)&" alt=��������>"
		end if
	end if
	response.write boardstat
%>
</TD><TD vAlign=top width=* class=tablebody1>
<table width=100% cellpadding=2 cellspacing=0>
<tr><td class=tablebody1><a href="list.asp?boardid=<%=rs1(0)%>">
<font color=<%=Forum_body(14)%>><%=rs1(1)%></font></a>
</td></tr>
<tr><td class=tablebody1>
<%
if rs1(8)<>"" then
	response.write "<table align=left><tr><td><a href=list.asp?boardid="&rs1(0)&"><img src="&rs1(8)&" align=top border=0></a></td><td width=20></td></tr></table>"
end if
%>
<%=rs1(3)%></td></tr></table>
</TD><TD vAlign=center align=middle class=tablebody2 width=80>
<%=master_2%>
</TD>
<TD vAlign=center noWrap align=middle width=38 class=tablebody1><%=rs1(7)%></TD>
<TD vAlign=center noWrap align=middle width=38 class=tablebody1><%=rs1(4)%></TD>
<TD noWrap width=168 class=tablebody2>
���ߣ�<a href="dispuser.asp?id=<%=htmlencode(LastPostInfo(5))%>" target=_blank><%=htmlencode(LastPostInfo(0))%></a>��<a href="dispbbs.asp?Boardid=<%=rs1(0)%>&ID=<%=LastPostInfo(6)%>&replyID=<%=LastPostInfo(1)%>&skin=1"><IMG border=0 src="<%=Forum_info(7)%>lastpost.gif" title="ת����<%=htmlencode(LastPostInfo(3))%>"></a><br>
���⣺<a href="dispbbs.asp?Boardid=<%=rs1(0)%>&ID=<%=LastPostInfo(6)%>&replyID=<%=LastPostInfo(1)%>&skin=1"><%=htmlencode(left(LastPostInfo(3),10))%>..</a><br>
<%= FormatDateTime(LastPostInfo(2),1) %>��<%= FormatDateTime(LastPostInfo(2),4) %>
</TD></TR>
<%
'============================End===============================
	master_2=""
	rs1.movenext
	loop
rs.movenext
loop
set rs=nothing
set rs1=nothing
end sub
%>
</table>
<BR>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
<TR><Th  colSpan=2 align=left height=25>-=> ������̳</Th></TR>
<TR>
<TD vAlign=top class=tablebody1 width=100% >
<table width=100% >
<%
	dim linkinfo
	i=7
	sql="select boardname,readme,url from bbslink where islogo=0 order by id"
	set rs=conn.execute(sql)
	if not rs.eof and not rs.bof then
	do while not rs.eof
	if i>6 then 
	response.write "<tr><td width=""16%"">"
	i=1
	else
	response.write "<td width=""16%"">"
	end if
	response.write "<a href="&rs(2)&" target=_blank title="""&rs(1)&""">"&rs(0)&"</a>"
	rs.movenext
	response.write "</td>"
	if i=6 then response.write "</tr>"
	i=i+1
	loop
	end if
	i=7
	sql="select boardname,readme,url,logo from bbslink where islogo=1 order by id"
	set rs=conn.execute(sql)
	if not rs.eof and not rs.bof then
	response.write "<tr><td colspan=6><hr align=center size=1 color="&forum_body(27)&"></td></tr>"
	do while not rs.eof
	if i>6 then 
	response.write "<tr><td width=""16%"">"
	i=1
	else
	response.write "<td width=""16%"">"
	end if
	response.write "<a href="&rs(2)&" target=_blank><img src="""&rs(3)&""" border=0 alt="""&rs(0)&"��"&rs(1)&""" height=31 width=88></a>"
	rs.movenext
	response.write "</td>"
	if i=6 then response.write "</tr>"
	i=i+1
	loop
	end if
	set rs=nothing
%>
</table>
</TD>
</TR>
</table><BR>
<table cellpadding=5 cellspacing=1 class=tableborder1 align=center style="word-break:break-all;">
<TR><Th align=left colSpan=2 height=25>-=> �û�������Ϣ</Th></TR>
<TR>
<TD vAlign=top class=tablebody1 height=25 width=100% >
<%
	dim userip,userip2
	userip = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
	userip2 = Request.ServerVariables("REMOTE_ADDR")
	if userip = ""  then
		response.write "������ʵ�ɣ� �ǣ�"&userip2&"��"
	else
		response.write "������ʵ�ɣ� �ǣ�"&userip&"��"
	end if
	response.write usersysinfo(Request.ServerVariables("HTTP_USER_AGENT"),2)&"��"&usersysinfo(Request.ServerVariables("HTTP_USER_AGENT"),1)
%>
</TD></TR>

<TR><Th colSpan=2 align=left id=tabletitlelink height=25 style="font-weight:normal"><b>-=> ��̳����ͳ��</b>&nbsp;
<%
	if request("action")="show" then
		response.write "[<a href=index.asp?action=off>�ر���ϸ�б�</a>]"
	else
		if cint(Forum_Setting(14))=1 and request("action")<>"off" then
		response.write "[<a href=index.asp?action=off>�ر���ϸ�б�</a>]"
		else
		response.write "[<a href=index.asp?action=show>��ʾ��ϸ�б�</a>]"
		end if
	end if
%>
[<a href=boardstat.asp?reaction=online>�鿴�����û�λ��</a>]</Th></TR>
<TR>
<TD width=100% vAlign=top class=tablebody1>  Ŀǰ��̳���ܹ��� <b><%=clng(allonline())%></b> �����ߣ�����ע���Ա <b><%=onlinenum%></b> �ˣ��ÿ� <b><%=guestnum%></b> �ˡ�<br>
��ʷ������߼�¼�� <b><%=Maxonline%></b> ��ͬʱ���ߣ�����ʱ���ǣ�<%=formatdatetime(MaxonlineDate,1)%>&nbsp;<%=formatdatetime(MaxonlineDate,4)%>
<BR>
<font color=<%=Forum_body(8)%>>����ͼ��</font>��<img src=<%=Forum_info(7)&Forum_pic(0)%>> ��̳��   ����<img src=<%=Forum_info(7)&Forum_pic(1)%>> ��̳̳��   ��   <img src=<%=Forum_info(7)&Forum_pic(2)%>> ��̳���   ��   <img src=<%=Forum_info(7)&Forum_pic(3)%>> ��ͨ��Ա   ��   <img src=<%=Forum_info(7)&Forum_pic(4)%>> ���˻������Ա<hr size=1 color=<%=forum_body(27)%>>
<table width=100% border=0 cellspacing=0 cellpadding=0>
<%
	if request("action")="off" then
		call onlineuser(0,0,0)
	elseif request("action")="show" then
		call onlineuser(1,1,0)
	else
		call onlineuser(Forum_Setting(14),Forum_Setting(15),0)
	end if
%>
</table></TD></TR>
</TABLE><BR>
<table cellspacing=1 cellpadding=3 width="<%=Forum_body(12)%>" border=0 align=center>
<tr><td align=center><img src="<%=Forum_info(7)&Forum_pic(6)%>" align="absmiddle">&nbsp;û���µ�����&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src="<%=Forum_info(7)&Forum_pic(7)%>" align="absmiddle">&nbsp;���µ�����&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src="<%=Forum_info(7)&Forum_pic(14)%>" align="absmiddle">&nbsp;����������̳</td></tr>
</table>
<%
if instr(scriptname,"index.asp")>0 or instr(scriptname,"list.asp")>0 then
	if Forum_ads(2)=1 then
	call admove()
	end if
	if Forum_ads(13)=1 then
	call fixup()
	end if
end if
call footer()
%>
<!--#include file="online_l.asp"-->
<!--#include file="inc/ad_fixup.asp"-->