<%
Rem ==========��̳ͨ�ú���=========
Rem �ж������Ƿ�����
function isInteger(para)
       on error resume next
       dim str
       dim l,i
       if isNUll(para) then 
          isInteger=false
          exit function
       end if
       str=cstr(para)
       if trim(str)="" then
          isInteger=false
          exit function
       end if
       l=len(str)
       for i=1 to l
           if mid(str,i,1)>"9" or mid(str,i,1)<"0" then
              isInteger=false 
              exit function
           end if
       next
       isInteger=true
       if err.number<>0 then err.clear
end function

'ҳ�������ʾ��Ϣ
sub dvbbs_error()
%>
<br>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1 style="width:75%">
<tr align=center>
<th width="100%" height=25 colspan=2>��̳������Ϣ
</td>
</tr>
<tr>
<td width="100%" class=tablebody1 colspan=2>
<b>��������Ŀ���ԭ��</b><br><br>
<li>���Ƿ���ϸ�Ķ���<a href="boardhelp.asp?boardid=<%=boardid%>">�����ļ�</a>����������û�е�½���߲�����ʹ�õ�ǰ���ܵ�Ȩ�ޡ�
<%=errmsg%>
</td></tr>
<%if not founduser then%>
<form action="login.asp?action=chk" method=post>
    <tr>
    <th valign=middle colspan=2 align=center height=25>�����������û����������½</td></tr>
    <tr>
    <td valign=middle class=tablebody1>�����������û���</td>
    <td valign=middle class=tablebody1><INPUT name=username type=text> &nbsp; <a href=reg.asp>û��ע�᣿</a></td></tr>
    <tr>
    <td valign=middle class=tablebody1>��������������</font></td>
    <td valign=middle class=tablebody1><INPUT name=password type=password> &nbsp; <a href=lostpass.asp>�������룿</a></td></tr>
    <tr>
    <td class=tablebody1 valign=top width=30% ><b>Cookie ѡ��</b><BR> ��ѡ����� Cookie ����ʱ�䣬�´η��ʿ��Է������롣</td>
    <td valign=middle class=tablebody1>                <input type=radio name=CookieDate value=0 checked>�����棬�ر��������ʧЧ<br>
                <input type=radio name=CookieDate value=1>����һ��<br>
                <input type=radio name=CookieDate value=2>����һ��<br>
                <input type=radio name=CookieDate value=3>����һ��<br>                </td></tr>
    <tr>
    <td valign=top width=30% class=tablebody1><b>�����½</b><BR> ������ѡ�������½����̳��Ա�����û��б�����������Ϣ��</td>
    <td valign=middle class=tablebody1>                <input type=radio name=userhidden value=2 checked>������½<br>
                <input type=radio name=userhidden value=1>�����½<br>
                </td></tr>
	<input type=hidden name=comeurl value="<%=Request.ServerVariables("HTTP_REFERER")%>">
    <tr>
    <td class=tablebody2 valign=middle colspan=2 align=center><input type=submit name=submit value="�� ½">&nbsp;&nbsp;<input type=button name="back" value="�� ��" onclick="location.href='<%=Request.ServerVariables("HTTP_REFERER")%>'"></td></tr>
</form>
<%else%>
    <tr>
    <td class=tablebody2 valign=middle colspan=2 align=center><a href="<%=Request.ServerVariables("HTTP_REFERER")%>"><<������һҳ</a></td></tr>
<%end if%>
</table>
<%
end sub

sub dvbbs_suc()
%>
<br>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
<tr align=center>
<th width="100%">��̳�ɹ���Ϣ
</td>
</tr>
<tr>
<td width="100%" class=tablebody1>
<b>�����ɹ���</b><br><br>
<%=sucmsg%>
</td></tr>
<tr align=center><td width="100%" class=tablebody2>
<a href="<%=Request.ServerVariables("HTTP_REFERER")%>"> << ������һҳ</a>
</td></tr>
</table>
<%
end sub

rem �����ַ�
function ChkBadWords(fString)
    dim bwords,ii
    if not(isnull(BadWords) or isnull(fString)) then
    bwords = split(BadWords, "|")
    for ii = 0 to ubound(bwords)
        fString = Replace(fString, bwords(ii), string(len(bwords(ii)),"*")) 
    next
    ChkBadWords = fString
    end if
end function

Rem ����HTML����
function HTMLEncode(fString)
if not isnull(fString) then
    fString = replace(fString, ">", "&gt;")
    fString = replace(fString, "<", "&lt;")

    fString = Replace(fString, CHR(32), "&nbsp;")
    fString = Replace(fString, CHR(9), "&nbsp;")
    fString = Replace(fString, CHR(34), "&quot;")
    fString = Replace(fString, CHR(39), "&#39;")
    fString = Replace(fString, CHR(13), "")
    fString = Replace(fString, CHR(10) & CHR(10), "</P><P> ")
    fString = Replace(fString, CHR(10), "<BR> ")

    fString=ChkBadWords(fString)
    HTMLEncode = fString
end if
end function

Rem ���˱��ַ�
function HTMLcode(fString)
if not isnull(fString) then
    fString = Replace(fString, CHR(13), "")
    fString = Replace(fString, CHR(10) & CHR(10), "</P><P>")
    fString = Replace(fString, CHR(10), "<BR>")
    HTMLcode = fString
end if
end function

'�û�IP����
function LockIP(sip)
	dim str1,str2,str3,str4
	dim num
	LockIP=false
	if isnumeric(left(sip,2)) then
		str1=left(sip,instr(sip,".")-1)
		sip=mid(sip,instr(sip,".")+1)
		str2=left(sip,instr(sip,".")-1)
		sip=mid(sip,instr(sip,".")+1)
		str3=left(sip,instr(sip,".")-1)
		str4=mid(sip,instr(sip,".")+1)
		if isNumeric(str1)=0 or isNumeric(str2)=0 or isNumeric(str3)=0 or isNumeric(str4)=0 then
	
		else
			num=cint(str1)*256*256*256+cint(str2)*256*256+cint(str3)*256+cint(str4)-1
			sql="select count(*) from LockIP where ip1 <="&num&" and ip2 >="&num&""
			set rs=conn.execute(sql)
			if rs(0)>0 then 
				LockIP=true
			end if
			set rs=nothing
		end if
	end if
end function

Rem �жϷ����Ƿ������ⲿ
function ChkPost()
	dim server_v1,server_v2
	chkpost=false
	server_v1=Cstr(Request.ServerVariables("HTTP_REFERER"))
	server_v2=Cstr(Request.ServerVariables("SERVER_NAME"))
	if mid(server_v1,8,len(server_v2))<>server_v2 then
		chkpost=false
	else
		chkpost=true
	end if
end function

Rem �ж��û���Դ
function address(sip)
	dim str1,str2,str3,str4
	dim num
	dim country,city
	dim irs
	if isnumeric(left(sip,2)) then
	if sip="127.0.0.1" then sip="192.168.0.1"
	str1=left(sip,instr(sip,".")-1)
	sip=mid(sip,instr(sip,".")+1)
	str2=left(sip,instr(sip,".")-1)
	sip=mid(sip,instr(sip,".")+1)
	str3=left(sip,instr(sip,".")-1)
	str4=mid(sip,instr(sip,".")+1)
	if isNumeric(str1)=0 or isNumeric(str2)=0 or isNumeric(str3)=0 or isNumeric(str4)=0 then

	else
		num=cint(str1)*256*256*256+cint(str2)*256*256+cint(str3)*256+cint(str4)-1
		sql="select Top 1 country,city from address where ip1 <="&num&" and ip2 >="&num&""
		set irs=server.createobject("adodb.recordset")
		irs.open sql,conn,1,1
		if irs.eof and irs.bof then 
		country="����"
		city=""
		else
		country=irs(0)
		city=irs(1)
		end if
		irs.close
		set irs=nothing
	end if
	address=country&city
	else
	address="δ֪"
	end if
end function

function iif(expression,returntrue,returnfalse)
	if expression=0 then
		iif=returnfalse
	else
		iif=returntrue
	end if
end function

function iiif(express,expression,returntrue,returnfalse)
	if express>expression then
		iiif=returnfalse
	else
		iiif=returntrue
	end if
end function

function iimg(expression,returnfalse,returntrue)
	if expression="" or isnull(expression) then
		iimg=returnfalse
	else
		iimg=returntrue
	end if
end function

Rem ����SQL�Ƿ��ַ�
function checkStr(str)
	if isnull(str) then
		checkStr = ""
		exit function 
	end if
	checkStr=replace(str,"'","''")
end function

function chkboardmaster(boardid)
	dim brs,ckboardmaster
	chkboardmaster=false
	if master then
		chkboardmaster=true
	else
	if boardmaster then
	sql="select boardmaster,boardid from board where boardid="&boardid
	set brs=server.createobject("adodb.recordset")
	set brs=conn.execute(sql)
	if brs.eof and brs.bof then
		chkboardmaster=false
	else
		ckboardmaster=split(brs(0),"|")
		for i=0 to ubound(ckboardmaster)
			if trim(ckboardmaster(i))=trim(membername) then
			chkboardmaster=true
			else
			chkboardmaster=false
			end if
			if chkboardmaster then exit for
		next
	end if
	brs.close
	set brs=nothing
	else
	chkboardmaster=false
	end if
	end if
end function

Rem �û�����
sub activeonline()
dim ComeFrom,actCome,statuserid
statuserid=replace(Request.ServerVariables("REMOTE_HOST"),".","")
if not founduser then
	session("userid")=statuserid
	sql="select id,boardid from online where id="&cstr(session("userid"))
	set rs=conn.execute(sql)
	if rs.eof and rs.bof then
		ComeFrom=""
		actCome=""
		sql="insert into online(id,username,userclass,ip,startime,lastimebk,boardid,browser,stats,actforip,UserGroupID,actCome,userhidden) values ("&statuserid&",'����','����','"&Request.ServerVariables("REMOTE_HOST")&"',Now(),Now(),"&boardid&",'"&Request.ServerVariables("HTTP_USER_AGENT")&"','"&replace(stats,"'","")&"','"&Request.ServerVariables("HTTP_X_FORWARDED_FOR")&"',7,'"&actCome&"',"&userhidden&")"
	else
		sql="update online set lastimebk=Now(),boardid="&boardid&",stats='"&replace(stats,"'","")&"' where id="&cstr(session("userid"))
	end if
	conn.execute(sql)
else
	if founderr then
	boardid=0
	stats="������Ϣ"
	end if
	sql="select id,boardid from online where userid="&userid
	set rs=conn.execute(sql)
	if rs.eof and rs.bof then
	ComeFrom=""
	actCome=""
		sql="insert into online(id,username,userclass,ip,startime,lastimebk,boardid,browser,stats,actforip,UserGroupID,actCome,userhidden,userid) values ("&statuserid&",'"&membername&"','"&memberclass&"','"&Request.ServerVariables("REMOTE_HOST")&"',Now(),Now(),"&boardid&",'"&Request.ServerVariables("HTTP_USER_AGENT")&"','"&replace(stats,"'","")&"','"&Request.ServerVariables("HTTP_X_FORWARDED_FOR")&"',"&UserGroupID&",'"&actCome&"',"&userhidden&","&userid&")"
	else
		sql="update online set lastimebk=Now(),boardid="&boardid&",stats='"&replace(stats,"'","")&"' where userid="&userid
	end if
	conn.execute(sql)
	rs.close
	if session("userid")<>"" then
	Conn.Execute("delete from online where id="&session("userid"))
	session("userid")=""
	end if
end if
set rs=nothing
Rem ɾ����ʱ�û�
sql="Delete FROM online WHERE DATEDIFF('s', lastimebk, now()) > "&Forum_Setting(8)&"*60"
Conn.Execute sql
end sub

sub footer()
endtime=timer()
%>
<p>
<TABLE cellSpacing=0 cellPadding=0 width="<%=Forum_body(12)%>" border=0 align=center>
<tr><td align=center>
<%=Forum_ads(1)%>
</td></tr>
<tr><td align=center>
<%=Version%><br>
<%=Copyright%> , ҳ��ִ��ʱ�䣺<%=FormatNumber((endtime-startime)*1000,3)%>����
</td></tr>
</table>
</body>
</html>
<%
	CloseDatabase
end sub

sub nav()
%>
<html>
<head>
<META http-equiv=Content-Type content=text/html; charset=gb2312>
<meta name=keywords content="�����ȷ�,������̳,dvbbs">
<title><%=Forum_info(0)%>--<%=stats%></title>
<!--#include file="Forum_css.asp"-->
<!--#include file="Forum_js.asp"-->
</head>
<body <%=Forum_body(11)%>>
<table cellspacing=0 cellpadding=0 align=center style="border:1px <%=Forum_body(27)%> solid; border-top-width: 0px; border-right-width: 1px; border-bottom-width: 0px; border-left-width: 1px;width:<%=Forum_body(12)%>;">
<tr><td width=100%>
<table width=100% align=center border=0 cellspacing=0 cellpadding=0>
<tr><td class=TopDarkNav height=9></td></tr>
<tr><td height=70 class=TopLighNav2>
<TABLE border=0 width="100%" align=center>
<TR>
<TD align=left width="25%"><a href="<%= Forum_info(3) %>"><img border=0 src='<%= Forum_info(6) %>'></a></TD>
<TD Align=center width="65%">
<%if isnull(Forum_ads(0)) or Forum_ads(0)="" then%>
<object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0" width="468" height="60"><param name=movie value="http://www.dvbbs.net/skin/default/dvbanner.swf"><param name=quality value=high><param name=menu value=false><embed src="http://www.dvbbs.net/skin/default/dvbanner.swf" quality=high pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="468" height="60"></embed></object>
<%else%>
<%=Forum_ads(0)%>
<%end if%>
</td>
<td align=right style="line-height: 15pt" width="10%">
<a href=#><span style="CURSOR: hand" onClick="window.external.AddFavorite('<%=Forum_info(1)%>', '<%=Forum_info(0)%>')">�����ղ�</span></a>
<br><a href="mailto:<%=Forum_info(5)%>">��ϵ����</a>
<br><a href="boardhelp.asp">������·</a>
</td>
</td></tr>
</table>
</td></tr>
<tr><td class=TopLighNav height=9></td></tr>
        <tr> 
          <td class=TopLighNav1 height=22  valign="middle">&nbsp;&nbsp;
<%if not founduser then%>
<a href="login.asp">��½</a> <img src=<%=Forum_info(7)%>navspacer.gif align=absmiddle> <a href="reg.asp">ע��</a>
<%else%>
<%if userhidden=2 then%><a href="cookies.asp?action=hidden&userid=<%=userid%>">����</a><%else%><a href="cookies.asp?action=online&userid=<%=userid%>">����</a><%end if%> <img src=<%=Forum_info(7)%>navspacer.gif align=absmiddle> <a href="login.asp">�ص�½</a> <img src=<%=Forum_info(7)%>navspacer.gif align=absmiddle> <a href="usermanager.asp">�û��������</a>
<%end if%>
 <img src=<%=Forum_info(7)%>navspacer.gif align=absmiddle> <a href="query.asp?boardid=<%=boardid%>">����</a> <img src=<%=Forum_info(7)%>navspacer.gif align=absmiddle> <a href="boardstat.asp?boardid=<%=boardid%>">��̳״̬</a> <img src=<%=Forum_info(7)%>navspacer.gif align=absmiddle> <a href="boardhelp.asp?boardid=<%=boardid%>">����</a><%if founduser then%> <img src=<%=Forum_info(7)%>navspacer.gif align=absmiddle> <a href="logout.asp">�˳�</a><%else%> <img src=<%=Forum_info(7)%>navspacer.gif align=absmiddle> <a href="dispuser.asp?boardid=<%=boardid%>&action=permission">������ʲô</a><%end if%>
<%if master then response.write " <img src="&Forum_info(7)&"navspacer.gif align=absmiddle> <a href=admin_index.asp>����</a> <img src="&Forum_info(7)&"navspacer.gif align=absmiddle> <a href=""recycle.asp"">����վ</a>"%>
			</td>
        </tr>
</table>
</td></tr>
</table>
<%
if Cint(GroupSetting(0))=0 and (instr(scriptname,"reg.asp")=0 and instr(scriptname,"login.asp")=0) then
	Errmsg=Errmsg+"<br>"+"<li>��û���������̳��Ȩ�ޣ���<a href=login.asp>��½</a>����ͬ����Ա��ϵ��"
	call head_var(1)
	call dvbbs_error()
	call footer()
	response.end
end if
end sub

sub head_var(navline)
%>
<BR>
<table cellspacing=1 cellpadding=3 align=center border=0 width="<%=Forum_body(12)%>">
<tr>
<%if not founduser then%>
<td height=25>
>> <%if foundboard then%><%=BoardReadme%><%else%>��ӭ���� <B><%=Forum_info(0)%></B><%end if%>
<%else%>
<td width=65% >
>> ��ӭ <B><%=membername%></B>  [ <a href="JavaScript:openScript('messanger.asp?action=new',500,400)">������</a> :: <a href="topicwithme.asp?s=2">���������</a> :: <a href="topicwithme.asp?s=1">���������</a> :: <a href="dispuser.asp?id=<%=userid%>&boardid=<%=boardid%>&action=permission">������ʲô</a> ]
</td><td width=35% align=right>
<%if Cint(newincept())>Cint(0) then%>
<bgsound src="pic/mail.wav" border=0>
<img src=<%=Forum_info(7)%>msg_new_bar.gif> <a href="usersms.asp?action=inbox">�ҵ��ռ���</a> (<a href="javascript:openScript('messanger.asp?action=read&id=<%=inceptid(1)%>&sender=<%=inceptid(2)%>',500,400)"><font color="<%=Forum_body(8)%>"><%=newincept()%> ��</font></a>)
<%else%>
<img src=<%=Forum_info(7)%>msg_no_new_bar.gif> <a href="usersms.asp?action=inbox">�ҵ��ռ���</a> (<font color=gray>0 ��</font>)
<%end if%>
<%end if%>
</td></tr>
</table>
<table cellspacing=1 cellpadding=3 align=center class=tableBorder2>
<tr><td height=25 valign=middle>
<img src="<%=Forum_info(7)%>Forum_nav.gif" align=absmiddle> <a href=index.asp><%=Forum_info(0)%></a> �� 
<%
if navline=1 then
	response.write stats
else
	response.write "<a href=list.asp?boardid="&boardid&">"&boardtype&"</a> ��  "&htmlencode(replace(stats,boardtype,""))
end if
%>
<a name=top></a>
</td></td>
</table>
<br>
<%
end sub

'ͳ������
function newincept()
rs=conn.execute("Select Count(id) From Message Where flag=0 and issend=1 and delR=0 And incept='"& membername &"'")
    newincept=rs(0)
	set rs=nothing
	if isnull(newincept) then newincept=0
end function
function inceptid(stype)
	set rs=conn.execute("Select top 1 id,sender From Message Where flag=0 and issend=1 and delR=0 And incept='"& membername &"'")
	if stype=1 then
	inceptid=rs(0)
	else
	inceptid=rs(1)
	end if
	set rs=nothing
end function

Rem ��ð����û���Ȩ������
function GetBoardPermission()
	dim pmrs
	GetBoardPermission=false
	set pmrs=conn.execute("select PSetting from BoardPermission where Boardid="&Boardid&" and GroupID="&UserGroupID)
	if not (pmrs.eof and pmrs.bof) then
		GetBoardPermission=true
		GroupSetting=split(pmrs(0),",")
	else
		GetBoardPermission=false
	end if
	if FoundUser then
	set pmrs=conn.execute("select uc_Setting from UserAccess where uc_BoardID="&BoardID&" and uc_UserID="&userID)
	if not(pmrs.eof and pmrs.bof) then
		UserPermission=split(pmrs(0),",")
		GroupSetting=split(pmrs(0),",")
		FoundUserPer=true
	end if
	end if
	set pmrs=nothing
end function
%>
