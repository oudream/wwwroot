<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/char_board.asp" -->
<%
'=========================================================
' File: elist.asp
' Version:5.0
' Date: 2002-9-10
' Script Written by satan
'=========================================================
' Copyright (C) 2001,2002 AspSky.Net. All rights reserved.
' Web: http://www.aspsky.net,http://www.dvbbs.net
' Email: info@aspsky.net,eway@aspsky.net
'=========================================================

	dim Ers,Esql
	dim totalrec
	dim currentpage,page_count,Pcount
	dim guests
	dim topic
	dim rs1,sql1
	dim endpage

	if boardmaster or master then
		guestlist=""
	else
		guestlist=" lockboard<>1 and "
	end if
	stats="��̳����"
	currentPage=request("page")
	if BoardID="" or not isInteger(BoardID) or BoardID=0 then
		Errmsg=Errmsg+"<br>"+"<li>����İ����������ȷ�����Ǵ���Ч�����ӽ��롣"
		founderr=true
	else
		BoardID=clng(BoardID)
	end if
	if currentpage="" then
		currentpage=1
	elseif not isInteger(currentpage) then
		currentpage=1
	else
		currentpage=clng(currentpage)
	end if

	if founderr then
		call nav()
		call head_var(1)
	else
		call nav()
		call head_var(2)
		call main()
		if founderr then call dvbbs_error()
	end if
	call footer()

sub main()
	if Cint(GroupSetting(41))=0 then
		Errmsg=Errmsg+"<br>"+"<li>��û���������̳�������ӵ�Ȩ�ޣ���<a href=login.asp>��½</a>����ͬ����Ա��ϵ��"
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
	call activeonline()
	dim onlinenum,guestnum
	onlinenum=online(boardid)
	guestnum=guest(boardid)
	if trim(boardmasterlist)<>"" then
		master_1=split(boardmasterlist, "|")
		for i = 0 to ubound(master_1)
		master_2=""+master_2+"<a href=""dispuser.asp?name="+master_1(i)+""" target=_blank title=����鿴�ð�������>"+master_1(i)+"</a>&nbsp;"
		next
	else
		master_2="�ް���"
	end if
%>
<TABLE cellSpacing=0 cellPadding=0 width=<%=Forum_body(12)%> border=0 align=center>
<tr>
<td align=center width=100% valign=middle>
<SCRIPT LANGUAGE='JavaScript' SRC='js/fader.js' TYPE='text/javascript'></script>
<SCRIPT LANGUAGE='JavaScript' TYPE='text/javascript'>
prefix="";
arNews = ["<%
sql="select top 1 title,addtime from bbsnews where boardid="&boardid&"  order by id desc"
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
<BR>
<TABLE cellpadding=3 cellspacing=1 class=tableborder1 align=center>
<TBODY>
<TR>
<Th height=25 width="100%" align=left id=tabletitlelink style="font-weight:normal">������<b><%=allonline()%></b>�ˣ�����<%= boardtype %>�Ϲ��� <b><%= onlinenum %></b> λ��Ա�� <b><%= guestnum %></b> λ����.�������� <b><font color=<%=Forum_body(8)%>><%= todaynum %></font></b> 
<%
	if request("action")="show" then
		response.write "[<a href=elist.asp?action=off&boardID="& boardid&"&amp;page=1&skin="& skin &">�ر���ϸ�б�</a>]"
	else
		if cint(Forum_Setting(14))=1 and request("action")<>"off" then
		response.write "[<a href=elist.asp?action=off&boardID="& boardid&"&amp;page=1&skin="& skin &">�ر���ϸ�б�</a>]"
		else
		response.write "[<a href=elist.asp?action=show&boardID="& boardid&"&amp;page=1&skin="& skin &">��ʾ��ϸ�б�</a>]"
		end if
	end if
%>
</Th></TR>
<%
	if request("action")="off" then
		call onlineuser(0,0,boardid)
	elseif request("action")="show" then
		call onlineuser(1,1,boardid)
	else
		call onlineuser(Forum_Setting(14),Forum_Setting(15),boardid)
	end if
%>
</td></tr></TBODY></TABLE>
<BR>
<table cellpadding=0 cellspacing=0 border=0 width=<%=Forum_body(12)%> align=center valign=middle><tr>
<td align=center width=2> </td>
<td align=left> <a href="announce.asp?boardid=<%= boardid %>"><img src="<%=Forum_info(7)&Forum_boardpic(1)%>" border=0 alt="������"></a>
&nbsp;&nbsp;<a href="vote.asp?boardid=<%=boardid%>"><img src="<%=Forum_info(7)&Forum_boardpic(2)%>" border=0 alt="������ͶƱ"></a>
&nbsp;&nbsp;<a href="smallpaper.asp?boardid=<%=boardid%>"><img src="<%=Forum_info(7)&Forum_boardpic(3)%>" border=0 alt="����С�ֱ�"></a></td>
<td align=right><img src="<%=Forum_info(7)%>team2.gif" align=absmiddle>  <%= master_2 %></td></tr></table>
<table cellpadding=3 cellspacing=1 class=tableborder1 align=center>
<TR valign=middle>
<Th height=27 width=32>״̬</Th>
<Th width=*>�� ��</Th>
<Th width=80>�� ��</Th>
<Th width=195>������ | �ظ���</Th>
</TR>
<%
	sql="select * from BestTopic where boardID="&cstr(boardID)&" ORDER BY id desc"
	'response.write sql
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,1
	if rs.bof and rs.eof then
	response.write "<tr><td colSpan=4 width=100% class=tablebody1>����̳���޾�������</td></tr>"
	else
		rs.PageSize = Forum_Setting(11)
		rs.AbsolutePage=currentpage
		page_count=0
      	totalrec=rs.recordcount
		while (not rs.eof) and (not page_count = rs.PageSize)
'========================Best Topic Info=======================
response.write "<TR valign=middle>"
response.write "<TD class=tablebody2 width=32 height=27 align=center>"
response.write "<img src="""&Forum_info(7)&Forum_statePic(4)&""" alt=��������>"
response.write "</TD>"
response.write "<TD align=left class=tablebody1 width=*>"
response.write "  <img src='"& Forum_info(8) & rs("Expression") &"' width=15 height=15>  "
response.write "<a href=""dispbbs.asp?boardID="& boardID &"&ID="& rs("rootid") &"&replyID="& rs("announceid")&"&skin=1"">"
if len(rs("title"))>26 then
	response.write htmlencode(left(rs("title"),26))&"..."
else
	response.write htmlencode(rs("title"))
end if
response.write "</a></TD>"
response.write "<TD class=tablebody2 width=80 align=center><a href=dispuser.asp?id="& rs("postuserid") &" target=_blank>"& rs("postusername") &"</a></TD>"
response.write "<TD align=left class=tablebody1 width=195>&nbsp;&nbsp;"&FormatDateTime(rs("dateandtime"),2)&"&nbsp;"&FormatDateTime(rs("dateandtime"),4)&"&nbsp;<font color="&Forum_body(8)&">|</font>&nbsp;------"
response.write "</TD></TR>"
'============================End=============================
		page_count = page_count + 1
		rs.movenext
		wend
	end if
  	if totalrec mod Forum_Setting(11)=0 then
     		Pcount= totalrec \ Forum_Setting(11)
  	else
     		Pcount= totalrec \ Forum_Setting(11)+1
  	end if
%>
</table>
<table border=0 cellpadding=0 cellspacing=3 width="<%=Forum_body(12)%>" align=center>
<tr><td valign=middle nowrap>
ҳ�Σ�<b><%=currentpage%></b>/<b><%=Pcount%></b>ҳ
&nbsp;
ÿҳ<b><%=Forum_Setting(11)%></b> ����<b><%=totalrec%></b></td>
<td valign=middle nowrap align=right>��ҳ�� 
<%
	if currentpage > 4 then
	response.write "<a href=""?page=1&boardid="&boardid&""">[1]</a> ..."
	end if
	if Pcount>currentpage+3 then
	endpage=currentpage+3
	else
	endpage=Pcount
	end if
	for i=currentpage-3 to endpage
	if not i<1 then
		if i = clng(currentpage) then
        response.write " <font color="&Forum_body(8)&">["&i&"]</font>"
		else
        response.write " <a href=""?page="&i&"&boardid="&boardid&""">["&i&"]</a>"
		end if
	end if
	next
	if currentpage+3 < Pcount then 
	response.write "... <a href=""?page="&Pcount&"&boardid="&boardid&""">["&Pcount&"]</a>"
	end if

	response.write "</td></tr></table>"
	rs.close
	set rs=nothing
%>
<table border=0 cellpadding=0 cellspacing=3 width="<%=Forum_body(12)%>" align=center>
<tr>
<td valign=middle nowrap> <div align=right>
<select onchange="if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}">
<option selected>��ת��̳��...</option>
<%
	sql="select id,class from class order by orders"
	set rs=conn.execute(sql)
	if not(rs.eof and rs.bof) then
	do while not rs.eof
	response.write "<option>�� "& rs(1) &"</option>"

		sql1="select boardid,boardtype from board where "&guestlist&" class="& rs(0)&" order by orders"
		set rs1=conn.execute(sql1)
		if rs1.eof and rs1.bof then
			response.write "<option>û����̳</option>"
		else
			do while not rs1.eof
				response.write "<option value=""elist.asp?boardid="&rs1(0)&""">����" & rs1(1) & "</option>"
			rs1.movenext
			loop
		end if
	rs.movenext
	loop
	end if
	set rs=nothing
	set rs1=nothing
%>
</select><div></td></tr></table>
<BR>
<table cellspacing=1 cellpadding=3 width=100% class=tableborder1 align=center><tr>
<th width=80% align=left>��-=> <%= Forum_info(0) %>ͼ��</th>
<th noWrap width=20% align=right>����ʱ���Ϊ - <%=Forum_info(9)%> &nbsp;</th>
</tr>
<tr><td colspan=2 class=tablebody1>
<table cellspacing=4 cellpadding=0 width=92% border=0 align=center>
<tr><td><img src=<%=Forum_info(7)&Forum_statePic(0)%>> ���ŵ�����</td>
<td><img src=<%=Forum_info(7)&Forum_statePic(1)%>> �ظ�����10��</td>
<td><img src=<%=Forum_info(7)&Forum_statePic(2)%>> ����������</td>
<td><img src=<%=Forum_info(7)&Forum_statePic(3)%>> �̶����˵����� </td>
<td> <img src=<%=Forum_info(7)&Forum_statePic(4)%>> �������� </td>
<td> <img src=<%=Forum_info(7)&Forum_statePic(12)%>> ͶƱ���� </td>
</tr>
</table>
</td></tr></table>
<%
end sub
%>
<!--#include file="online_l.asp"-->