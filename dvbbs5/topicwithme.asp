<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<%
	dim totalrec
	dim n
	dim orders,ordername
	dim currentpage,page_count,Pcount
	currentPage=request("page")
	if not founduser then
		Errmsg="����û�е�½����<a href=login.asp>��½</a>�����"
		founderr=true
	end if
	orders=request("s")
	if orders="" or not isInteger(orders) then
		orders=1
	else
		orders=clng(orders)
	end if
	if orders=1 then
	ordername="�Ҳ��������"
	sql="select top 200 * from topic where topicid in (select top 200 rootid from "&NowUseBBS&" where postuserid="&userid&" order by announceid desc) order by topicid desc"
	elseif orders=2 then
	sql="select top 200 topicID,boardID,postUserName,child,title,DateAndTime,hits,Expression,locktopic,istop,isbest,isvote from topic where postuserid="&userid&" and child>0 ORDER BY topicid desc"
	ordername="�ѱ��ظ����ҵķ���"
	else
	ordername="�Ҳ��������"
	sql="select top 200 * from topic where topicid in (select top 200 rootid from "&NowUseBBS&" where postuserid="&userid&" order by announceid desc) order by topicid desc"
	end if	
	if currentpage="" or not isInteger(currentpage) then
		currentpage=1
	else
		currentpage=clng(currentpage)
	end if
	stats=ordername
if founderr then
	call nav()
	call head_var(1)
else
	call nav()
	call head_var(1)
	call AnnounceList1()
	call listPages3()
end if
call activeonline()
call footer()

sub AnnounceList1()
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,1
	if rs.bof and rs.eof then
		call showEmptyBoard1()
	else
		rs.PageSize = Forum_Setting(11)
		rs.AbsolutePage=currentpage
		page_count=0
		totalrec=rs.recordcount
		call showpagelist1()
	end if
	end sub

	REM ��ʾ�����б�	
	sub showPageList1()
	i=0
	response.write "<p align=center>��ǰ��̳���ݺܿ��ܽ����˷ֱ�������<BR>��ǰ����ʾ���ڵ�ǰ���ݱ�����ǰ200�����������Ҫ��ѯ���������뵽����ҳ��</p>"
	response.write "<TABLE cellPadding=1 cellSpacing=1 class=tableborder1 align=center>"
	response.write "<TR align=middle>"
	response.write "<Th height=27 width=32 id=tabletitlelink>״̬</th>"
	response.write "<Th width=*>�� ��  (��<img src="&Forum_info(7)&"plus.gif>����չ�������б�)</Th>"
	response.write "<Th width=80>�� ��</Th>"
	response.write "<Th width=64>�ظ�/����</Th>"
	response.write "<Th width=195>������ | �ظ���</Th>"
	response.write "</TR>"
	dim Ers, Eusername, Edateandtime,body
	dim vrs,votenum,pnum,p,iu,votenum_1
       while (not rs.eof) and (not page_count = rs.PageSize)
	boardid=rs("boardid")
response.write "<TR align=middle>"&_
				"<TD class=tablebody2 width=32 height=27>"

if rs("istop")<>1 and lockboard<>1 and rs("locktopic")<>1 and rs("isvote")<>1 and rs("isbest")<>1 and rs("child")<10 then
	response.write "<img src="""&Forum_info(7)&Forum_statePic(0)&""" alt=��������>"
elseif rs("isvote")=1 then
	response.write "<img src="""&Forum_info(7)&Forum_statePic(12)&""" alt=ͶƱ����>"
elseif rs("istop")=1 then
	response.write "<img src="""&Forum_info(7)&Forum_statePic(3)&""" alt=�̶�����>"
elseif rs("isbest")=1 then
	response.write "<img src="""&Forum_info(7)&Forum_statePic(4)&""" alt=��������>"
elseif rs("child")>=10 then
	response.write "<img src="""&Forum_info(7)&Forum_statePic(1)&""" alt=��������>"
elseif rs("locktopic")=1 then
	response.write "<img src="""&Forum_info(7)&Forum_statePic(2)&""" alt=������������>"
elseif lockboard=1 then
	response.write "<img src="""&Forum_info(7)&Forum_statePic(2)&""" alt=����̳������>"
else
	response.write "<img src="""&Forum_info(7)&Forum_statePic(0)&""" alt=��������>"
end if

response.write "</TD>"&_
				"<TD align=left  class=tablebody1 width=*>"

	response.write "<img src='"& Forum_info(7) &"nofollow.gif' id='followImg"& rs("topicid") &"'>"

response.write "<a href=""dispbbs.asp?boardID="& rs("boardid") &"&ID="& rs("topicid") &""">"

if len(rs("title"))>26 then
	response.write ""&left(htmlencode(rs("title")),26)&"..."
else
	response.write htmlencode(rs("title"))
end if
	response.write "</a>"
if (rs("child")+1)>cint(Forum_Setting(12)) then
	response.write "&nbsp;&nbsp;[��ҳ��"
	Pnum=(Cint(rs("child")+1)/cint(Forum_Setting(12)))+1
	for p=1 to Pnum
	response.write " <a href='dispbbs.asp?boardID="& boardID &"&ID="& rs("topicid") &"&star="&P&"'><FONT color="&Forum_body(8)&"><b>"&p&"</b></font></a> "
	next
	response.write "]"
end if

response.write "</TD>"&_
				"<TD class=tablebody2 width=80><a href=dispuser.asp?name="& rs("postusername") &">"& rs("postusername") &"</a></TD>"&_
				"<TD class=tablebody1 width=64>"
	response.write ""& rs("child") &"/"& rs("hits") &""

response.write "</TD><TD align=left  class=tablebody2 width=195>&nbsp;"

	response.write "&nbsp;"&_
						FormatDateTime(rs("dateandtime"),2)&"&nbsp;"&FormatDateTime(rs("dateandtime"),4)&_
						"&nbsp;<font color="&Forum_body(8)&">|</font>&nbsp;------"

response.write "</TD></TR>"

	  page_count = page_count + 1
          rs.movenext
        wend
	response.write "</table>"
	end sub

	sub listPages3()
	dim endpage
	Pcount=rs.PageCount
	response.write "<table border=0 cellpadding=0 cellspacing=3 width="""&Forum_body(12)&""" align=center>"&_
			"<tr><td valign=middle nowrap>"&_
			"<font color="&Forum_body(13)&">ҳ�Σ�<b>"&currentpage&"</b>/<b>"&Pcount&"</b>ҳ"&_
			"ÿҳ<strong>"&Forum_Setting(11)&"</b> ����<b>"&totalrec&"</b></td>"&_
			"<td valign=middle nowrap><div align=right><p>��ҳ�� "

	if currentpage > 4 then
	response.write "<a href=""?page=1&s="&orders&""">[1]</a> ..."
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
        response.write " <a href=""?page="&i&"&s="&orders&""">["&i&"]</a>"
		end if
	end if
	next
	if currentpage+3 < Pcount then 
	response.write "... <a href=""?page="&Pcount&"&s="&orders&""">["&Pcount&"]</a>"
	end if
	response.write "</p></div></font></td></tr></table>"
	rs.close
	set rs=nothing	
	end sub 

	sub showEmptyBoard1()
%>
<TABLE cellPadding=4 cellSpacing=1 class=tableborder1 align=center>
<TBODY>
<TR align=middle>
<Th height=25>״̬</Th>
<Th>�� ��  (�������Ϊ���´����)</Th>
<Th>�� ��</Th> 
<Th>�ظ�/����</Th> 
<Th>���»ظ�</Th>
</TR> 
<tr><td colSpan=5 width=100% class=tablebody1>�������ݣ���ӭ��������</td></tr>
</TBODY></TABLE>
<%
	end sub
%>