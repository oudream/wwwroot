<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<%
'=========================================================
' File: toplist.asp
' Version:5.0
' Date: 2002-9-7
' Script Written by satan
'=========================================================
' Copyright (C) 2001,2002 AspSky.Net. All rights reserved.
' Web: http://www.aspsky.net,http://www.dvbbs.net
' Email: info@aspsky.net,eway@aspsky.net
'=========================================================

dim orders,ordername
dim currentpage,page_count,Pcount
dim totalrec,endpage
dim bbsnum,usernum
dim select1,select2,select3,select4,select5,select6,select7,select8

if Cint(GroupSetting(1))=0 then
	Errmsg=Errmsg+"<br>"+"<li>��û���������̳��Ա���ϵ�Ȩ�ޣ���<a href=login.asp>��½</a>����ͬ����Ա��ϵ��"
	founderr=true
end if

if founderr then
	call nav()
	call head_var(1)
	call dvbbs_error()
else
	currentPage=request("page")
	if not isInteger(request("orders")) or request("orders")="" then
		orders=1
	else
		orders=request("orders")
	end if
	select case orders
	case 1
		orders=1
		ordername="��������Top" & Forum_Setting(68)
		select1="selected"
	case 2
		orders=2
		ordername="�����û�ע��"
		select2="selected"
	case 3
		orders=3
		ordername=Forum_Setting(68) & "����"
		select3="selected"
	case 7
		orders=7
		ordername="�����û��б�"
		select7="selected"
	case 8
		orders=8
		ordername="�����Ŷ�"
		select8="selected"
	case else
		orders=1
		ordername="��������Top" & Forum_Setting(68)
		select1="selected"
	end select
	
	stats=ordername
	call nav()
	call head_var(1)
	call main()
end if
call footer()

sub main()

	set rs=conn.execute("select top 1 BbsNum,UserNum from config where active=1")
	bbsnum=rs(0)
	usernum=rs(1)
%>
<table cellpadding=6 cellspacing=1 align=center class=tableborder1>
<form method=POST action=toplist.asp>
<tr>
<td colspan=5 valign=top width=350 class=tabletitle2>&nbsp;>> <B><%=ordername%></B> <<<BR><BR>
&nbsp;��ע���û����� <%=usernum%> �� &nbsp; ���������� <%=bbsnum%> ƪ</td>
<td colspan=6 align=right class=tabletitle2>
<select name=orders onchange='javascript:submit()'>
<option value=1 <%=select1%>>��������Top<%=Forum_Setting(68)%></option>
<option value=2 <%=select2%>>����ע���û�</option>
<option value=3 <%=select3%>><%=Forum_Setting(68)%>����</option>
<option value=7 <%=select7%>>�����û��б�</option>
<option value=8 <%=select8%>>�����Ŷ�</option>
</select>
</td></tr></form>
<tr> 
<th>�û���</th>
<th>Email</th>
<th>OICQ</th>
<th>��ҳ</th>
<th>����Ϣ</th>
<th>ע��ʱ��</th>
<th>�ȼ�״̬</th>
<th>��������</th>
<th>�Ʋ�</th></tr>
<%
	if currentpage="" or not isInteger(currentpage) then
		currentpage=1
	else
		currentpage=clng(currentpage)
		if err then
			currentpage=1
			err.clear
		end if
	end if
	set rs=server.createobject("adodb.recordset")
	select case orders
	case 1
		sql="select top "&Forum_Setting(68)&" username,useremail,userclass,oicq,homepage,article,addDate,userwealth as wealth,userid from [user] order by article desc"
	case 2
		sql="select top "&Forum_Setting(68)&" username,useremail,userclass,oicq,homepage,article,addDate,userwealth as wealth,userid from [user] order by AddDate desc"
	case 3
		sql="select top "&Forum_Setting(68)&" username,useremail,userclass,oicq,homepage,article,addDate,userwealth as wealth,userid from [user] order by userwealth desc"
	case 7
		sql="select username,useremail,userclass,oicq,homepage,article,addDate,userwealth as wealth,userid from [user] order by userid desc"
	case 8
		sql="select username,useremail,userclass,oicq,homepage,article,addDate,userwealth as wealth,userid from [user] where usergroupid<=3 order by usergroupid,article desc"
	case else
		sql="select top "&Forum_Setting(68)&" username,useremail,userclass,oicq,homepage,article,addDate,userwealth as wealth,userid from [user] order by article desc"
	end select
	rs.open sql,conn,1
	if rs.eof and rs.bof then
		response.write "<tr><td colspan=9 class=tablebody1>������û���κ��û����ݡ�</font></td></tr>"
	else
	if orders=7 then
	totalrec=userNum
	else
	totalrec=rs.recordcount
	end if
  	if totalrec mod Forum_Setting(11)=0 then
     		Pcount= totalrec \ Forum_Setting(11)
  	else
     		Pcount= totalrec \ Forum_Setting(11)+1
  	end if
	RS.MoveFirst
	if currentpage > Pcount then currentpage = Pcount
   	if currentpage<1 then currentpage=1
	RS.Move (currentpage-1) * Forum_Setting(11)
	page_count=0
	do while not rs.eof and page_count < Clng(Forum_Setting(11))
%>
<tr bgcolor=<%=Forum_body(4)%>>
<td class=tablebody1>&nbsp;<a href=dispuser.asp?id=<%=rs("userid")%> target=_blank><%=rs("username")%></a></td>
<td align=center class=tablebody2><a href=mailto:<%=rs("useremail")%>><img border=0 src=<%=Forum_info(7)&Forum_TopicPic(10)%>></a></td>
<td align=center class=tablebody1>
<%=iimg(rs("oicq"),"û��","<a href=http://search.tencent.com/cgi-bin/friend/user_show_info?ln="&rs("oicq")&" target=_blank><img src="&Forum_info(7)&Forum_TopicPic(11)&" alt='�鿴 OICQ:"&rs("oicq")&" ������' border=0 ></a>")%>
</td>
<td align=center class=tablebody2>
<%=iimg(rs("homepage"),"û��","<a href="&rs("homepage")&" target=_blank><img border=0 src="&Forum_info(7)&Forum_TopicPic(14)&"></a>")%>
</td>
<td align=center class=tablebody1><a href=javascript:openScript('messanger.asp?action=new&touser=<%=htmlencode(rs("username"))%>',500,400)><img src=<%=Forum_info(7)&Forum_TopicPic(7)%> border=0></a></td>
<td align=center class=tablebody2><%=rs("addDate")%></td>
<td align=center class=tablebody1> <%=rs("userclass")%><br>
</td>
<td align=center class=tablebody2><%=rs("article")%></td>
<td align=center class=tablebody1><%=rs("wealth")%></td>
</tr>
<%
		page_count=page_count+1
		rs.movenext
	loop
	end if
%>
</table>
<table border=0 cellpadding=0 cellspacing=3 width="<%=Forum_body(12)%>" align=center>
<tr><td valign=middle nowrap>
<font color=<%=Forum_body(13)%>>ҳ�Σ�<b><%=currentpage%></b>/<b><%=Pcount%></b>ҳ
&nbsp;
ÿҳ<b><%=Forum_Setting(11)%></b> ����<b><%=totalrec%></b></font></td>
<td valign=middle nowrap><font color=<%=Forum_body(13)%>><div align=right><p>��ҳ�� 
<%
	if currentpage > 4 then
	response.write "<a href=""?page=1&orders="&orders&""">[1]</a> ..."
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
			response.write " <a href=""?page="&i&"&orders="&orders&""">["&i&"]</a>"
		end if
	end if
	next
	if currentpage+3 < Pcount then
		response.write "... <a href=""?page="&Pcount&"&orders="&orders&""">["&Pcount&"]</a>"
	end if
%>
</p></div></font></td></tr></table>
<%
	rs.close
	set rs=nothing
	call activeonline()
end sub
%>