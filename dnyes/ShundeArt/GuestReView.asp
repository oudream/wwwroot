<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp" -->
<!--#include file="ChkManage.asp" -->
<%
if not(request.cookies(Forcast_SN)("ManageKEY")="super" or request.cookies(Forcast_SN)("ManageKEY")="typemaster" or request.cookies(Forcast_SN)("ManageKEY")="bigmaster") then
	Show_Err("�Բ������ĺ�̨����Ȩ�޲���������")
	response.end
else
	PageShowSize = 10            'ÿҳ��ʾ���ٸ�ҳ
	MyPageSize   = 16           'ÿҳ��ʾ����������
		
	If Not IsNumeric(Request("page")) Or IsEmpty(Request("page")) Or Request("page") <=0 Then
		MyPage=1
	Else
		MyPage=Int(Abs(Request("page")))
	End if
	
	set rs=server.CreateObject("ADODB.RecordSet")
	if reviewcheckshow="1" then
		rs.Source="select * from "& db_Review_Table &" where newsid is not null and newsid < 1 order by reviewID desc"
	else
		rs.Source="select * from "& db_Review_Table &" where passed=0 and newsid is not null newsid < 1 order by reviewID desc"
	end if
	
	rs.Open rs.Source,conn,1,1
	If Not rs.eof then
		rs.PageSize     = MyPageSize
		MaxPages        = rs.PageCount
		rs.absolutepage = MyPage
		total           = rs.RecordCount
		%>
<html><head>
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=site.css rel=stylesheet>
<title><%=copyright%><%=version%>&nbsp;<%=ver%> - �������</title>
</head>
<body topmargin="0">
<!--#include file=Admin_Top.asp-->
<div align="center">
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#C0C0C0" width="100%" id="AutoNumber1">
<tr class="TDtop">
	<td colspan="8" height=25 class="TDtop" align="center">�� <B>���Թ���</B> �� </td>
</tr>
<tr>
	<td colspan=8  align=center height=25>�� <%=total%> ������ǰ�� <%=Mypage%>/<%=Maxpages%> ҳ��ÿҳ <%=MyPageSize%> ��</td>
</tr>
<tr class="TDtop1" align="center">
	<td width="5%">ID��</td>
	<td width="7%">����</td>
	<td width="33%">������Ҫ</td>
	<td width="12%">����</td>
	<td width="22%">����</td>
	<td width="6%">����</td>		
	<td width="15%">ִ��</td>
</tr>
		<%
		for i=1 to rs.PageSize
			if not rs.EOF then
				content=trim(rs("content"))
				DisplayContent=mid(Content,1,18)
				%>
<form method="post" action="CheckReView.asp">
<tr bgcolor=ffffff>
	<td align=center ><%=rs("ReViewID")%></td>
	<td align=center><%if rs("NewsID")=0 then%>��ͨ<%else%><font color=red>���Ļ�</font><%end if%></td>
	<td align="left"><a href='GuestBook.asp?ReViewID=<%=rs("ReViewID")%>' target=_blank><%=htmlencode(DisplayContent)%>...</a></td>
	<td align=center><%=trim(rs("author"))%>��</td>
	<td align=center><a href='mailto:<%=trim(rs("email"))%>'><%=trim(rs("email"))%></a></td>
	<td align=center><%=Month(rs("UpdateTime")) %>-<%=Day(rs("UpdateTime")) %></td>
	<td align=center >
		<input type=hidden name="ReviewID" value="<%=rs("ReviewID")%>">
		<input type=hidden name="Guest" value=1>
		<input type=hidden name="passed" value="<%=rs("passed")%>">
		<input type=submit value="<%if rs("passed")=0 then%>ͨ��<%else%>ȡ��<%end if%>���" class=button target=_blank title="�������ť<%if rs("passed")=0 then%>ͨ��<%else%>ȡ��<%end if%>���" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
		<a href=DelReView2_alert.asp?ReViewID=<%=rs("ReViewID")%> style="font-size: 9pt;  color: #000000; background-color: #ffffff; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#EAEAF4'" onMouseOut ="this.style.backgroundColor='#ffffff'">ɾ��</a>
	</td>
</tr>
</form>
				<%
				rs.MoveNext
			end if
		next
		%>
<tr>
	<td colspan=8  align=center height=25>�� <%=total%> ������ǰ�� <%=Mypage%>/<%=Maxpages%> ҳ��ÿҳ <%=MyPageSize%> �� 
		<%
		url="ReViewManage.asp?" 
		PageNextSize=int((MyPage-1)/PageShowSize)+1
		Pagetpage=int((total-1)/rs.PageSize)+1
		if PageNextSize >1 then
			PagePrev=PageShowSize*(PageNextSize-1)
			Response.write "<a class=black href='" & Url & "page=" & PagePrev & "' title='��" & PageShowSize & "ҳ'>��һ��ҳ</a> "
			Response.write "<a class=black href='" & Url & "page=1' title='��1ҳ'>ҳ��</a> "
		end if
		if MyPage-1 > 0 then
			Prev_Page = MyPage - 1
			Response.write "<a class=black href='" & Url & "page=" & Prev_Page & "' title='��" & Prev_Page & "ҳ'>��һҳ</a> "
		end if
	
		if Maxpages>=PageNextSize*PageShowSize then
			PageSizeShow = PageShowSize
		Else
			PageSizeShow = Maxpages-PageShowSize*(PageNextSize-1)
		End if
		If PageSizeShow < 1 Then PageSizeShow = 1
		for PageCounterSize=1 to PageSizeShow
			PageLink = (PageCounterSize+PageNextSize*PageShowSize)-PageShowSize
			if PageLink <> MyPage Then
				Response.write "<a class=black href='" & Url & "page=" & PageLink & "'>[" & PageLink & "]</a> "
			else
				Response.Write "<B>["& PageLink &"]</B> "
			end if
			If PageLink = MaxPages Then Exit for
		Next
		if Mypage+1 <=Pagetpage  then
			Next_Page = MyPage + 1
			Response.write "<a class=black href='" & Url & "page=" & Next_Page & "' title='��" & Next_Page & "ҳ'>��һҳ</A>"
		end if
		if MaxPages > PageShowSize*PageNextSize then
			PageNext = PageShowSize * PageNextSize + 1
			Response.write " <A class=black href='" & Url & "page=" & Pagetpage & "' title='��"& Pagetpage &"ҳ'>ҳβ</A>"
			Response.write " <a class=black href='" & Url & "page=" & PageNext & "' title='��" & PageShowSize & "ҳ'>��һ��ҳ</a>"
		End if
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		%>
	</td>
</tr>
</table>
</div>
</body>
</html>
<!--#include file=Admin_Bottom.asp-->
	<%else
		Show_Message("û��Ҫ��˵�����")
		response.end
	end if
end if%>