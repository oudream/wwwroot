<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="ChkUser.asp" -->
<!--#include file="ChkManage.asp" -->
<%
if not(request.cookies(Forcast_SN)("ManageKEY")="super" or request.cookies(Forcast_SN)("ManageKEY")="bigmaster" or request.cookies(Forcast_SN)("ManageKEY")="check" or request.cookies(Forcast_SN)("ManageKEY")="typemaster") then
	Show_Err("�Բ������Ĺ���Ȩ�޲�����")
	response.end
else
	if linkmana="1" or (request.cookies(Forcast_SN)("ManagePurview")="99999" and request.cookies(Forcast_SN)("purview")="99999") then
		PageShowSize = 10            'ÿҳ��ʾ���ٸ�ҳ
		MyPageSize   = 10           'ÿҳ��ʾ����������
			
		If Not IsNumeric(Request("page")) Or IsEmpty(Request("page")) Or Request("page") <=0 Then
			MyPage=1
		Else
			MyPage=Int(Abs(Request("page")))
		End if
		
		set rs=server.CreateObject("ADODB.RecordSet")
		if linkpass="1" then
			rs.Source="select * from "& db_link_Table &" order by ID desc"
		else
			rs.Source="select * from "& db_link_Table &" where pass=0 order by ID desc"
		end if
		rs.Open rs.Source,conn,1,1
		
		If Not rs.eof then
			rs.PageSize     = MyPageSize
			MaxPages         = rs.PageCount
			rs.absolutepage = MyPage
			total            = rs.RecordCount
			%>
<html><head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=site.css rel=stylesheet>
<title><%=copyright%><%=version%>&nbsp;<%=ver%> - �������ӹ���</title>
</head>
<body topmargin="0">
<!--#include file=Admin_Top.asp-->
<table border="1" cellpadding="3" cellspacing="0" style="border-collapse: collapse" bordercolor="#C0C0C0" width="100%" id="AutoNumber1">
<tr class="TDtop">
	<td colspan=8 width="100%" height="25" align="center">�� <strong>��������</strong> ��</td>
</tr>
<tr>
	<td colspan=8 align=center height=25>
		�� <%=total%> ������ǰ�� <%=Mypage%>/<%=Maxpages%> ҳ��ÿҳ <%=MyPageSize%> �� 
		<a href="linkreg.asp">�������</a>
	</td>
</tr>
<tr class="TDtop1" align="center" >
	<td width="6%">ID��</td>
	<td width="6%">����</td>
	<td width="20%">��վ����</td>
	<td width="20%">��վLOGO</td>
	<td width="14%">վ��</td>
	<td width="12%">ִ��</td>
	<td width="6%">�༭</td>
	<td width="6%">ɾ��</td>
</tr>
			<%
			for i=1 to rs.PageSize
				if not rs.EOF then
					%>
<form method="POST" action="<%if rs("pass")=0 then%>LinkPass.asp<%else%>LinkUnPass.asp<%end if%>?ID=<%=rs("ID")%>">
<tr bgcolor=ffffff>
	<td align=center><%=rs("ID")%></td>
	<td align=center  > 
		<%if rs("linktype")=2 then%>��ҳLOGO����<%else%><%if rs("linktype")=1 then%>����LOGO����<%else%>����<%end if%><%end if%>
	</td>
	<td align=center >
		<a href="<%=rs("weburl")%>" title="<%=rs("content")%>" target="_blank"><%=rs("webname")%></a>
	</td>
	<td height="35" align=center valign="middle" > 
		<%if rs("linktype")<>0 then%><a target="_blank" title="����鿴ԭͼ" href="<%=rs("logo")%>"><img src="<%=rs("logo")%>" border=0 width=88 height=31></a><%else%>��������<%end if%>
	</td>
	<td align=center  >
		<a href="mailto:<%=rs("email")%>"><%=rs("webmaster")%></a>
	</td>
	<td align=center  >
		<input type=hidden name=ID value="<%=rs("ID")%>">
		<input type=submit value="<%if rs("pass")="0" then%>ͨ��<%else%>ȡ��<%end if%>���" class=button onMouseOver="window.status='�������ť<%if rs("pass")="0" then%>ͨ��<%else%>ȡ��<%end if%>�����վ';return true;" onMouseOut="window.status='';return true;"target=_blank title="�������ť<%if rs("pass")="0" then%>ͨ��<%else%>ȡ��<%end if%>�����վ" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
	</td>
	<td align="center"  >
		<a href='LinkEdit.asp?id=<%=rs("ID")%>'>�༭</a>
	</td>
	<td align=center  >
		<a href='LinkDel.asp?id=<%=rs("ID")%>'>ɾ��</a>
	</td>
</tr>
</form>
					<%
					rs.MoveNext
				end if
			next
			%>
<tr>
	<td colspan=8  align=center height=25>
		�� <%=total%> ������ǰ�� <%=Mypage%>/<%=Maxpages%> ҳ��ÿҳ <%=MyPageSize%> �� 
			<%
			if request.cookies(Forcast_SN)("KEY")="super" then			
				url="linkmanage.asp?" 
			else
				url="linkmanage.asp?"
			end if									
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
			%>
	</td>
</tr>
</table>
</body>
</html>
<!--#include file=Admin_Bottom.asp-->
			<%
		else
			Show_Err("û��Ҫ�������ӵ���վ<br><br><a href='linkreg.asp'>�������</a>")
			response.end
		End If
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
	else
		Show_Err("�Բ��𣬸ù����Ѿ�������ϵͳ����Ա�رգ���û��Ȩ�޲�����")
		response.end
	end if
end if
%>