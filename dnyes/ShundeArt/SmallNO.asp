<!--#include file="Conn.ASP"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.ASP"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp" -->
<%
dim typename,bigclassname,jingyong
bigclassid=request("bigclassid")
set rs=server.createobject("adodb.recordset")
sql="select * from "& db_BigClass_Table &" where bigclassid="&bigclassid
rs.open sql,conn,1,3
bigclassname=rs("bigclassname")
typeid=rs("typeid")
rs.close
set rs=nothing
set rs=server.createobject("adodb.recordset")
sql="select * from "& db_Type_Table &" where typeid="&typeid
rs.open sql,conn,1,3
typename=rs("typename")
rs.close
set rs=nothing

set rs=server.createobject("adodb.recordset")
sql="select * from "& db_User_Table &" where  "& db_User_Name &"='"&request.cookies(Forcast_SN)("username")&"'"
rs.open sql,ConnUser,1,3
jingyong=rs("jingyong")
rs.close
set rs=nothing

if r(request.cookies(Forcast_SN)("key")="selfreg" and jingyong=0) then
	PageShowSize = 10            'ÿҳ��ʾ���ٸ�ҳ
	MyPageSize   = 16           'ÿҳ��ʾ����������
	
	If Not IsNumeric(Request("page")) Or IsEmpty(Request("page")) Or Request("page") <=0 Then
		MyPage=1
	Else
		MyPage=Int(Abs(Request("page")))
	End if
	
	set rs=server.CreateObject("ADODB.RecordSet")
	if request.cookies(Forcast_SN)("KEY")="super" or request.cookies(Forcast_SN)("key")="check" or request.cookies(Forcast_SN)("key")="bigmaster" or request.cookies(Forcast_SN)("key")="typemaster" then
		rs.Source="select * from "& db_News_Table &" where BigClassid="&BigClassid&" and smallclassid is null order by NewsID desc"
	end if
		if request.cookies(Forcast_SN)("key")="selfreg" then
			rs.Source="select * from "& db_News_Table &" where BigClassid="&BigClassid&" and smallclassid is null and editor='" & request.cookies(Forcast_SN)("username") & "' order by NewsID desc"
		end if
		rs.Open rs.Source,conn,3,1
	
	If Not rs.eof then
		rs.PageSize     = MyPageSize
		MaxPages         = rs.PageCount
		rs.absolutepage = MyPage
		total            = rs.RecordCount
		%>
<html><head>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=site.css rel=stylesheet>
<title><%=copyright%><%=version%>&nbsp;<%=ver%> - ���¹���</title>
</head>
<body topmargin="0">
<!--#include file=Admin_Top.asp-->
<div align="center">
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#C0C0C0" width="100%" id="AutoNumber1">
<tr>
	<td width="100%"  height="30" align="center">�� �� <b><font color=red><a href="BigClassManage.asp?typeid=<%=typeid%>"><%=typename%></a></font></b> �� �� <b><font color=red><a href="SmallClassManage.asp?bigclassid=<%=bigclassid%>"><%=bigclassname%></a></font></b> ����С�������</td>
</tr>
<tr>
	<td colspan=2  align=center height=25>�� <%=total%> ������ǰ�� <%=Mypage%>/<%=Maxpages%> ҳ��ÿҳ <%=MyPageSize%> ��&nbsp;&nbsp;<a href='NewsAdd1.asp?BigClassID=<%=bigclassid%>'><U>����������</U></a></td>
</tr>
<tr>
<td colspan=2align=center height=25>
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#C0C0C0" width="100%" id="AutoNumber1">
<tr align="center" bgcolor="#FFFFFF">
<td width="7%">ID��</td>
<td width="45%">�� &nbsp;&nbsp;&nbsp;��</td>
<td width="16%">��&nbsp;&nbsp;&nbsp;��</td>
<td width="6%">�༭</td>
<td width="6%">���</td>
<td width="6%">�̶�</td>
<td width="6%">ɾ��</td>		
</tr>
</table>
</td>
</tr>
<tr>
<td width="100%" >
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#C0C0C0" width="100%" id="AutoNumber1">
		<%
		for i=1 to rs.PageSize
			if not rs.EOF then
				title=htmlencode4(trim(rs("title")))
				%>
<tr bgcolor=ffffff><td width=7% align=center ><%if rs("checkked")<>1 then%><font color=red><%end if%><%=rs("NewsID")%></font>��</td>
<td width=45% >&nbsp;<a href='ReadNews.asp?NewsID=<%=rs("NewsID")%>' target=_blank title="�鿴���¡�<%=Title%>��"><font color="<%=rs("titlecolor")%>"><%=left(title,40)%></font></a><%if rs("picname")<>"" then%><font color=red>[ͼ]<%end if%><%if request.cookies(Forcast_SN)("key")="super" and request.cookies(Forcast_SN)("purview")="99999" then%><font color=red>(<%=rs("editor")%>)</font><%end if%></td>
<td width=16% align=center ><%=rs("UpdateTime")%>��</td>
<td width=6% align=center ><a href='NewsEdit.asp?NewsID=<%=rs("newsid")%>'>�༭</a></td>
<td width=6% align=center ><a href='NewsCheck.asp?NewsID=<%=rs("newsid")%>'><%if rs("checkked")<>1 then%>���<%else%>����<%end if%></a></td>
<td width=6% align=center ><a href='Newsgu.asp?NewsID=<%=rs("newsid")%>'><%if rs("istop")<>1 then%>�̶�<%else%>�ѹ�<%end if%></a></td>
<td width=6% align=center ><a href='newsdel.asp?newsid=<%=rs("newsid")%>'>ɾ��</a></td>
<tr>
				<%
				rs.MoveNext
			end if
		next
		%>
</table>
</td>
<tr>
    <td colspan=2  align=center height=25>�� <%=total%> ������ǰ�� <%=Mypage%>/<%=Maxpages%> 
            ҳ��ÿҳ <%=MyPageSize%> �� 
            <%
url="smallno.asp?bigclassid="&bigclassid&"&"
								
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
</div>
</body>
</html>
<!--#include file=Admin_Bottom.asp-->
		<%rs.close
		set rs=nothing
		conn.close
		set conn=nothing
	else
		Show_Err("��û�д˴�������С�������<br><br><a href='SmallClassManage.asp?bigclassid="& bigclassid &"'>����</a>")
		response.end
	End If
end if%>
