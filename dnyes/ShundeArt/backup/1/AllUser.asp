<%@ Language=VBScript%>
<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->

<%order=checkstr(request("order"))
order2=checkstr(request("order2"))
set rs=server.CreateObject("ADODB.RecordSet")
rs.Source="select * from "& db_Type_Table &" order by typeorder"
rs.Open rs.Source,conn,1,1

dim ArraytypeID(10000),ArraytypeName(10000),Arraytypecontent(10000)
typeCount=rs.RecordCount
for i=1 to typeCount
	ArraytypeID(i)=rs("typeID")
	ArraytypeName(i)=rs("typeName")
	Arraytypecontent(i)=rs("typecontent")
	rs.MoveNext
next
rs.Close
set rs=nothing
PageShowSize = 10	'ÿҳ��ʾ���ٸ�ҳ
MyPageSize   = 50	'ÿҳ��ʾ���ٸ��û���
If Not IsNumeric(Request("page")) Or IsEmpty(Request("page")) Or Request("page") <=0 Then
	MyPage=1
Else
	MyPage=Int(Abs(Request("page")))
End if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>ȫ����Ա__<%=jjgn%></title>
<LINK href=news.css rel=stylesheet>
</head>
<body topmargin="0">
<!--#include file="top.asp"-->
<table width="760" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
	<tr valign="top"> 
		<td width="760" class="daohang"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
			<tr> 
				<td height="3" bgcolor="#FFFFFF"><img src="images/kb.gif" width="9" height="3"></td>
			</tr>
			<tr> 
				<td height="25" background="images/menu-guestbook.gif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font class=m_tittle>&nbsp;</font>��Ŀ����&nbsp;&nbsp;<b>��ǰλ�ã�<a class="daohang" href="./" >��վ��ҳ</a>&gt;&gt;ȫ����Ա</b></td>
			</tr>
			<tr> 
				<td height="25" background="images/menu-guest-l.gif"><div align="center"><a class=daohang href="./" ><script language=javascript src="./zongg/ad.asp?i=13"></script></a></div></td>
			</tr>
			<tr> 
				<td height="25" background="images/menu-guestbook.gif"></td>
			</tr>
			<tr> 
				<td height="25" background="images/menu-guest-l.gif">
					<div align="center"> 
						<table width="95%" height="156" border="0" align="center" cellpadding="0" cellspacing="0" id="AutoNumber3" style="border-collapse: collapse">
							<tr> 
								<td height="10" valign=top>&nbsp;</td>
							</tr>
							<tr> 
								<td height="10" valign=top>
									<table width="100%" border="0" cellpadding="0" cellspacing="0">
<%Response.write "<tr>"& chr(10)
Response.write "<td width='20%' height='20' colspan='5'>"
function getusernum()
	dim rs
	rs=ConnUser.execute("Select Count("& db_User_Id &") from "& db_User_Table &"")
	getusernum=rs(0)
	set rs=nothing
	if isnull(getusernum) then getusernum=0
end function
Response.write "��վ����"& getusernum &"λ��Ա&nbsp;&nbsp;&nbsp;&nbsp;"
	Response.write "<a class=class href='AllUser.asp?order2=1' title='����'>[��</a>���û���<a class=class href='AllUser.asp?order2=2' title='����'>��]</a>&nbsp;&nbsp;"
	Response.write "<a class=class href='?order=time&order2=1' title='����'>[��</a>ע��ʱ��<a class=class href='?order=time&order2=2' title='����'>��]</a>&nbsp;&nbsp;"
	Response.write "<a class=class href='?order=number&order2=2' title='����'>[��</a>��������<a class=class href='?order=number&order2=1' title='����'>��]</a>&nbsp;&nbsp;"
Response.write "</td>"& chr(10)
Response.write "</tr>"& chr(10)

set rs1=server.CreateObject("ADODB.RecordSet")
if order="time" then
	if order2=1 then
		rs1.Source="SELECT * FROM "& db_User_Table &" order by "& db_User_RegDate
	else
		rs1.Source="SELECT * FROM "& db_User_Table &" order by "& db_User_RegDate &" desc"
	end if
else
	if order="number" then
		if order2=1 then
			rs1.Source="SELECT * FROM "& db_User_Table &" order by number"
		else
			rs1.Source="SELECT * FROM "& db_User_Table &" order by number desc"
		end if
	else
		if order2=1 then
			rs1.Source="SELECT * FROM "& db_User_Table &" order by "& db_User_Name
		else
			rs1.Source="SELECT * FROM "& db_User_Table &" order by "& db_User_Name &" desc"
		end if
	end if
end if
rs1.Open rs1.Source,ConnUser,3,1
If Not rs1.eof then
	rs1.PageSize     = MyPageSize
	MaxPages         = rs1.PageCount
	rs1.absolutepage = MyPage
	total            = rs1.RecordCount
	i = 1
	do until rs1.Eof or (i > MyPageSize)
		Response.write "<tr>"& chr(10)
		for ii=1 to 8	'�û�ÿ������
			Response.write "<td width='12.5%' height='20'>" 
			if not rs1.eof and (i <= MyPageSize) then
				Response.write "<a class=class href='User.asp?User="& rs1(db_User_Name) &"' title='"
				if order="time" then
					Response.write "["& rs1(db_User_RegDate) &"]"
				end if
				Response.write "'>"& left(rs1(db_User_Name),6) &"</a>"
				if order="number" then
					Response.write "&nbsp;&nbsp;<font color=red>["& rs1("Number") &"]</font>"
				end if
				rs1.movenext
				i = i + 1  
			end if
			Response.write "</td>"& chr(10)			 
		next
		Response.write "</tr>"& chr(10)
	loop
	Response.write "</table>"& chr(10)
	Response.write "</td>"& chr(10)
        Response.write "</tr>"& chr(10)
	Response.write "<tr>"& chr(10)
	Response.write "<td width='100%' align=center valign='middle'>�� "& total &" λ��Ա����ǰ�� "& Mypage &"/"& Maxpages &" ҳ��ÿҳ "& MyPageSize &" λ��Ա <br><br>"
	url="AllUser.asp?order="&order									
	PageNextSize=int((MyPage-1)/PageShowSize)+1
	Pagetpage=int((total-1)/rs1.PageSize)+1
	if PageNextSize >1 then
		PagePrev=PageShowSize*(PageNextSize-1)
		Response.write "<a class=black href='" & Url & "&page=" & PagePrev & "' title='��" & PageShowSize & "ҳ'>��һ��ҳ</a> "
		Response.write "<a class=black href='" & Url & "&page=1' title='��1ҳ'>ҳ��</a> "
	end if
	if MyPage-1 > 0 then
		Prev_Page = MyPage - 1
		Response.write "<a class=black href='" & Url & "&page=" & Prev_Page & "' title='��" & Prev_Page & "ҳ'>��һҳ</a> "
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
				Response.write "<a class=black href='" & Url & "&page=" & PageLink & "'>[" & PageLink & "]</a> "
			else
				Response.Write "<B>["& PageLink &"]</B> "
			end if
			If PageLink = MaxPages Then Exit for
		Next
	
		if Mypage+1 <=Pagetpage  then
			Next_Page = MyPage + 1
			Response.write "<a class=black href='" & Url & "&page=" & Next_Page & "' title='��" & Next_Page & "ҳ'>��һҳ</A>"
		end if
	
		if MaxPages > PageShowSize*PageNextSize then
			PageNext = PageShowSize * PageNextSize + 1
			Response.write " <A class=black href='" & Url & "&page=" & Pagetpage & "' title='��"& Pagetpage &"ҳ'>ҳβ</A>"
			Response.write " <a class=black href='" & Url & "&page=" & PageNext & "' title='��" & PageShowSize & "ҳ'>��һ��ҳ</a>"
		End if
					
	End If
	rs1.close
	set rs1=nothing
			%>
									</td>
								</tr>
								<tr> 
									<td height="10" valign=top>&nbsp;</td>
								</tr>
							</table>
						</div>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr valign="top">
		<td height="20" background="images/menu-guest-t.gif">��</td>
		<td>��</td>
		<td height="20" background="images/menu-guest-t.gif" class="daohang"></td>
	</tr>
</table>
<!--#include file="bottom.asp"-->
</body>
</html>