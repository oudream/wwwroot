<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp" -->
<!--#include file="ChkURL.asp"-->
<%
IF not(request.cookies(Forcast_SN)("KEY")="super" or request.cookies(Forcast_SN)("KEY")="check") THEN
	response.redirect "index_face.asp"
	response.end
else
	usernamecookie=CheckStr(request.cookies(Forcast_SN)("UserName"))
	passwdcookie=CheckStr(trim(Request.cookies(Forcast_SN)("passwd")))
	KEYcookie=CheckStr(trim(request.cookies(Forcast_SN)("KEY")))
	if usernamecookie="" or passwdcookie="" then
		response.redirect "login.asp"
		response.end
	else
		'�ж��û��ĺϷ���
  		set rs=server.createobject("adodb.recordset")
  		sql="select * from "& db_User_Table &" where "& db_User_Name &"='"&usernamecookie&"'"
  		rs.open sql,ConnUser,1,1
  		if rs.eof and rs.bof then
             		response.redirect "login.asp"
            		response.end
         	end if
  		IF passwdcookie<>rs(db_User_Password) THEN
          		response.redirect "login.asp"
           		response.end
      		END IF
   		'�����ж��û�����ʵ�������û������Ƕ�Ӧ���ж�
   		if KEYcookie<>rs("OSKEY") then
      			response.redirect "index_face.asp"
      			response.end
    		end if
		rs.close
		set rs=nothing
	END IF
END IF

PageShowSize = 10            'ÿҳ��ʾ���ٸ�ҳ
MyPageSize   = 16           'ÿҳ��ʾ����������
	
If Not IsNumeric(Request("page")) Or IsEmpty(Request("page")) Or Request("page") <=0 Then
	MyPage=1
Else
	MyPage=Int(Abs(Request("page")))
End if

set rs=server.CreateObject("ADODB.RecordSet")
rs.Source="select * from "& db_Review_Table &" order by NewsID desc"
rs.Open rs.Source,conn,3,1

If Not rs.eof then
	rs.PageSize     = MyPageSize
	MaxPages         = rs.PageCount
	rs.absolutepage = MyPage
	total            = rs.RecordCount
	%>
<html><head>
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=site.css rel=stylesheet>
<title><%=copyright%><%=version%>&nbsp;<%=ver%>- -ɾ������</title>
</head>
<body topmargin="0">
<!--#include file=Admin_Top.asp-->
<div align="center">
<table border="1" width="750" cellspacing="0" cellpadding="0" bgcolor="#000000" style="border-collapse: collapse" bordercolor="<%=border%>">
<tr>
<td width="100%" bgcolor="<%=m_top%>" height="30" align="center"><b>ɾ �� �� ��</b></td>
</tr>
<tr>
<td  align=center height=25>�� <%=total%> ������ǰ�� <%=Mypage%>/<%=Maxpages%> ҳ��ÿҳ <%=MyPageSize%> ��</td>
</tr>
<tr>
<td bgcolor=#abb8d6 align=center height=25>
<table width="100%" border="1" cellspacing="0" cellpadding="0" height="25" style="border-collapse: collapse" bordercolor="<%=border%>">
<tr align="center" bgcolor="#FFFFFF">
<td width="6%">����ID</td>
<td width="6%">����ID</td>		
<td width="40%">������Ҫ</td>
<td width="12%">��&nbsp;&nbsp;��</td>
<td width="22%">��&nbsp;&nbsp;��</td>
<td width="6%">��&nbsp;&nbsp;��</td>
<td width="8%">ִ��</td>
</tr>
</table>
</td>
</tr>
<tr>
<td bgcolor=#ffffff>
<table border="1" cellspacing="0" width="100%" cellpadding="0" bgcolor="#abb8d6" style="border-collapse: collapse" bordercolor="<%=border%>">

<%
for i=1 to rs.PageSize
if not rs.EOF then
content=trim(rs("content"))
DisplayContent=mid(Content,1,25)
%>
<form method="POST" action="DelReview2_alert.asp"><input type=hidden name="ReViewID" value="<%=rs("ReViewID")%>">

<%if (delreview="1" and request.cookies(Forcast_SN)("key")="input") or (shdelreview="1" and request.cookies(Forcast_SN)("key")="check") or request.cookies(Forcast_SN)("key")="super" then%>
<tr bgcolor=ffffff align="center">
<td width="6%"><%=rs("ReViewID")%>��</td>
<td width="6%"><%=rs("NewsID")%>��</td>
<td width="40%" align="left">&nbsp;<a  href='ReadView.asp?ReViewID=<%=rs("ReViewID")%>' target=_blank><%=DisplayContent%>...</a></td>
<td width="12%"><%=trim(rs("author"))%>��</td>
<td width="22%"><a href="mailto:<%=trim(rs("email"))%>"><%=trim(rs("email"))%></a>��</td>
<td width="6%"><%=Month(rs("UpdateTime")) %>-<%=Day(rs("UpdateTime")) %></td>			
<td width="8%">
<input type=submit value="ɾ��" class=button></td>
</tr>
</form>
<%
end if
rs.MoveNext
end if
next
%>
</table>
</td>
<tr>
<td colspan=2  align=center height=25>�� <%=total%> ������ǰ�� <%=Mypage%>/<%=Maxpages%> ҳ��ÿҳ <%=MyPageSize%> �� 
            <%
if request.cookies(Forcast_SN)("KEY")="super" then			
url="DelReView1.asp?" 
else
url="DelReView1.asp?BigClassName=" & Request("BigClassName") & "&"
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
  </center>
</div>
<br>
<%
else
%><meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=site.css rel=stylesheet>
<title><%=copyright%><%=version%>&nbsp;<%=ver%>- -ɾ������</title>
<body topmargin="0"><!--#include file=Admin_Top.asp-->
<p>��</p>
<p>��</p>
<p>��</p>
<div align="center">
  <center>
<table border="1" width="450" bgcolor="#000000" cellpadding="0" cellspacing="0" bordercolor="<%=border%>" style="border-collapse: collapse">
  <tr> 
    <td colspan=2 align="center" bgcolor="ffffff"> 
      <p>��</p>
      <p>û�����ݿ�ѡ</p>
      <p>��</p>
    </td>
  </tr>
</table>
  </center>
</div>
<%				
End If
%>
</body>
</html>
<%
rs.close
set rs=nothing
conn.close
set conn=nothing
%>