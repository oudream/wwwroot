<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp" -->

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
		'判断用户的合法性
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
   		'下面判断用户级别实际在有用户级别是都应该判断
   		if KEYcookie<>rs("OSKEY") then
      			response.redirect "index_face.asp"
      			response.end
    		end if
           	rs.close
           	set rs=nothing
	END IF
END IF

PageShowSize = 10            '每页显示多少个页
MyPageSize   = 16           '每页显示多少条文章
	
If Not IsNumeric(Request("page")) Or IsEmpty(Request("page")) Or Request("page") <=0 Then
	MyPage=1
Else
	MyPage=Int(Abs(Request("page")))
End if

set rs=server.CreateObject("ADODB.RecordSet")

if reviewcheckshow="1" then
	rs.Source="select * from "& db_Review_Table &" order by reviewID desc"
else
	rs.Source="select * from "& db_Review_Table &" where passed=0 order by reviewID desc"
end if

rs.Open rs.Source,conn,1,1

If Not rs.eof then
	rs.PageSize     = MyPageSize
	MaxPages         = rs.PageCount
	rs.absolutepage = MyPage
	total            = rs.RecordCount
%>

<html>
<head>
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=site.css rel=stylesheet>
<title><%=copyright%><%=version%>&nbsp;<%=ver%> - 审核评论</title>
</head>
<body topmargin="0"><!--#include file=Admin_Top.asp--><br>
<div align="center">
  <center>
<table border="1" width="750" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="<%=border%>">
<tr>
    <td width="100%" bgcolor="<%=m_top%>" height="30" align="center"><b>审核评论</b></td>
</tr>
<tr>
    <td colspan=2  align=center height=25>共 <%=total%> 条，当前第 <%=Mypage%>/<%=Maxpages%> 页，每页 <%=MyPageSize%> 条</td>
</tr>
<tr>
<td colspan=2 bgcolor=#abb8d6 align=center height=25>
<table width="100%" border="1" cellspacing="0" cellpadding="0" height="25" style="border-collapse: collapse" bordercolor="<%=border%>">
<tr align="center" bgcolor="#FFFFFF">
<td width="5%">ID号</td>
<td width="7%">文章ID</td>
<td width="33%">内容提要</td>
<td width="12%">作者</td>
<td width="22%">信箱</td>
<td width="6%">日期</td>		
<td width="15%">执行</td>
</tr>
</table>
</td>
</tr>
<tr>
<td width="100%" bgcolor="#abb8d6">
<table border="1" cellspacing="0" width="100%" cellpadding="0" style="border-collapse: collapse" bordercolor="<%=border%>">
<%
for i=1 to rs.PageSize
if not rs.EOF then
content=trim(rs("content"))
DisplayContent=mid(Content,1,18)
%>
<form method="POST" action="<%if rs("passed")="0" then%>checkreview4.asp<%else%>checkreview3.asp<%end if%>?reviewid=<%=rs("reviewID")%>">
<tr bgcolor=ffffff>
<td width=5% align=center ><%=rs("reviewID")%></td>
<td width=7% align=center><a href="readnews.asp?newsid=<%=rs("newsid")%>" target="_blank"><%=rs("newsid")%></a></td>
<td width="33%" align="left">&nbsp;<a  href='ReadView.asp?ReViewID=<%=rs("ReViewID")%>' target=_blank><%=DisplayContent%>...</a></td>
<td width="12%" align=center><%=trim(rs("author"))%>　</td>
<td width="22%" align=center><a href="mailto:<%=trim(rs("email"))%>"><%=trim(rs("email"))%></a>　</td>
<td width="6%" align=center><%=Month(rs("UpdateTime")) %>-<%=Day(rs("UpdateTime")) %></td>
<td width=15% align=center ><input type=hidden name=reviewiD value="<%=rs("reviewID")%>">
<input type=submit value="<%if rs("passed")="0" then%>通过<%else%>取消<%end if%>审核" class=button onMouseOver="window.status='按这个按钮<%if rs("passed")="0" then%>通过<%else%>取消<%end if%>审核评论';return true;" onMouseOut="window.status='';return true;"target=_blank title="按这个按钮<%if rs("passed")="0" then%>通过<%else%>取消<%end if%>审核评论"></td>
</form>

<%
rs.MoveNext
end if
next
%>

</table>
</td>
<tr>
    <td colspan=2  align=center height=25>共 <%=total%> 条，当前第 <%=Mypage%>/<%=Maxpages%> 页，每页 <%=MyPageSize%> 条 
            <%
if request.cookies(Forcast_SN)("KEY")="super" then			
	url="checkreview1.asp?"
else
	url="checkreview1.asp?"
end if									
PageNextSize=int((MyPage-1)/PageShowSize)+1
Pagetpage=int((total-1)/rs.PageSize)+1

if PageNextSize >1 then
PagePrev=PageShowSize*(PageNextSize-1)
Response.write "<a class=black href='" & Url & "page=" & PagePrev & "' title='上" & PageShowSize & "页'>上一翻页</a> "
Response.write "<a class=black href='" & Url & "page=1' title='第1页'>页首</a> "
end if
if MyPage-1 > 0 then
Prev_Page = MyPage - 1
Response.write "<a class=black href='" & Url & "page=" & Prev_Page & "' title='第" & Prev_Page & "页'>上一页</a> "
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
Response.write "<a class=black href='" & Url & "page=" & Next_Page & "' title='第" & Next_Page & "页'>下一页</A>"
end if

if MaxPages > PageShowSize*PageNextSize then
PageNext = PageShowSize * PageNextSize + 1
Response.write " <A class=black href='" & Url & "page=" & Pagetpage & "' title='第"& Pagetpage &"页'>页尾</A>"
Response.write " <a class=black href='" & Url & "page=" & PageNext & "' title='下" & PageShowSize & "页'>下一翻页</a>"
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
%>
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=site.css rel=stylesheet>
<title><%=copyright%><%=version%>&nbsp;<%=ver%> - 审核评论</title>
<body topmargin="0"><!--#include file=Admin_Top.asp-->
<div align="center">
  <center>
<table border="1" width="750" bgcolor="#000000" cellpadding="0" cellspacing="0" bordercolor="<%=border%>" style="border-collapse: collapse">
  <tr> 
    <td colspan=2 align="center" bgcolor="ffffff"> 
      <p>　</p>
      <p>没有要审核的评论</p>
      <p>　</p>
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