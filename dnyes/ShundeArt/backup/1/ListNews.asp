<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc" -->
<!--#include file="chkuser.asp" -->
<%
smallclassid=request("smallclassid")
if smallclassid="" then
	Show_Err("未指定参数<br><br><a href='javascript:history.back()'>回去重来</a>")
	response.end
else
	if not IsNumeric(smallclassid) then
		Show_Err("非法参数<br><br><a href='javascript:history.back()'>回去重来</a>")
		response.end
	end if
end if
	dim smallclassname,bigclassname,typename,bigclassid,typeid
	set rs=server.createobject("adodb.recordset")
	sql="select * from "& db_SmallClass_Table &" where smallclassid="&smallclassid
	rs.open sql,conn,1,3
	smallclassname=rs("smallclassname")
	bigclassid=rs("bigclassid")
	master6=rs("smallclassma")
	if master6<>"" then
	smallmaster6=split(master6,"|")
	 for i=0 to ubound(smallmaster6)
		if request.cookies(Forcast_SN)("username")=trim(smallmaster6(i)) then
		smallclassmaster6=true
		exit for
		else
		smallclassmaster6=false
		end if
	 next
	end if
	set rs1=server.createobject("adodb.recordset")
	sql1="select * from "& db_BigClass_Table &" where bigclassid="&rs("bigclassid")
	rs1.open sql1,conn,1,3
	bigclassname=rs1("bigclassname")
	typeid=rs1("typeid")
	master5=rs1("bigclassmaster")
	if master5<>"" then
	bigmaster5=split(master5,"|")
	 for i=0 to ubound(bigmaster5)
		if request.cookies(Forcast_SN)("username")=trim(bigmaster5(i)) then
		bigclassmaster5=true
		exit for
		else
		bigclassmaster5=false
		end if
	 next
	end if
	set rs2=server.createobject("adodb.recordset")
	sql2="select * from "& db_Type_Table &" where typeid="&typeid
	rs2.open sql2,conn,1,3
	typename=rs2("typename")
	master4=rs2("typemaster")
	if master4<>"" then
	tmaster4=split(master4,"|")
	 for i=0 to ubound(tmaster4)
		if request.cookies(Forcast_SN)("username")=trim(tmaster4(i)) then
		typemaster4=true
		exit for
		else
		typemaster4=false
		end if
	 next
	end if
	rs2.close
	set rs2=nothing
	rs1.close
	set rs1=nothing
	rs.close
	set rs=nothing
	if typemaster4=true or request.cookies(Forcast_SN)("key")="super" or request.cookies(Forcast_SN)("key")="check" or bigclassmaster5=true or smallclassmaster6=true or request.cookies(Forcast_SN)("key")="selfreg" then
	
	dim jingyong
	set rs=server.createobject("adodb.recordset")
	sql="select * from "& db_User_Table &" where  "& db_User_Name &"='"&request.cookies(Forcast_SN)("username")&"'"
	rs.open sql,ConnUser,1,3
	jingyong=rs("jingyong")
	rs.close
	set rs=nothing
	
	if request.cookies(Forcast_SN)("key")="super" or request.cookies(Forcast_SN)("key")="check" or request.cookies(Forcast_SN)("key")="bigmaster" or request.cookies(Forcast_SN)("key")="smallmaster" or (request.cookies(Forcast_SN)("key")="selfreg" and jingyong=0) or request.cookies(Forcast_SN)("key")="typemaster" then%>
	<%
	PageShowSize = 10            '每页显示多少个页
	MyPageSize   = 16           '每页显示多少条文章
		
	If Not IsNumeric(Request("page")) Or IsEmpty(Request("page")) Or Request("page") <=0 Then
	MyPage=1
	Else
	MyPage=Int(Abs(Request("page")))
	End if
	
	
	set rs=server.CreateObject("ADODB.RecordSet")
	if request.cookies(Forcast_SN)("KEY")="super" or request.cookies(Forcast_SN)("key")="check" or request.cookies(Forcast_SN)("key")="bigmaster" or request.cookies(Forcast_SN)("key")="typemaster" then
	rs.Source="select * from "& db_News_Table &" where smallclassid=" & smallclassid & " order by NewsID desc"
	end if
	if request.cookies(Forcast_SN)("key")="selfreg" then
	rs.Source="select * from "& db_News_Table &" where smallclassid=" & smallclassid & " and editor='" & request.cookies(Forcast_SN)("username") & "' order by NewsID desc"
	end if
	if request.cookies(Forcast_SN)("key")="smallmaster" then
		if smallmana=1 then
	rs.Source="select * from "& db_News_Table &" where smallclassid=" & smallclassid & " order by NewsID desc"
		else
	rs.Source="select * from "& db_News_Table &" where smallclassid=" & smallclassid & " and editor='" & request.cookies(Forcast_SN)("username") & "' order by NewsID desc"
		end if
	end if
	rs.Open rs.Source,conn,3,1
	
	If Not rs.eof then
	rs.PageSize     = MyPageSize
	MaxPages         = rs.PageCount
	rs.absolutepage = MyPage
	total            = rs.RecordCount
	%>
<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=site.css rel=stylesheet>
<title><%=copyright%><%=version%>&nbsp;<%=ver%> - 管理文章</title>
</head>
<body topmargin="0"><!--#include file=Admin_Top.asp-->
<div align="center">
<table border="1" cellpadding="3" cellspacing="0" style="border-collapse: collapse" bordercolor="#C0C0C0" width="100%" id="AutoNumber1">
<tr class="TDtop">
  <td colspan=7 height="20" align="left" ><a href=TypeManage.asp>全部总栏</a></td>
</tr>
<tr class="TDtop1">
    <td colspan=7 width="100%" height="30" align="left" >&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src=images/menu1.gif><b><a href="BigClassManage.asp?typeid=<%=typeid%>"><font color=red><%=typename%></font></a></b>&nbsp;&nbsp;&nbsp;&nbsp;其他总栏：
<%
Set rs1 = Server.CreateObject("ADODB.Recordset")
sql1 ="SELECT  * From "& db_Type_Table &" where typeid<>"&typeid&" order by typeorder"
RS1.open sql1,Conn,3,3
do while not rs1.EOF
master1=rs1("typemaster")
if master1<>"" then
tmaster1=split(master1,"|")
 for i=0 to ubound(tmaster1)
	if request.cookies(Forcast_SN)("username")=trim(tmaster1(i)) then
	typemaster1=true
	exit for
	else
	typemaster1=false
	end if
 next
end if
if typemaster1=true or request.cookies(Forcast_SN)("key")="super" or request.cookies(Forcast_SN)("key")="check" or request.cookies(Forcast_SN)("key")="selfreg" then

%><a href=BigClassManage.asp?typeid=<%=rs1("typeid")%>><%=rs1("typename")%></a>
<%end if
rs1.movenext
loop
rs1.close
set rs1=nothing%>
      <br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src=images/menu1.gif><b><a href="SmallClassManage.asp?bigclassid=<%=bigclassid%>"><font color=red><%=bigclassname%></font></a></b>
      &nbsp;&nbsp;&nbsp;&nbsp;其他大类：
<%
Set rs1 = Server.CreateObject("ADODB.Recordset")
sql1 ="SELECT  * From "& db_BigClass_Table &" where typeid="&typeid&" and bigclassid<>"&bigclassid&" order by bigclassorder"
RS1.open sql1,Conn,3,3
do while not rs1.EOF
master=rs1("bigclassmaster")
if master<>"" then
bigmaster=split(master,"|")
 for i=0 to ubound(bigmaster)
	if request.cookies(Forcast_SN)("username")=trim(bigmaster(i)) then
	bigclassmaster=true
	exit for
	else
	bigclassmaster=false
	end if
 next
end if
if (bigclassmaster=true and request.cookies(Forcast_SN)("key")="bigmaster") or request.cookies(Forcast_SN)("key")="super" or request.cookies(Forcast_SN)("key")="check" or (request.cookies(Forcast_SN)("key")="typemaster" and typemaster4=true) or request.cookies(Forcast_SN)("key")="selfreg" then
%><a href=SmallClassManage.asp?bigclassid=<%=rs1("bigclassid")%>><%=rs1("bigclassname")%></a>
<%end if
rs1.movenext
loop
rs1.close
set rs1=nothing%>
      <br>
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src=images/menu1.gif><b><a href="listnews.asp?smallclassid=<%=smallclassid%>"><font color=red><%=smallclassname%></font></a></b>
      &nbsp;&nbsp;&nbsp;&nbsp;其他小类：
<%if request.cookies(Forcast_SN)("key")="super" or request.cookies(Forcast_SN)("key")="check" or request.cookies(Forcast_SN)("key")="typemaster" or request.cookies(Forcast_SN)("key")="bigmaster" or request.cookies(Forcast_SN)("key")="selfreg" then
Set rs1 = Server.CreateObject("ADODB.Recordset")
sql1 ="SELECT  * From "& db_SmallClass_Table &" where bigclassid="&bigclassid&" and smallclassid<>"&smallclassid&" order by smallclassorder"
RS1.open sql1,Conn,3,3
do while not rs1.EOF
%><a href=ListNews.asp?SmallClassID=<%=rs1("smallclassid")%>><%=rs1("smallclassname")%></a>
<%rs1.movenext
loop
rs1.close
set rs1=nothing
else
Set rs1 = Server.CreateObject("ADODB.Recordset")
sql1 ="SELECT * From "& db_SmallClass_Table &" where smallclassid<>"&smallclassid&" order by SmallClassorder"
RS1.open sql1,Conn,3,3
do while not rs1.EOF
master2=rs1("smallclassma")
if master2<>"" then
smallmaster=split(master2,"|")
 for i=0 to ubound(smallmaster)
	if request.cookies(Forcast_SN)("username")=trim(smallmaster(i)) then
	smallclassmaster=true
	exit for
	else
	smallclassmaster=false
	end if
 next
end if
if smallclassmaster=true then
%><a href=listnews.asp?smallclassid=<%=rs1("smallclassid")%>><%=rs1("smallclassname")%></a>
<%end if
rs1.movenext
loop
rs1.close
set rs1=nothing
end if
%></td>
</tr>
<tr>
    <td colspan=7 align=center height=25>共 <%=total%> 条，当前第 <%=Mypage%>/<%=Maxpages%> 
            页，每页 <%=MyPageSize%> 条（ID号为红色，表示该文章未经过审核）&nbsp;&nbsp;<a href='<%if eWebEditor="1" then%>NewsAdd1_eWebEditor.asp<%else%>NewsAdd1.asp<%end if%>?smallclassid=<%=smallclassid%>'><U>发表新文章</U></a></td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
<td width="7%">ID号</td>
<td width="45%">标 &nbsp;&nbsp;&nbsp;题</td>
<td width="16%">日&nbsp;&nbsp;&nbsp;期</td>
<td width="6%">编辑</td>
<td width="6%">审核</td>
<td width="6%">固顶</td>
<td width="6%">删除</td>		
</tr>
<%
for i=1 to rs.PageSize
if not rs.EOF then
title=htmlencode4(trim(rs("title")))
%>
<%
if request.cookies(Forcast_SN)("key")="super" or request.cookies(Forcast_SN)("key")="bigmaster" or request.cookies(Forcast_SN)("key")="smallmaster" or (request.cookies(Forcast_SN)("KEY")="check" and checkmod="1") or request.cookies(Forcast_SN)("key")="typemaster" or request.cookies(Forcast_SN)("key")="selfreg" then
%>	
<tr>
<td align=center ><%if rs("checkked")<>1 then%><font color=red><%end if%><%=rs("NewsID")%></font>　</td>
<td>&nbsp;<a href='ReadNews.asp?NewsID=<%=rs("NewsID")%>' target=_blank title="查看文章“<%=Title%>”"><font color="<%=rs("titlecolor")%>"><%=left(title,40)%></font></a><%if rs("picname")<>"" then%><font color=red>[图]<%end if%><%if request.cookies(Forcast_SN)("key")="super" and request.cookies(Forcast_SN)("purview")="99999" then%><font color=red>(<%=rs("editor")%>)</font><%end if%></td>
<td align=center ><%=rs("UpdateTime")%>　</td>
<td align=center ><a href='NewsEdit.asp?Newsid=<%=rs("newsid")%>'>编辑</a></td>
<td align=center >
	<%IF not(request.cookies(Forcast_SN)("KEY")="super" or request.cookies(Forcast_SN)("KEY")="check" or request.cookies(Forcast_SN)("key")="typemaster" or request.cookies(Forcast_SN)("key")="bigmaster" or request.cookies(Forcast_SN)("key")="smallmaster") THEN%>
		编辑
	<%else%>
		<a href='NewsCheck.asp?Newsid=<%=rs("newsid")%>'><%if rs("checkked")<>1 then%>审核<%else%>已审<%end if%></a>
	<%end if%>
</td>
<td align=center >
	<%IF not(request.cookies(Forcast_SN)("KEY")="super" or request.cookies(Forcast_SN)("KEY")="check" or request.cookies(Forcast_SN)("key")="typemaster" or request.cookies(Forcast_SN)("key")="bigmaster"  or request.cookies(Forcast_SN)("key")="smallmaster") THEN%>
		<%if rs("istop")<>1 then%>固顶<%else%>已固<%end if%>
	<%else%>
		<a href='NewsGu.asp?Newsid=<%=rs("newsid")%>'><%if rs("istop")<>1 then%>固顶<%else%>已固<%end if%></a>
	<%end if%>
</td>
<td align=center ><a href='NewsDel.asp?Newsid=<%=rs("newsid")%>'>删除</a></td>
</tr>
<%
end if
rs.MoveNext
end if
next
%>
<tr>
    <td colspan=7  align=center height=25>共 <%=total%> 条，当前第 <%=Mypage%>/<%=Maxpages%> 
            页，每页 <%=MyPageSize%> 条 
            <%
url="listnews.asp?smallclassid=" & smallclassid & "&"
							
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
</div>
<%else%>
<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=site.css rel=stylesheet>
<title><%=copyright%><%=version%>&nbsp;<%=ver%> - 管理文章</title>
</head>
<body topmargin="0"><!--#include file=Admin_Top.asp-->
<div align="center">
<table border="1" cellpadding="3" cellspacing="0" style="border-collapse: collapse" bordercolor="#C0C0C0" width="100%" id="AutoNumber1">
<tr>
    <td width="100%" bgcolor="<%=m_top%>" height="30" align="left">
<a href=typemanage.asp>全部总栏</a><br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src=images/menu1.gif><b><a href="BigClassManage.asp?typeid=<%=typeid%>"><font color=red><%=typename%></font></a></b>&nbsp;&nbsp;&nbsp;&nbsp;其他总栏：
<%
Set rs1 = Server.CreateObject("ADODB.Recordset")
sql1 ="SELECT  * From "& db_Type_Table &" where typeid<>"&typeid&" order by typeorder"
RS1.open sql1,Conn,3,3
do while not rs1.EOF
master1=rs1("typemaster")
if master1<>"" then
tmaster1=split(master1,"|")
 for i=0 to ubound(tmaster1)
	if request.cookies(Forcast_SN)("username")=trim(tmaster1(i)) then
	typemaster1=true
	exit for
	else
	typemaster1=false
	end if
 next
end if
if typemaster1=true or request.cookies(Forcast_SN)("key")="super" or request.cookies(Forcast_SN)("key")="check" or request.cookies(Forcast_SN)("key")="selfreg" then
%><a href=BigClassManage.asp?typeid=<%=rs1("typeid")%>><%=rs1("typename")%></a>
<%end if
rs1.movenext
loop
rs1.close
set rs1=nothing%>
      <br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src=images/menu1.gif><b><a href="SmallClassManage.asp?bigclassid=<%=bigclassid%>"><font color=red><%=bigclassname%></font></a></b>
      &nbsp;&nbsp;&nbsp;&nbsp;其他大类：
<%
Set rs1 = Server.CreateObject("ADODB.Recordset")
sql1 ="SELECT  * From "& db_BigClass_Table &" where typeid="&typeid&" and bigclassid<>"&bigclassid&" order by bigclassorder"
RS1.open sql1,Conn,3,3
do while not rs1.EOF
if (bigclassmaster=true and request.cookies(Forcast_SN)("key")="bigmaster") or request.cookies(Forcast_SN)("key")="super" or request.cookies(Forcast_SN)("key")="check" or (request.cookies(Forcast_SN)("key")="typemaster" and typemaster4=true) or request.cookies(Forcast_SN)("key")="selfreg" then
%><a href=SmallClassManage.asp?bigclassid=<%=rs1("bigclassid")%>><%=rs1("bigclassname")%></a>
<%end if
rs1.movenext
loop
rs1.close
set rs1=nothing%>
      <br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src=images/menu1.gif><b><a href="listnews.asp?smallclassid=<%=smallclassid%>"><font color=red><%=smallclassname%></font></a></b>
      &nbsp;&nbsp;&nbsp;&nbsp;其他小类：
<%if request.cookies(Forcast_SN)("key")="super" or request.cookies(Forcast_SN)("key")="check" or request.cookies(Forcast_SN)("key")="typemaster" or request.cookies(Forcast_SN)("key")="bigmaster" or request.cookies(Forcast_SN)("key")="selfreg" then
Set rs1 = Server.CreateObject("ADODB.Recordset")
sql1 ="SELECT  * From "& db_SmallClass_Table &" where bigclassid="&bigclassid&" and smallclassid<>"&smallclassid&" order by smallclassorder"
RS1.open sql1,Conn,3,3
do while not rs1.EOF
%><a href=listnews.asp?smallclassid=<%=rs1("smallclassid")%>><%=rs1("smallclassname")%></a>
<%rs1.movenext
loop
rs1.close
set rs1=nothing
else
Set rs1 = Server.CreateObject("ADODB.Recordset")
sql1 ="SELECT * From "& db_SmallClass_Table &" where smallclassid<>"&smallclassid&" order by SmallClassorder"
RS1.open sql1,Conn,3,3
do while not rs1.EOF
master2=rs1("smallclassma")
if master2<>"" then
smallmaster=split(master2,"|")
 for i=0 to ubound(smallmaster)
	if request.cookies(Forcast_SN)("username")=trim(smallmaster(i)) then
	smallclassmaster=true
	exit for
	else
	smallclassmaster=false
	end if
 next
end if
if smallclassmaster=true then
%><a href=listnews.asp?smallclassid=<%=rs1("smallclassid")%>><%=rs1("smallclassname")%></a>
<%end if
rs1.movenext
loop
rs1.close
set rs1=nothing
end if
%>
</td>
</tr>
<tr>
    <td  align=center height=25>该小类下还未有文章<input type=button value="发表新文章" class=button onclick="javascript:window.open('newsadd1.asp?smallclassid=<%=smallclassid%>','_self','')"  style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'"></td>
</tr>
</table>
</body>
</html>
<%end if
rs.close
set rs=nothing
conn.close
set conn=nothing
end if
else
response.write "<script>alert('对不起，你无权操作！');history.go(-1);</Script>"
Response.End 
end if
%>
<!--#include file=Admin_Bottom.asp-->