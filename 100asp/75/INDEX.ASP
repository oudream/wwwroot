<%@ language=vbscript%>
<html>
<head>
<title>实例　搜索引擎</title>
<%
const MaxPerPage=15 '分页显示的纪录个数
dim sql 
dim rs
dim gstBookID
dim totalPut 
dim CurrentPage
dim TotalPages
dim i,j 
%>
</head>
<body background="/image/bg.gif">
<P align=center><U><EM><STRONG>实例　搜索引擎</STRONG></EM></U></P>
<table border=0 cellspacing="1" bgcolor="#ffcb7d" align=center>
<tr bgcolor=lavenderblush bordercolor=lightgoldenrodyellow><td>请输入查询条件&nbsp;&nbsp;</td>
<form action="index.asp" id=form1 name=form1>
<td>学院名称&nbsp;<input name="name" maxlength="80" >
<input type="submit" value="提交" name="B1" class="buttonface">
 <input type="reset" value="全部重写" name="B2" class="buttonface"> 
</td></form></tr>
</table>
<%
sql=""
name=request("name")
if name<>"" then
	sql="select * from mytable where name like '%" & name & "%'"
else
	sql="select * from mytable"
end if
set rs=server.createobject("adodb.recordset")
conn = "DBQ=" + server.mappath("message.mdb") + ";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
rs.open sql,conn,1,1
rs.pagesize=MaxPerPage 
If trim(Request("Page"))<>"" and IsNumeric(Request("Page")) then
	CurrentPage= CLng(request("Page")) 
	If CurrentPage< 1 then
		CurrentPage= 1 
	ElseIf CurrentPage> rs.PageCount then 
		CurrentPage = rs.PageCount 
	End If 
Else 
	CurrentPage= 1 
End If 
if rs.eof then 
	response.write "<p align='center'> 记录添加中......</p>" 
else 
	totalPut=rs.recordcount '记录总数 
	if CurrentPage=1 then 
		showpages 
		showContent 
	else 
		if (currentPage-1)*MaxPerPage<totalPut then 
			rs.move(currentPage-1)*MaxPerPage 
			dim bookmark 
			bookmark=rs.bookmark '移动到开始显示的记录位置 
			showpages 
			showContent 
		else 
			currentPage=1 
			showContent 
		end if
	end if 
	rs.close 
end if 
set rs=nothing 
set conn=nothing 
sub showContent
%>
<p><TABLE border=1 bordercolor="#ffdeab" width="100%"  align=center  cellspacing=0>
 <tr>
 <td>大学名称</td>
 <td>大学排行</td>
 <td>名次升降</td>
 <td>综合得分</td>
 <td>学术声誉得分</td>
 <td>学术地位得分</td>
 <td>学术成果得分</td>
 <td>学生情况得分</td>
 <td>物资资源得分</td>
 <td>教师资源得分</td></tr>
<%
dim i 
i=0 
do while not (rs.eof or err)
%><tr>
<td><%=rs("name")%></td>
<td><%=rs("id")%></td>
<td> <%=rs("名次升降")%></td>
<td><%=rs("综合得分")%></td>
<td><%=rs("学术声誉得分")%></td>
<td><%=rs("学术地位得分")%></td>
<td><%=rs("学术成果得分")%></td>
<td><%=rs("学生情况得分")%></td>
<td><%=rs("教师资源得分")%></td>
<td><%=rs("物资资源得分")%></td>
</tr>
 <%
i=i+1
if i>=MaxPerPage then exit do
 rs.movenext 
loop
%></TABLE>
<%end sub
sub showpages() 
%>
<p align='left'>
	 <%
	dim n,k 
	if (totalPut mod MaxPerPage)=0 then
	 n= totalPut \ MaxPerPage
	else
	 n= totalPut \ MaxPerPage + 1
	end if%>全国重点大学排行榜&nbsp;：共计有<%=rs.recordcount%> 所大学<%
	k=currentPage
if k<>1 then
	response.write "[<b>"+"<a href='index.asp?name="+cstr(name)+"&page=1'>首页</a></b>] "
	response.write "[<b>"+"<a href='index.asp?name="+cstr(name)+"&page="+cstr(k-1)+"'>上一页</a></b>] "
end if
if k<>n then
	response.write "[<b>"+"<a href='index.asp?name="+cstr(name)+"&page="+cstr(k+1)+"'>下一页</a></b>] "
	response.write "[<b>"+"<a href='index.asp?name="+cstr(name)+"&page="+cstr(n)+"'>尾页</a></b>] "
end if
end sub
%>
</p>
</body>
</html>