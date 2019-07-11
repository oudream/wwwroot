<html>
<title>分页技术</title>
<head>
<%
   const MaxPerPage=25
   dim sql 
   dim rs
   dim totalPut   
   dim CurrentPage
   dim TotalPages
   dim i,j 
%>
</head>
<body bgcolor=Thistle>
<P align=center><FONT face=方正舒体><FONT size=5><STRONG>实例&nbsp;&nbsp;&nbsp; 
分页技术</STRONG>

<%
conn = "DBQ=" + server.mappath("mydb.mdb") + ";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};" 
sql = "SELECT * FROM mytable"
set rs=server.createobject("adodb.recordset") 
rs.open SQL,conn,1,1
rs.MoveFirst
rs.pagesize=MaxPerPage
howmanyfields=rs.Fields.Count-1

If trim(Request("Page"))<>"" then
	CurrentPage= CLng(request("Page")) 
	If CurrentPage> rs.PageCount then 
		CurrentPage = rs.PageCount 
	End If 
Else 
	CurrentPage= 1 
End If 

if rs.eof then 
	response.write "<p align='center'> ERROR!</p>" 
else 
	totalPut=rs.recordcount
	if CurrentPage<>1 then 
		if (currentPage-1)*MaxPerPage<totalPut then 
			rs.move(currentPage-1)*MaxPerPage 
			dim bookmark 
			bookmark=rs.bookmark
		end if 
	end if

	dim n,k 
	if (totalPut mod MaxPerPage)=0 then  
		n= totalPut \ MaxPerPage
	else  
		n= totalPut \ MaxPerPage + 1  
	end if%>
	</FONT></FONT></P>
<P>PAGE <%=currentpage%> OF <%=n%>   共<%=rs.recordcount%> 纪录
	<%	k=currentPage
		if k<>1 then
			response.write "[<b>"+"<a href='index.asp?page=1'>首页</a></b>] "
			response.write "[<b>"+"<a href='index.asp?page="+cstr(k-1)+"'>上一页</a></b>] "
		else
			Response.Write "[首页] [上一页]"
		end if
		if k<>n then
			response.write "[<b>"+"<a href='index.asp?page="+cstr(k+1)+"'>下一页</a></b>] "
			response.write "[<b>"+"<a href='index.asp?page="+cstr(n)+"'>尾页</a></b>] "
		else
			Response.Write "[下一页] [尾页]"
		end if
	%>

</P>

<TABLE border=1 align=center>
	<tr>
	<%
	for i= 0 to howmanyfields%>
		<TD><B><%=rs(i).name%></B></TD>
	<%
	next	
	i=0
	do while not rs.eof and i<maxperpage%>
	<tr align=middle>
		<%for j=0 to howmanyfields%>
		<td>
			&nbsp;<%=rs(j)%>&nbsp;
		</td>
		<%next%>
	</tr>
	<%
	i=i+1
	rs.movenext 
	loop
	%>
</TABLE>
<%
end if 
rs.close 
set rs=nothing 
%>
</body>
</html>