<HTML>
<%
dim IID
IID=request("id")
if IID="" then Response.Redirect "index.asp"%>
<HEAD>
<title>同学录</title>
</HEAD>
<BODY>
<P align=center><FONT face=隶书 size=6>实例&nbsp; 同学录</FONT></P>
<table align=center><tr><td>
<%
set rs=server.createobject("adodb.recordset") 
conn = "DBQ=" + server.mappath("mydb.mdb") + ";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};" 
sql="select * from mytable where id="+IID
rs.open sql,conn,1,1
if rs.EOF then
	Response.Write "数据库中暂无资料!"
else
	%>
	<FORM action="modifynews.asp?id=<%=rs("id")%>" method=post  name=form1><FONT face=华文彩云 size=4>姓名：</FONT>
	</td><td><FONT face=华文彩云 size=4><INPUT  name=text1 value=<%=rs("username")%>></FONT></td></tr>
	  <TR><td><FONT face=华文彩云 size=4>性别：</FONT>
	</td><td>
		<SELECT name=sel1>
		<OPTION selected><%=rs("sex")%></OPTION>
		<OPTION>男</OPTION>
		<OPTION>女</OPTION></SELECT></td></TR>
	  <TR><td><FONT face=华文彩云 size=4>生日：</FONT>
	</td><td><FONT face=华文彩云 size=4><INPUT name=text2 value=<%=rs("borndate")%>></FONT></td></TR>
	  <TR><td><FONT face=华文彩云 size=4>联系电话：</FONT>
	</td><td><FONT face=华文彩云 size=4><INPUT name=text3 value=<%=rs("phone")%>></FONT></td></TR>
	  <TR><td><FONT face=华文彩云 size=4>传呼：</FONT>
	</td><td><FONT face=华文彩云 size=4><INPUT name=text4 value=<%=rs("BP")%>></FONT></td></TR>
	  <TR><td><FONT face=华文彩云 size=4>手机：</FONT>
	</td><td><FONT face=华文彩云 size=4><INPUT name=text5 value=<%=rs("policy")%>></FONT></td></TR>
	  <TR><td><FONT face=华文彩云 size=4>家庭住址：</FONT>
	</td><td><FONT face=华文彩云 size=4><INPUT name=text6 value=<%=rs("address")%>></FONT></td></TR>
	  <TR><td><FONT face=华文彩云 size=4>所在公司：</FONT>
	</td><td><FONT face=华文彩云 size=4><INPUT name=text7 value=<%=rs("company")%>></FONT></td></TR>
	  <TR><td><FONT face=华文彩云 size=4>电子信箱：</FONT>
	</td><td><FONT face=华文彩云 size=4><INPUT name=text8 value=<%=rs("email")%>></FONT></td></TR>
	  <TR><td colspan=2 align=center><FONT face=华文彩云 size=4>
	<INPUT type="submit" value="提 交" id=submit1 name=submit1> 
	</FONT></FORM></td></TR></table>
<%
end if
%>
</BODY>
</HTML>
