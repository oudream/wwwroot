<HTML>
<%
dim IID
IID=request("id")
if IID="" then Response.Redirect "index.asp"%>
<HEAD>
<title>ͬѧ¼</title>
</HEAD>
<BODY>
<P align=center><FONT face=���� size=6>ʵ��&nbsp; ͬѧ¼</FONT></P>
<table align=center><tr><td>
<%
set rs=server.createobject("adodb.recordset") 
conn = "DBQ=" + server.mappath("mydb.mdb") + ";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};" 
sql="select * from mytable where id="+IID
rs.open sql,conn,1,1
if rs.EOF then
	Response.Write "���ݿ�����������!"
else
	%>
	<FORM action="modifynews.asp?id=<%=rs("id")%>" method=post  name=form1><FONT face=���Ĳ��� size=4>������</FONT>
	</td><td><FONT face=���Ĳ��� size=4><INPUT  name=text1 value=<%=rs("username")%>></FONT></td></tr>
	  <TR><td><FONT face=���Ĳ��� size=4>�Ա�</FONT>
	</td><td>
		<SELECT name=sel1>
		<OPTION selected><%=rs("sex")%></OPTION>
		<OPTION>��</OPTION>
		<OPTION>Ů</OPTION></SELECT></td></TR>
	  <TR><td><FONT face=���Ĳ��� size=4>���գ�</FONT>
	</td><td><FONT face=���Ĳ��� size=4><INPUT name=text2 value=<%=rs("borndate")%>></FONT></td></TR>
	  <TR><td><FONT face=���Ĳ��� size=4>��ϵ�绰��</FONT>
	</td><td><FONT face=���Ĳ��� size=4><INPUT name=text3 value=<%=rs("phone")%>></FONT></td></TR>
	  <TR><td><FONT face=���Ĳ��� size=4>������</FONT>
	</td><td><FONT face=���Ĳ��� size=4><INPUT name=text4 value=<%=rs("BP")%>></FONT></td></TR>
	  <TR><td><FONT face=���Ĳ��� size=4>�ֻ���</FONT>
	</td><td><FONT face=���Ĳ��� size=4><INPUT name=text5 value=<%=rs("policy")%>></FONT></td></TR>
	  <TR><td><FONT face=���Ĳ��� size=4>��ͥסַ��</FONT>
	</td><td><FONT face=���Ĳ��� size=4><INPUT name=text6 value=<%=rs("address")%>></FONT></td></TR>
	  <TR><td><FONT face=���Ĳ��� size=4>���ڹ�˾��</FONT>
	</td><td><FONT face=���Ĳ��� size=4><INPUT name=text7 value=<%=rs("company")%>></FONT></td></TR>
	  <TR><td><FONT face=���Ĳ��� size=4>�������䣺</FONT>
	</td><td><FONT face=���Ĳ��� size=4><INPUT name=text8 value=<%=rs("email")%>></FONT></td></TR>
	  <TR><td colspan=2 align=center><FONT face=���Ĳ��� size=4>
	<INPUT type="submit" value="�� ��" id=submit1 name=submit1> 
	</FONT></FORM></td></TR></table>
<%
end if
%>
</BODY>
</HTML>
