<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="css.css" type=text/css rel=stylesheet>
</head>
<body text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="50">&nbsp;</td>
  </tr>
</table>

<!--#include file="adminconn.asp" -->
<%
sql="select * from sort" 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1 

if rs.eof or rs.bof then 
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' None data ! ');</SCRIPT>")
response.End()
end if
%>
<table width="97%" border="1" align="center" cellpadding="2" cellspacing="0" bordercolor="#CCCCCC">
  <tr bgcolor="#FFFFFF"> 
    <td height="22" colspan="8" align="center" class="header"> </td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <td height="22" colspan="8" align="center" class="header">种类列表</td>
  </tr>
  <tr> 
    <td width="100" height="30" bgcolor="#FFFFFF">名称</td>
    <td width="600" bgcolor="#FFFFFF">备注</td>
  </tr>
  <%
do while not rs.eof 
%>
  <tr> 
    <td height="25" bgcolor="#FFFFFF"><a href="sort_edit.asp?sort_id=<%=rs("sort_id")%>"><%=rs("sname")%></a>&nbsp;</td>
    <td bgcolor="#FFFFFF"><%=rs("memo")%>&nbsp;</td>
  </tr>
  <%
rs.movenext
loop
%>
  <tr bgcolor="#FFFFFF"> 
    <td height="35" colspan="8">&nbsp;</td>
  </tr>
</table>
<%
rs.close
set rs=nothing
%>
</body>
</html>
