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
sql="select * from user where adminlevel=1 or adminlevel=3 order by id desc" 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1 

if rs.eof or rs.bof then 
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' None data ! ');</SCRIPT>")
response.End()
end if
%>
<table width="97%" border="1" align="center"  cellpadding="2" cellspacing="0" bordercolor="#CCCCCC">
  <tr bgcolor="#FFFFFF"> 
    <td height="22" colspan="4"  class="header"> </td>
  </tr>
  <tr align="center" bgcolor="#FFFFFF"> 
    <td height="22" colspan="4"  class="header">管理员列表</td>
  </tr>
  <tr> 
    <td width="200" height="30" bgcolor="#FFFFFF">姓名 -- 地址 -- 电话 -- 传真</td>
    <td width="150"  bgcolor="#FFFFFF">邮箱 -- QQ -- MSN</td>
    <td width="250"  bgcolor="#FFFFFF">备注</td>
  </tr>
  <%
do while not rs.eof 
%>
  <tr> 
    <td width="200" bgcolor="#FFFFFF">&nbsp;&nbsp;&nbsp;&nbsp; <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td><%=rs("Uname")%></td>
        </tr>
        <tr> 
          <td><%=rs("addr")%></td>
        </tr>
        <tr> 
          <td><%=rs("tel2")%></td>
        </tr>
        <tr> 
          <td><%=rs("tel3")%></td>
        </tr>
        <tr> 
          <td><%=rs("fax")%></td>
        </tr>
      </table></td>
    <td width="150"  bgcolor="#FFFFFF">&nbsp; 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><%=rs("email")%></td>
        </tr>
        <tr>
          <td><%=rs("qq")%></td>
        </tr>
        <tr>
          <td><%=rs("msn")%></td>
        </tr>
        <tr>
          <td>&nbsp;</td>
        </tr>
        <tr>
          <td>&nbsp;</td>
        </tr>
      </table></td>
    <td width="250"  bgcolor="#FFFFFF"><%=rs("memo")%>&nbsp;</td>
  </tr>
  <%
rs.movenext
i=i+1                                                                     
if i>=rs.pagesize then exit do                                                           
loop
rs.movefirst
%>
  <tr bgcolor="#FFFFFF"> 
    <td height="35" colspan="4">&nbsp;</td>
  </tr>
</table>
<%
rs.close
set rs=nothing
%>
</body>
</html>
