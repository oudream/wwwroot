<!--#include file="mymefaq5411jkjkh/favorible/showme.asp"-->
<!--#include file="css.asp"-->
<%on error resume next%>
<%
id=request("id")
sql="select * from myfaq where id="&id
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1
newstype=rs("newstype")
subject=rs("subject")
message=rs("message")
message=replace(message,chr(13),"<br>")
idate=rs("idate")
rs.close
set rs=nothing
%>
<html>
<head>
<title><%=Application("y_it_itname")%>-<%=newstype%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0">
<table width="413" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="413" height="149" valign="top" background="images/413x100.gif"><table width="100%" height="133" border="0" cellpadding="5" cellspacing="0">
        <tr> 
          <td height="123" valign="bottom">首页&gt;&gt;相关知识</td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td height="20">
            <div align="center">=== <%=subject%> ===<br>
            </div>
            <table width="96%" align="center" border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td bgcolor="#f0f0f0"><img src="%D2%DA%C1%F7%CD%F8%C2%E7%20-%20DNS99_CN.files/#" width="1" height="1"></td>
              </tr>
            </table></td>
        </tr>
      </table>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td rowspan="2" width="1"></td>
          <td> <table width="100%" border="0" cellspacing="1" cellpadding="5" align="center">
              <tr> 
                <td class="ptxt"><br> &nbsp;&nbsp;&nbsp;&nbsp;<%=message%> </td>
              </tr>
            </table>
            <br> <br> </td>
        </tr>
        <tr> 
          <td height="20">&nbsp; <div align="right"><%=Application("y_it_itname")%>&nbsp;&nbsp;<%=idate%>&nbsp;&nbsp;&nbsp;</div></td>
        </tr>
        <tr> 
          <td colspan="2" height="1"></td>
        </tr>
      </table> </td>
  </tr>
</table>
</body>
</html>
