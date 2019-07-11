<!--#include file="dbm_conn.asp" -->
<HTML>
<HEAD>
<TITLE></TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="CSS_b.css" rel="stylesheet" type="text/css">
</HEAD>
<BODY body bgcolor=transparent leftMargin=0 
topMargin=0 marginheight="0" marginwidth="0">
<%
Set rs=Server.CreateObject("ADODB.RecordSet") 
sql="select * from dbm_hildden order by dbm_hildden_id"
rs.open sql,conn,1,1 
if rs.eof or rs.bof then 
%>
<table width="97%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="20" align="center">暂无您所要的数据!</td>
  </tr>
</table>
<%
response.End()
end if
%>
<table width="120" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="2">&nbsp;</td>
    <td height="25"><font color="#CCCCCC"> 
      <%
do while not rs.eof 
%>
      <a href="http://<%=rs("web_addr")%>" target="_blank"><%=rs("htitle")%></a> <br>
      </font>
      <table width="75%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td height="3"><img src="1.gif" width="1" height="1"></td>
        </tr>
      </table>
      <font color="#CCCCCC"> 
      <%
rs.movenext
loop
rs.close
set rs=nothing
%>
      </font>&nbsp;</td>
  </tr>
</table>
</BODY></HTML>
