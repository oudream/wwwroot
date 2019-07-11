<HTML>
<HEAD>
<TITLE></TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="CSS_b.css" rel="stylesheet" type="text/css">
</HEAD>
<BODY body bgcolor=transparent leftMargin=0 
topMargin=0 marginheight="0" marginwidth="0">
<TABLE cellSpacing=0 cellPadding=0 width=330 border=0>
  <TBODY>
    <TR> 
      <TD width=25></TD>
    </TR>
    <TR> 
      <TD vAlign=top align=left width=370>
<!--#include file="dbm_conn.asp" -->
<%
Set rs=Server.CreateObject("ADODB.RecordSet")
sql="select * from dbm_subs"
rs.open sql,conn,1,1 
if rs.eof or rs.bof then 
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' 没有合要求的作品 ! ');</SCRIPT>")
response.End()
end if
%>
<table width="660" border="0" cellpadding="0" cellspacing="0">
<%
do while not rs.eof 
%>
  <tr>
    <td height="22"><a href="View_Works.asp?subs_id=<%=rs("subs_id")%>" target="_parent"><%=rs("sname")%></a>&nbsp;</td>
  </tr>
  <%
rs.movenext
loop
rs.close
set rs=nothing
%></table>
	  </TD>
    </TR>
  </TBODY>
</TABLE>
</BODY></HTML>
