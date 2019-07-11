<HTML>
<HEAD>
<TITLE></TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="CSS.css" rel="stylesheet" type="text/css">
</HEAD>
<BODY bgColor=#cccccc leftMargin=0 background=../image/bg_main_004.gif 
topMargin=0 marginheight="0" marginwidth="0">
<!--#include file="dbm_conn.asp" -->
<%
subs_id=request("subs_id")
if subs_id<>"" then
sql="select * from dbm_subs where subs_id=" & subs_id 
else
sql="select * from dbm_subs order by subs_id desc"
end if
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1
brief = replace(rs("FILEDESC"),chr(13),"<br>")
%>
<table width="660" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td rowspan="2" valign="top"> <img align="middle" src="<%=rs("sfile")%>" width="493" height="371"><br>
      <font size="+1"><strong><br>
      <%=rs("FILETITLE")%><br>
      </strong></font>      <br> <br>
      <%=brief%>       <br> <br>       <br> <br> 
	  </td>
    <td width="15" rowspan="2" valign="top">&nbsp;</td>
    <td width="152" height="40" valign="bottom"><strong>[其它作品]<br>
      <img src="../image/line.gif" width="107" height="16"></strong></td>
  </tr>
  <tr> 
    <td valign="top">
<%
rs.close
sql="select * from dbm_subs order by subs_id desc"
rs.open sql,conn,1,1
do while not rs.eof 
%>

      <a href="Works_Show.asp?subs_id=<%=rs("subs_id")%>" class="a"><%=rs("FILETITLE")%></a><br> <br>
<%
rs.movenext
loop
%>
	  &nbsp;</td>
  </tr>
</table>
<%
rs.close
set rs=nothing
%>

</BODY></HTML>
