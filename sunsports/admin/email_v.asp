<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="css.css" type=text/css rel=stylesheet>
</head>
<body text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!--#include file="adminconn.asp" -->
<%
if trim(request("option"))="del" then
conn.execute "delete * from mail where id=" & request("id")
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Del Email is complete. ');</SCRIPT>")
end if
%>
<table width="100%" border="0" cellspacing="2" cellpadding="2">
  <tr> 
    <td>&nbsp;</td>
    <td><table width="100%" border="0" cellspacing="2" cellpadding="2">
        <tr> 
          <td height="10">&nbsp;</td>
        </tr>
        <tr> 
          <td align="center" class="header"><h2><font color="#000066"><br>
              Email View<br>
              </font></h2></td>
        </tr>
      </table></td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td> 
<!-- ========================================================================================================		  

													cup table ============star

 ======================================================================================================== -->
      <table width="100%" bordercolor="#CCCCCC" border="1" cellpadding="2" cellspacing="0">
        <tr> 
          <td width="100" align="center">NAME</td>
          <td width="100" align="center">EMAIL</td>
          <td width="100" align="center">SUBJECT</td>
          <td width="200" align="center">CONTENT</td>
          <td colspan="2" align="center">Edit/Del</td>
        </tr>
        <%
sql="select * from mail order by id desc" 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.Open sql,conn,1,1 
if rs.bof or rs.eof then response.Redirect("error.asp?err=1")
do while not rs.eof
%>
        <tr> 
          <td height="19" align="center"><%=rs("name")%>&nbsp;</td>
          <td align="center"><%=rs("email")%>&nbsp;</td>
          <td align="center"><%=rs("subject")%>&nbsp;</td>
          <td align="center"> <%
if len(rs("mailtext"))>32 then
response.Write(left(rs("mailtext"),30)&"..")
else
response.Write(rs("mailtext"))
end if
%> &nbsp;</td>
          <td width="30" align="center"><a href="email_show.asp?id=<%=rs("id")%>" title="<%=rs("subject")%>">Detail</a></td>
          <td width="30" align="center"><a href="email_v.asp?id=<%=rs("id")%>&option=del"% onClick="return confirm(' Are you sure to del the Email ?')">Del</a></td>
        </tr>
        <%
rs.movenext
loop
rs.close
set rs=nothing
%>
        <tr> 
          <td>&nbsp;</td>
          <td>&nbsp;</td>
          <td>&nbsp;</td>
          <td>&nbsp;</td>
          <td>&nbsp;</td>
          <td>&nbsp;</td>
        </tr>
      </table>
      <!-- ========================================================================================================		  

													cup table ============stop

 ======================================================================================================== -->
    </td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
</table>
</body>
</html>
