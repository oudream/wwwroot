<html>
<head>
<title>SYSTEM MANAGER</title>
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
sql="select * from color" 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1 

if rs.eof or rs.bof then response.Redirect("error.asp?err=003")

if not isempty(request("page")) then   
pagecount=cint(request("page"))   
else   
pagecount=1
end if

rs.pagesize=20
rs.AbsolutePage=pagecount
%>
<table width="96%" border="1" align="center" cellpadding="2" cellspacing="0" bordercolor="#CCCCCC">
  <tr align="center" bgcolor="#FFFFFF"> 
    <td height="22" colspan="4"><strong><font color="#000000" size="3" face="Verdana, Arial, Helvetica, sans-serif">Color 
      Edit/Del</font></strong> </td>
  </tr>
  <tr align="center" bgcolor="#FFFFFF"> 
    <td width="40%" height="22">Color Name</td>
    <td width="36%" height="22" colspan="2">Color</td>
    <td height="22">Operation</td>
  </tr>
  <%
do while not rs.eof 
%>
  <tr> 
    <td width="30%" height="22" align="center" bgcolor="#FFFFFF"><%=rs("colorname")%>&nbsp; </td>
    <td width="20%" bgcolor="<%=rs("color")%>">&nbsp;</td>
    <td width="20%" align="center" bgcolor="#FFFFFF"><%=rs("color")%>&nbsp;</td>
    <td width="24%" align="center" bgcolor="#FFFFFF"><a href="color_e.asp?page=<%=pagecount%>&id=<%=rs("id")%>">[ 
      EDIT | DEL ]</a> </td>
  </tr>
<%
rs.movenext
i=i+1                                                                     
if i>=rs.pagesize then exit do                                                           
loop
rs.movefirst
%>
  <tr bgcolor="#FFFFFF"> 
    <td height="35" colspan="4"> <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr> 
          <td> <table border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <%
zj=1
for zi=1 to rs.pagecount
%>
                <td width="50"><a href="color_v.asp?page=<%=zj%>">|&nbsp;<%=zj%>&nbsp;|</a></td>
                <%
zj=zj+1
next
%>
              </tr>
            </table></td>
          <td width="20%" align="left"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td> <div align="center"> Record :<font color=red><b><%=rs.recordcount%></b></font> Page <font color=red><b><%=pagecount%></b></font> of <font color=red><b><%=rs.pagecount%></b></font> </div></td>
              </tr>
            </table></td>
        </tr>
      </table></td>
  </tr>
  <tr align="center" bgcolor="#FFFFFF"> 
    <td height="35" colspan="4"> <a href="color_c.asp">[ 
      ADD ]</a> </td>
  </tr>
</table>
<%
rs.close
set rs=nothing
%>
</body>
</html>
