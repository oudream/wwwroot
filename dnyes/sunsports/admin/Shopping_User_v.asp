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
pid=request("pid")
if request("option")="del" then
conn.execute "delete * from player where pid=" & pid
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Del Shopping User is complete. ');</SCRIPT>")
end if
%>
<%
sql="select * from player where adminlevel=4 order by pid desc" 
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
<table width="97%" border="1" align="center" cellpadding="2" cellspacing="0" bordercolor="#CCCCCC">
  <tr align="center" bgcolor="#FFFFFF"> 
    <td height="22" colspan="8" class="header">List of <font color="#FF0000">Shopping 
      User</font></td>
  </tr>
  <tr> 
    <td width="7%" height="22" align="center" bgcolor="#FFFFFF">Userid</td>
    <td width="11%" align="center" bgcolor="#FFFFFF">Password</td>
    <td width="15%" align="center" bgcolor="#FFFFFF">Name</td>
    <td width="11%" align="center" bgcolor="#FFFFFF">Telphone</td>
    <td width="8%" align="center" bgcolor="#FFFFFF">Zip</td>
    <td width="13%" align="center" bgcolor="#FFFFFF">Contact</td>
    <td align="center" bgcolor="#FFFFFF">Email</td>
    <td width="16%" align="center" bgcolor="#FFFFFF">OPERATION</td>
  </tr>
  <%
do while not rs.eof 
%>
  <tr> 
    <td width="7%" height="22" align="center" bgcolor="#FFFFFF"><%=rs("uid")%>&nbsp;</td>
    <td width="11%" align="center" bgcolor="#FFFFFF"><%=rs("pwd")%>&nbsp;</td>
    <td width="15%" align="center" bgcolor="#FFFFFF"><%=rs("name")%>&nbsp;</td>
    <td align="center" bgcolor="#FFFFFF"><%=rs("tel")%>&nbsp;</td>
    <td align="center" bgcolor="#FFFFFF"><%=rs("zip")%>&nbsp;</td>
    <td width="13%" align="center" bgcolor="#FFFFFF"><%=rs("contact")%>&nbsp;</td>
    <td width="19%" align="center" bgcolor="#FFFFFF"><%=rs("email")%>&nbsp;</td>
    <td align="center" bgcolor="#FFFFFF"> <a href="Shopping_User_e.asp?pid=<%=rs("pid")%>"> 
      EDIT </a> &nbsp; &nbsp; <a href="Shopping_User_v.asp?option=del&pid=<%=rs("pid")%>" onClick="return confirm('Are you sure del it ?')"> 
      DEL </a> </td>
  </tr>
  <%
rs.movenext
i=i+1                                                                     
if i>=rs.pagesize then exit do                                                           
loop
rs.movefirst
%>
  <tr bgcolor="#FFFFFF"> 
    <td height="35" colspan="8"> <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr> 
          <td> <table border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <%
zj=1
for zi=1 to rs.pagecount
%>
                <td width="50"><a href="Shopping_User_v.asp?page=<%=zj%>">|&nbsp;<%=zj%>&nbsp;|</a></td>
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
</table>
<%
rs.close
set rs=nothing
%>
</body>
</html>
