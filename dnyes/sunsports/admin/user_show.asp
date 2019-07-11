<!--#include file="adminconn.asp" -->
<%
sql="select * from player where pid="&request("pid")
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1
%>
	
<table width="90%" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#6184A3">
  <tr align="center" bgcolor="#FFFFFF"> 
    <td height="30" colspan="2" ><b>SIGN UP COMPLETE</b></td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <td width="28%" height="30">User Level:</td>
    <td width="72%"> <%
   	Select Case rs("adminlevel")
           Case 0 response.Write("player")
           Case 1 response.Write("administrator")
           Case 2 response.Write("manager")
           Case 3 response.Write("visitor")
           Case 4 response.Write("shopping user")
           Case 5 response.Write("asst.manager")
      End Select
		  %>
      &nbsp; </td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <td width="28%" height="30">Reg Time:</td>
    <td width="72%"><%=rs("inserttime")%>&nbsp;</td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <td width="28%" height="30">User Nric:</td>
    <td width="72%"><%=rs("nric")%>&nbsp;</td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <td width="28%" height="30">Jersey:</td>
    <td width="72%"><%=rs("jersey")%>&nbsp;</td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <td width="28%" height="30">Team id:</td>
    <td width="72%"><%=rs("tid")%>&nbsp;</td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <td width="28%" height="30">User id:</td>
    <td width="72%"><%=rs("uid")%>&nbsp;</td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <td height="30">Password:</td>
    <td><%=rs("pwd")%>&nbsp;</td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <td height="30">Full Name:</td>
    <td><%=rs("name")%>&nbsp;</td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <td height="30">Telphone:</td>
    <td><%=rs("tel")%>&nbsp;</td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <td height="30">Email:</td>
    <td><%=rs("email")%>&nbsp;</td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <td height="30">ZIP:</td>
    <td><%=rs("zip")%>&nbsp;</td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <td height="30">Contact:</td>
    <td><%=rs("contact")%>&nbsp;</td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <td height="30" colspan="2">&nbsp;</td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <td height="30" colspan="2"><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td height="36" align="center">&nbsp;</td>
          <td width="50%" align="center"><a href="Shopping_Order_V.asp?page=<%=request("page")%>">go back</a></td>
        </tr>
      </table></td>
  </tr>
</table>
<%
rs.close
set rs=nothing
%>
