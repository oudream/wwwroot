<html>
<head>
<title>SYSTEM MANAGER</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="50">&nbsp;</td>
  </tr>
</table>

<!--#include file="adminconn.asp" -->
<%
sql="select * from field order by fid desc" 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1 

if rs.eof or rs.bof then response.Redirect("error.asp?err=003")

if not isempty(request("page")) then   
pagecount=cint(request("page"))   
else   
pagecount=1
end if

rs.pagesize=5
rs.AbsolutePage=pagecount
%>
<table border="0" cellspacing="1" cellpadding="2" bgcolor="#000000" width="97%" align="center">
  <tr align="center" bgcolor="#FFFFFF"> 
    <td height="22" colspan="6"><strong><font color="#000000" size="3" face="Verdana, Arial, Helvetica, sans-serif">List 
      of Field</font></strong> </td>
  </tr>
  <%
do while not rs.eof 
%>
  <tr> 
    <td width="10%" height="22" bgcolor="#FFFFFF"><img src=<%="../"&rs("photo")%> border="0" width="200" height="120"></td>
    <td colspan="4" bgcolor="#FFFFFF"> <table width="100%" border="1" cellspacing="2" cellpadding="2">
        <tr> 
          <td width="50%" height="22"><%=rs("name")%></td>
          <td width="50%">Night -- <%=rs("tfnight")%></td>
        </tr>
        <tr> 
          <td height="22" colspan="2"><%=rs("location")%></td>
        </tr>
        <tr align="left" valign="top"> 
          <td height="100" colspan="2">NOTE : <%=rs("memo")%></td>
        </tr>
      </table></td>
    <td width="24%" align="center" bgcolor="#FFFFFF"><a href="field_e.asp?page=<%=pagecount%>&fid=<%=rs("fid")%>">[ 
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
    <td height="35" colspan="6"> <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr> 
          <td> <table border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <%
zj=1
for zi=1 to rs.pagecount
%>
                <td width="50"><a href="field_v.asp?page=<%=zj%>">|&nbsp;<%=zj%>&nbsp;|</a></td>
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
    <td height="35" colspan="6"> <a href="field_c.asp?page=<%=pagecount%>">[ 
      ADD ]</a> </td>
  </tr>
</table>
<%
rs.close
set rs=nothing
%>
</body>
</html>
