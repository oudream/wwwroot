<!--#include file="Top.asp" -->
<html>
<head>
<title>SYSTEM MANAGER</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body bgcolor="#FFFFCC" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="50">&nbsp;</td>
  </tr>
</table>

<%
adminlevel=request("adminlevel")
if not ( adminlevel=1 or adminlevel=2 or adminlevel=3 ) then response.Redirect("error.asp?err=001")
if adminlevel=1 then 
sql="select * from user where adminlevel=1 order by id desc"
admin_str="Administrators"
end if
if adminlevel=2 then 
sql="select * from player where adminlevel=2 order by id desc" 
admin_str="Managers"
end if
if adminlevel=3 then 
sql="select * from user where adminlevel=3 order by id desc" 
admin_str="Visitors"
end if

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
    <td height="22" colspan="7"><strong><font color="#000000" size="3" face="Verdana, Arial, Helvetica, sans-serif">List 
      of <%=admin_str%></font></strong> </td>
  </tr>
  <tr> 
    <td width="10%" height="22" bgcolor="#FFCC33">Userid</td>
    <td width="13%" bgcolor="#FFCC33">Password</td>
    <td width="10%" bgcolor="#FFCC33">Name</td>
    <td width="15%" bgcolor="#FFCC33">Contact</td>
    <td bgcolor="#FFCC33">Email</td>
    <td bgcolor="#FFCC33">Nric/Passport NO.</td>
    <td bgcolor="#FFCC33">OPERATION</td>
  </tr>
  <%
do while not rs.eof 
%>
  <tr> 
    <td width="10%" height="22" bgcolor="#FFFF99"><%=rs("uid")%></td>
    <td width="13%" bgcolor="#FFFF99"><%=rs("pwd")%></td>
    <td width="10%" bgcolor="#FFFF99"><%=rs("name")%></td>
    <td width="15%" bgcolor="#FFFF99"><%=rs("contact")%></td>
    <td width="10%" bgcolor="#FFFF99"><%=rs("email")%></td>
    <td bgcolor="#FFFF99"><%=rs("nric")%></td>
    <td bgcolor="#FFFF99"><a href="User_e.asp?id=<%=rs("id")%>&page=<%=pagecount%>&adminlevel=<%=rs("adminlevel")%>">[ 
      EDIT | DEL ]</a> </td>
  </tr>
  <%
rs.movenext
i=i+1                                                                     
if i>=rs.pagesize then exit do                                                           
loop
rs.movefirst
%>
  <tr bgcolor="#FFFFCC"> 
    <td height="35" colspan="7"> <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td> <table border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <%
zj=1
for zi=1 to rs.pagecount
%>
                <td width="50"><a href="User_v.asp?page=<%=zj%>&adminlevel=<%=rs("adminlevel")%>">|&nbsp;<%=zj%>&nbsp;|</a></td>
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
  <tr align="center" bgcolor="#FFFFCC">
    <td height="35" colspan="7">
<a href="User_c.asp?page=<%=pagecount%>&adminlevel=<%=rs("adminlevel")%>">[ ADD ]</a> 
	</td>
  </tr>
</table>
<%
rs.close
set rs=nothing
%>
</body>
</html>
