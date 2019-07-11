<script language="JavaScript">
<!--
function gettarg(targ,selObj,restore){ //v3.0
  eval(targ+".location='team_v.asp?cid="+selObj.options[selObj.selectedIndex].value+"'");
  if (restore) selObj.selectedIndex=0;
}
//-->
</script>
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="css.css" type=text/css rel=stylesheet>
</head>
<body text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">

<body>
<!--#include file="adminconn.asp" -->
<%
if trim(request("option"))="del" then



dim aplayer()
psql="select pid from player where tid=" & request("tid")
Set prs=Server.CreateObject("ADODB.RecordSet") 
prs.Open psql,conn,1,1 
if not ( prs.eof or prs.bof ) then
redim aplayer(prs.recordcount-1)
for i=0 to ubound(aplayer)
aplayer(i)=prs("pid")
prs.movenext
next
prs.close
set prs=nothing
else
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' There are not player in The League(cup) ! ');</SCRIPT>")
response.End()
prs.close
set prs=nothing
end if

for i=0 to ubound(aplayer)
conn.Execute("delete * from main where pid=" & aplayer(i))
conn.Execute("delete * from announcements where pid=" & aplayer(i))
conn.Execute("delete * from card where pid=" & aplayer(i))
conn.Execute("delete * from scorer where pid=" & aplayer(i))
conn.Execute("delete * from player where pid=" & aplayer(i))
next

erase aplayer

conn.execute("delete * from card where tid=" & request("tid"))
conn.execute("delete * from scorer where tid=" & request("tid"))
conn.execute("delete * from standing where tid=" & request("tid"))
conn.execute("delete * from match where htid=" & request("tid"))
conn.execute("delete * from match where atid=" & request("tid"))
conn.execute("delete * from player where tid=" & request("tid"))
conn.execute("delete * from ct where tid=" & request("tid"))
conn.execute("delete * from team where tid=" & request("tid"))


response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Del team is complete. ');</SCRIPT>")
end if
%>

<%
sql="select * from team order by name" 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.Open sql,conn,1,1 
i=0

if not isempty(request("page")) then   
pagecount=cint(request("page"))   
else   
pagecount=1
end if

rs.pagesize=30
rs.AbsolutePage=pagecount
%>
<table width="100%" border="0" cellspacing="2" cellpadding="2">
  <tr> 
    <td>&nbsp;</td>
    <td><table width="100%" border="0" cellspacing="2" cellpadding="2">
        <tr> 
          <td height="30" align="center"><b><font color=#000066 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        size=2>Team Edit/Del :</font></b></td>
        </tr>
        <tr> 
          <td height="30" align="center" valign="middle" class="header"><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td width="79%"> <table border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr> 
                      <%
zj=1
for zi=1 to rs.pagecount
%>
                      <td width="50"><a href="team_v.asp?page=<%=zj%>">|&nbsp;<%=zj%>&nbsp;|</a></td>
                      <%
zj=zj+1
next
%>
                    </tr>
                  </table></td>
                <td width="21%" align="left"> <table width="100%" border="0" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td> <div align="center"> Record :<font color=red><b><%=rs.recordcount%></b></font> 
                          Page <font color=red><b><%=pagecount%></b></font> of 
                          <font color=red><b><%=rs.pagecount%></b></font> </div></td>
                    </tr>
                  </table></td>
              </tr>
            </table></td>
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
          <td width="30" align="center">NO.</td>
          <td width="150" align="center">NAME</td>
          <td width="130" align="center">City</td>
          <td align="center">Create Date</td>
          <td width="130" align="center">Manager</td>
          <td width="130" align="center">Asst.Manager </td>
          <td colspan="2" align="center">Edit/Del</td>
        </tr>
<%
do while not rs.eof
%>
        <tr> 
          <td height="19" align="center">
            <%=rs("tid")%>&nbsp;</td>
          <td><%=rs("name")%>&nbsp;</td>
          <td align="center"><%=rs("city")%>&nbsp;</td>
          <td align="center"><%=rs("createdate")%>&nbsp;</td>
          <td align="center">
<%
psql="select * from player where pid=" & rs("mpid")
Set prs=Server.CreateObject("ADODB.RecordSet") 
prs.Open psql,conn,1,1 
if not ( prs.bof or prs.eof ) then 
response.Write(prs("name"))
end if
prs.close
set prs=nothing
%>
		  &nbsp;</td>
          <td align="center">
<%
psql="select * from player where pid=" & rs("apid")
Set prs=Server.CreateObject("ADODB.RecordSet") 
prs.Open psql,conn,1,1 
if not ( prs.bof or prs.eof ) then 
response.Write(prs("name"))
end if
prs.close
set prs=nothing
%>
            &nbsp;</td>
          <td width="30" align="center"><a href="team_e.asp?tid=<%=rs("tid")%>" title="<%=rs("name")%>">Edit</a></td>
          <td width="30" align="center"><a href="team_v.asp?tid=<%=rs("tid")%>&option=del"% onClick="return confirm(' Are you sure to del team [ <%=rs("name")%> ] form the tournament \n If you delete it , all data about the team would be delete too !!! ?')">Del</a></td>
        </tr>
<%
rs.movenext
i=i+1                                                                     
if i>=rs.pagesize then exit do                                                           
loop
%>
      </table>
      <!-- ========================================================================================================		  

													cup table ============stop

 ======================================================================================================== -->
    </td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr> 
          <td width="79%"> <table border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <%
zj=1
for zi=1 to rs.pagecount
%>
                <td width="50"><a href="team_v.asp?page=<%=zj%>">|&nbsp;<%=zj%>&nbsp;|</a></td>
                <%
zj=zj+1
next
%>
              </tr>
            </table></td>
          <td width="21%" align="left"> <table width="100%" border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td> <div align="center"> Record :<font color=red><b><%=rs.recordcount%></b></font> 
                    Page <font color=red><b><%=pagecount%></b></font> of <font color=red><b><%=rs.pagecount%></b></font> 
                  </div></td>
              </tr>
            </table></td>
        </tr>
      </table></td>
    <td>&nbsp;</td>
  </tr>
</table>
<%
rs.close
set rs=nothing
%>
</body>
</html>
