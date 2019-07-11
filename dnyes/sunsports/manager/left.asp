<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="css.css" type=text/css rel=stylesheet>
</head>
<body text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!--#include file="adminconn.asp" -->
<table width="180" border="0" align="center" cellpadding="0" cellspacing="0">
<%
if session("user_uid")<>"" then
%>
<%
psql="select * from player where uid='" & session("user_uid") & "'"
set prs=server.createobject("ADODB.Recordset")
prs.open psql,conn,1,1
if not ( prs.eof or prs.bof ) then
if isnull(prs("tid")) then
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' You have not joined the team yet . ');</SCRIPT>")
response.End()
end if
session("manager_pid")=prs("pid")
session("manager_tid")=prs("tid")
end if
prs.close
set prs=nothing
%>
<%
tsql="select * from team where tid=" & session("manager_tid")
set trs=server.createobject("ADODB.Recordset")
trs.open tsql,conn,1,1
if not ( trs.eof or trs.bof ) then
session("manager_tname")=trs("name")
end if
trs.close
set trs=nothing
%>
<%
if session("user_adminlevel")=2 then
%>
<!-- 
   ----------------------------------------------   ----------------------------------------------------
   														logined level1																
   ----------------------------------------------   ----------------------------------------------------
-->
  <tr> 
    <td width="1">&nbsp;</td>
    <td width="172" align="center"><TABLE height=100% width=168 border=0>
        <TBODY>
          <TR> 
            <TD vAlign=top align=left> <TABLE width=100% height="30" border=0>
                <TBODY>
                  <TR> 
                    <TD height="40" align="center" valign="middle"><strong>Manager</strong></TD>
                  </TR>
                </TBODY>
              </TABLE>
              <TABLE width=100% border=1 cellpadding="5" cellspacing="0" bordercolor="#CCCCCC" bgcolor="#f5f5f5">
                <TBODY>
                  <TR> 
                    <TD><strong>Team Management</strong></TD>
                  </TR>
                  <TR> 
                    <TD> <a href="team_e.asp" target="mainFrame">Team Edit</a></TD>
                  </TR>
                  <TR> 
                    <TD> <a href="team_table_e.asp" target="mainFrame">Team's 
                      Table Edit</a></TD>
                  </TR>
                  <TR> 
                    <TD><a href="team_t.asp" target="mainFrame">Team's 
                      Player Edit</a></TD>
                  </TR>
                </TBODY>
              </TABLE>
              <br>
              <TABLE width=100% border=1 cellpadding="5" cellspacing="0" bordercolor="#CCCCCC" bgcolor="#f5f5f5">
                <TBODY>
                  <TR> 
                    <TD><strong>Match Confirm &amp; Reports</strong></TD>
                  </TR>
                  <TR> 
                    <TD> <a href="match_no.asp" target="mainFrame">Match <font color="#FF0000">Confirm</font></a></TD>
                  </TR>
                  <TR>
                    <TD> <a href="match_r_no.asp" target="mainFrame"> Match <font color="#FF0000">Report</font></a></TD>
                  </TR>
                </TBODY>
              </TABLE>
              <br>
              <TABLE width=100% border=1 cellpadding="5" cellspacing="0" bordercolor="#CCCCCC" bgcolor="#f5f5f5">
                <TBODY>
                  <TR> 
                    <TD width=118 height="25"><strong>Announcements</strong></TD>
                  </TR>
                  <TR> 
                    <TD height="25"><a href="announcements_c.asp" target="mainFrame">Announcements 
                      Add </a></TD>
                  </TR>
                  <TR>
                    <TD height="25"> <a href="announcements_v.asp" target="mainFrame">Announcements 
                      Edit/Del</a></TD>
                  </TR>
                </TBODY>
              </TABLE>
              
            </TD>
          </TR>
        </TBODY>
      </TABLE> </td>
        <td width="10"></td>
  </tr>
<%
end if
%>
<!-- 
   ----------------------------------------------   ----------------------------------------------------
   														loginout																
   ----------------------------------------------   ----------------------------------------------------
-->
  <tr>
    <td>&nbsp;</td>
    <td height="50" align="center" valign="middle">
          <FORM action=login.asp method="post" id="form_logout" name="form_logout" target="_parent">
            <INPUT type=submit value="logout" id="option" name="option">
          </FORM>
	</td>
    <td>&nbsp;</td>
  </tr>
<%
else response.Redirect("error.asp?../err=001")
end if
%>
</table>
</body>
</html>