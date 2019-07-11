<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="css.css" type=text/css rel=stylesheet>
</head>
<body text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="185" border="0" align="center" cellpadding="0" cellspacing="0">
<%
if session("user_uid")<>"" then
%>
<%
if session("user_adminlevel")=1 then
%>
<!-- 
   ----------------------------------------------   ----------------------------------------------------
   														logined level1																
   ----------------------------------------------   ----------------------------------------------------
-->
  <tr> 
    <td>&nbsp;</td>
	    
    <td align="center"><TABLE height=100% width=100% border=0>
        <TBODY>
          <TR> 
            <TD vAlign=top align=left> <TABLE width=100% height="30" border=0>
                <TBODY>
                  <TR> 
                    <TD width="50%" height="40"><strong><a href="left.asp" target="leftFrame"><font color="#FF0000">SPORTS</font></a></strong></TD>
                    <TD width="50%"><strong><a href="Shopping_left.asp" target="leftFrame"><font color="#0066FF">PRODUCTS</font></a></strong></TD>
                  </TR>
                </TBODY>
              </TABLE><br>
              <TABLE width=100% border=1 cellpadding="5" cellspacing="0" bordercolor="#CCCCCC" bgcolor="#f5f5f5">
                <TBODY>
                  <TR> 
                    <TD><b>Data</b></TD>
                  </TR>
                  <TR> 
                    <TD><a href="data_match_c.asp" target="mainFrame" title="Fixture for league & Cup"><font color="#FF0000">Fixture</font> 
                      (League & Cup)</a> 216</TD>
                  </TR>
                  <TR> 
                    <TD><a href="data_match_c1.asp" target="mainFrame" title="Fixture for league & Cup"><font color="#FF0000">Fixture</font> 
                      (League & Cup)</a> 230</TD>
                  </TR>
                </TBODY>
              </TABLE>
              <br>
              <TABLE width=100% border=1 cellpadding="5" cellspacing="0" bordercolor="#CCCCCC" bgcolor="#f5f5f5">
                <TBODY>
                  <TR> 
                    <TD><b>Fixture</b></TD>
                  </TR>
                  <TR> 
                    <TD><a href="match_c.asp" target="mainFrame" title="Fixture for league & Cup"><font color="#FF0000">Fixture</font> 
                      (League & Cup)</a></TD>
                  </TR>
                  <TR> 
                    <TD><a href="match_v.asp" target="mainFrame">Fixture Edit 
                      / Del</a></TD>
                  </TR>
                  <TR> 
                    <TD><a href="friendlies_v.asp" target="mainFrame" title="Fixture for friendlies">Fixture (Friendlies)
                      </a></TD>
                  </TR>
                </TBODY>
              </TABLE>
              <br>
              <TABLE width=100% border=1 cellpadding="5" cellspacing="0" bordercolor="#CCCCCC" bgcolor="#f5f5f5">
                <TBODY>
                  <TR> 
                    <TD><b>Match Reports</b></TD>
                  </TR>
                  <TR> 
                    <TD><a href="match_r_vcall.asp" target="mainFrame">View Unsubmitted Reports</a></TD>
                  </TR>
                  <TR> 
                    <TD><a href="match_r_vdeff.asp" target="mainFrame">Resolve 
                      Conflicting Reports</a></TD>
                  </TR>
                  <TR> 
                    <TD><a href="match_r_vedit.asp" target="mainFrame">View Submitted Reports</a></TD>
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
                      Edit / Del</a></TD>
                  </TR>
                </TBODY>
              </TABLE>
              <br>
              <TABLE width=100% border=1 cellpadding="5" cellspacing="0" bordercolor="#CCCCCC" bgcolor="#f5f5f5">
                <TBODY>
                  <TR> 
                    <TD><strong>Fields Management</strong></TD>
                  </TR>
                  <TR> 
                    <TD><a href="field_book_c.asp" target="mainFrame"><font color="#FF0000">Field 
                      Booking</font></a></TD>
                  </TR>
                  <TR> 
                    <TD> <a href="field_book_v.asp" target="mainFrame">Field Booking 
                      Edit / Del</a></TD>
                  </TR>
                  <TR>
                    <TD><a href="field_c.asp" target="mainFrame">Field Add</a></TD>
                  </TR>
                  <TR>
                    <TD> <a href="field_v.asp" target="mainFrame">Field Edit / 
                      Del</a></TD>
                  </TR>
                </TBODY>
              </TABLE>
              <br>
              <TABLE width=100% border=1 cellpadding="5" cellspacing="0" bordercolor="#CCCCCC" bgcolor="#f5f5f5">
                <TBODY>
                  <TR> 
                    <TD><strong>Teams Management</strong></TD>
                  </TR>
                  <TR> 
                    <TD><a href="team_v.asp" target="mainFrame">Team Edit / Delete</a></TD>
                  </TR>
                  <TR> 
                    <TD><a href="team_c.asp" target="mainFrame">Team Add</a></TD>
                  </TR>
                  <TR> 
                    <TD> <a href="team_table.asp" target="mainFrame">Team's Table 
                      </a></TD>
                  </TR>
                  <TR> 
                    <TD><a href="team_t.asp" target="mainFrame">Team's Player 
                      </a></TD>
                  </TR>
                  <TR> 
                    <TD>&nbsp;</TD>
                  </TR>
                </TBODY>
              </TABLE>
              <br>
              <TABLE width=100% border=1 cellpadding="5" cellspacing="0" bordercolor="#CCCCCC" bgcolor="#f5f5f5">
                <TBODY>
                  <TR> 
                    <TD><strong>League/Cup Management</strong></TD>
                  </TR>
                  <TR> 
                    <TD><a href="league_c.asp" target="mainFrame">League Add </a></TD>
                  </TR>
                  <TR> 
                    <TD> <a href="league_v.asp" target="mainFrame">League Edit 
                      / Del</a></TD>
                  </TR>
                  <TR> 
                    <TD><a href="cup_c.asp" target="mainFrame">Cup Add </a></TD>
                  </TR>
                  <TR> 
                    <TD> <a href="Cup_v.asp" target="mainFrame">Cup Edit / Del</a></TD>
                  </TR>
                  <TR> 
                    <TD><a href="tournament_t.asp" target="mainFrame">Tournament's 
                      Team </a></TD>
                  </TR>
                </TBODY>
              </TABLE>
              <br>
              <TABLE width=100% border=1 cellpadding="5" cellspacing="0" bordercolor="#CCCCCC" bgcolor="#f5f5f5">
                <TBODY>
                  <TR> 
                    <TD><strong>Miscellaneous</strong></TD>
                  </TR>
                  <TR> 
                    <TD><a href="user_c.asp?adminlevel=1" target="mainFrame">Administrator 
                      Add </a></TD>
                  </TR>
                  <TR> 
                    <TD><a href="user_c.asp?adminlevel=3" target="mainFrame">Visitor 
                      Add </a></TD>
                  </TR>
                  <TR>
                    <TD> <a href="user_v.asp" target="mainFrame">All User Edit / Del</a></TD>
                  </TR>
                </TBODY>
              </TABLE>
              <br>
              <TABLE width=100% border=1 cellpadding="5" cellspacing="0" bordercolor="#CCCCCC" bgcolor="#f5f5f5">
                <TBODY>
                  <TR> 
                    <TD><strong>Others</strong></TD>
                  </TR>
                  <TR> 
                    <TD><a href="email_v.asp" target="mainFrame">Email View</a></TD>
                  </TR>
                  <TR> 
                    <TD><a href="events_c.asp" target="mainFrame">Events Add</a></TD>
                  </TR>
                  <TR> 
                    <TD> <a href="events_v.asp" target="mainFrame">Events Edit 
                      / Del</a></TD>
                  </TR>
                  <TR>
                    <TD><a href="color_v.asp" target="mainFrame">Colorl Management</a></TD>
                  </TR>
                </TBODY>
              </TABLE>
              
            </TD>
          </TR>
        </TBODY>
      </TABLE> </td>
        <td></td>
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
<br>
<br>
</body>
</html>