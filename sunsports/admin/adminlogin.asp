<LINK href="css.css" type=text/css rel=stylesheet>
<%
if session("user_uid")<>"" then
%>
<%
if session("user_adminlevel")=1 then
%>
<table width="100%" height="100%" border="0" cellpadding="2" cellspacing="2">
        <tr> 
          <td height="410" align="center">Welcome, <font color="#FF0000">administrator</font> 
            !You have logined</td>
        </tr>
        <tr>
          
    <td height="10" align="right" valign="bottom">&nbsp;</td>
        </tr>
</table>
<%
elseif session("user_adminlevel")=2 then 
%>
<table width="100%" height="100%" border="0" cellpadding="2" cellspacing="2">
        <tr> 
          <td height="410" align="center">Welcome, <font color="#FF0000">manager</font> 
            !You have logined</td>
        </tr>
        <tr>
          
    <td height="10" align="right" valign="bottom">&nbsp;</td>
        </tr>
</table>
<%
else
%>
<table width="100%" height="100%" border="0" cellpadding="2" cellspacing="2">
        <tr> 
          
    <td height="410" align="center">Sorry, Session time out or Your ID or password 
      that you put in is <font color="#FF0000">wrong</font> ! !<br>
      <br>
      Please login again</td>
        </tr>
        <tr>
          
    <td height="10" align="right" valign="bottom">&nbsp;</td>
        </tr>
</table>
<%
end if
%>
<%
else
%>
<table width="100%" height="100%" border="0" cellpadding="2" cellspacing="2">
        <tr> 
          
    <td height="410" align="center">Sorry, Session time out or Your ID or password 
      that you put in is <font color="#FF0000">wrong</font> ! !<br>
      <br>
      Please login again </td>
        </tr>
        <tr>
          
    <td height="10" align="right" valign="bottom">&nbsp;</td>
        </tr>
</table>
<%
end if
%>
