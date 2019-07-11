<%
if session("user_uid")<>"" then
%>
<table width="100%" border="0" cellspacing="2" cellpadding="2">
  <tr>
    <td height="429" align="center" valign="middle"><table width="100%" border="0" cellspacing="2" cellpadding="2">
        <tr>
          <td align="center">Welcome team manager , you have logined !</td>
        </tr>
      </table></td>
  </tr>
</table>
<%
if session("user_adminlevel")=1 then
%>
<table width="100%" border="0" cellspacing="2" cellpadding="2">
  <tr>
    <td height="429" align="center" valign="middle"><table width="100%" border="0" cellspacing="2" cellpadding="2">
        <tr>
          <td align="center">your level is administrator !</td>
        </tr>
      </table></td>
  </tr>
</table>
<%
elseif session("user_adminlevel")=2 then 
%>
<table width="100%" border="0" cellspacing="2" cellpadding="2">
  <tr>
    <td height="429" align="center" valign="middle"><table width="100%" border="0" cellspacing="2" cellpadding="2">
        <tr>
          <td align="center">your level is managment !</td>
        </tr>
      </table></td>
  </tr>
</table>
<%
else
%>
<table width="100%" border="0" cellspacing="2" cellpadding="2">
  <tr>
    <td height="429" align="center" valign="middle"><table width="100%" border="0" cellspacing="2" cellpadding="2">
        <tr>
          <td align="center">you are wrong 1 !</td>
        </tr>
      </table></td>
  </tr>
</table>
<%
end if
%>
<%
else
%>
<table width="100%" border="0" cellspacing="2" cellpadding="2">
  <tr>
    <td height="429" align="center" valign="middle"><table width="100%" border="0" cellspacing="2" cellpadding="2">
        <tr>
          <td align="center">you are wrong 2 !</td>
        </tr>
      </table></td>
  </tr>
</table>
<%
end if
%>
