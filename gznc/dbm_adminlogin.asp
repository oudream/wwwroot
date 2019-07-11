<table width="100%" height="100%" border="0" cellpadding="2" cellspacing="2">
        <tr> 
          <td height="410" align="center">您好 ： <font color="#FF0000"><%=session("user_allname")%></font> 
            您已经成功登录!</td>
        </tr>
        <tr>
          
    <td height="10" align="right" valign="bottom">&nbsp;</td>
        </tr>
</table>
<%
if session("user_adminlevel")<>1 and session("user_adminlevel")<>2 and session("user_adminlevel")<>9 then 
response.write "<script language='javascript'>window.close();</script>"
response.end
end if
%>
