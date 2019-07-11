<script language="JavaScript">
function checkform()
{
	if(login.uid.value.length==0) {
		alert("登录帐号不能为空！");
		login.uid.focus();
		return false;
	}
	if(login.pwd.value.length==0) {
		alert("登录密码不能为空！");
		login.pwd.focus();
		return false;
	}
}
</script>
<%
url=Request.ServerVariables("SCRIPT_NAME")
url=url & "?area=" & area & "&id=" & id & "&newsid=" & newsid
%>
                      
<table width="200" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td height="5" colspan="2"><img src="1.gif" width="1" height="1"></td>
  </tr>
  <tr> 
    <td> <%if session("y_it_uid")<>"" then%> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td height="20" width="16"> <div align="center">→</div></td>
          <td colspan="2">客户类型：<font color=red><%=session("y_it_userlevel")%></font></td>
        </tr>
        <tr> 
          <td height="20"> <div align="center">→</div></td>
          <td colspan="2">预 付 款：<font color=red><%=session("y_it_prepay")%> RMB</font></td>
        </tr>
        <tr> 
          <td height="20"> <div align="center">→</div></td>
          <td width="108"><a href="viewmyorders.asp">查看定单</a></td>
          <td width="108"><a href="domainmanage.asp" target="_blank">域名管理</a></td>
        </tr>
        <tr> 
          <td height="20"> <div align="center">→</div></td>
          <td><a href="usermodpass.asp">修改密码</a></td>
          <td><a href="usermod.asp">修改资料</a></td>
        </tr>
        <tr> 
          <td height="20"> <div align="center">→</div></td>
          <td><a href="login.asp?logout=yes">退出登陆</a></td>
          <td>&nbsp;</td>
        </tr>
      </table>
      <%else%> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <form action="login.asp" method="post" name="login" onSubmit="return checkform();">
          <tr> 
            <td height="22" align="center"> 会员： 
              <input maxlength=30 name=uid size=16 class="input"> </td>
          </tr>
          <tr> 
            <td height="32" align="center"> 密码： 
              <input maxlength=16 name=pwd size=16 type=password class="input"> 
            </td>
          </tr>
          <tr align="center"> 
            <td> <table width="100%" border="0" cellpadding="0" cellspacing="0">
                <tr> 
                  <td height="19" align="center"> <input type="hidden" name="url" value=<%=url%>> 
                    <input type="image" border="0" name="imageField2" src="images/login.gif" width="64" height="19"> 
                    <table width="75%" border="0" cellpadding="0" cellspacing="0">
                      <tr> 
                        <td height="1"><img src="1.gif" width="1" height="1"></td>
                      </tr>
                    </table></td>
                  <td height="19" align="center"> <a href="callmypass.asp"><img src="images/forgetpassword.gif" width="87" height="19" border="0"></a> 
                  </td>
                </tr>
              </table></td>
          </tr>
          <tr align="center"> 
            <td height="30"> <a href="userregprotocal.asp"><img src="images/reg.gif" width="146" height="19" border="0"></a>&nbsp;&nbsp;&nbsp;</td>
          </tr>
        </form>
      </table>
      <%end if%> </td>
  </tr>
</table>
<br>