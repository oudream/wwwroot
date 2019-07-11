<!--#include file="mymefaq5411jkjkh/favorible/showme.asp"-->
<%on error resume next
%>
<script language="JavaScript">
function CheckForm()
{

	if (document.modiinfo.pwdold.value.length == 0) {
		alert("请输入旧密码.");
		document.modiinfo.pwdold.focus();
		return false;
	}
	if (document.modiinfo.pwd.value.length == 0) {
		alert("请输入新密码.");
		document.modiinfo.pwd.focus();
		return false;
	}
	if (document.modiinfo.pwd.value != document.modiinfo.PW_Again.value) {
		alert("您两次输入的密码不一样！请重新输入.");
		document.modiinfo.pwd.focus();
		return false;
	}
    if (document.modiinfo.pwd.value.length < 6)
  {
    	alert("密码不能少于6个字符！")
    	document.modiinfo.pwd.focus()
    	return false
  }
   if (document.modiinfo.pwd.value.length > 20)
  {
    	alert("密码不能超过20个字符！")
    	document.modiinfo.pwd.focus()
    	return false
  }      
	
	return true;
}
</script>
<%
uid=request("uid")
pwdold=request("oldpass")
if trim(pwdold)="" or trim(uid)="" then
response.Redirect("index.asp")
response.End()
end if 
%>
<%
if trim(request("errorstr"))="yes" then
%>
<script language="JavaScript">
alert("你输入的旧密码错误，请重新输入")
</script>
<%
nowtime=now()
end if
%>
<html>
<head>
<title>信网 数信科技 - 密码修改</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="css.css" type="text/css">

</head>

<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0">
<br>
<br>
<br>
<table width="750" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#000000">
  <tr> 
    <td valign="top" bgcolor="#FFFFFF"> <div align="right"><br>
        <table width="563" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#000000">
          <tr> 
            <td height="5" bgcolor="#FFFFFF"></td>
          </tr>
          <tr> 
            <td height="35" bgcolor="#FFFFFF"><p
                         align="center" style="margin-top:3px;"> 
              <div align="center">=== 忘记密码 - 修改面板 ===</div> </td>
          </tr>
          <tr> 
            <td bgcolor="#FFFFFF"><br> &nbsp;&nbsp;&nbsp;&nbsp; <br> <br>
              &nbsp;&nbsp;&nbsp;&nbsp; <p
                         align="center" style="margin-top:3px;">
              </td>
          </tr>
        </table>
        <div align="center"><br>
        </div>
        <table width="563" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr> 
            <td colspan="2" height="1"></td>
          </tr>
          <tr> 
            <td width="1"></td>
            <td width="562"> <table width="100%" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#000000">
                <%
Set rs=Server.CreateObject("ADODB.RecordSet") 
sql="select * from user where uid='"&uid&"'"
rs.open sql,conn,1,1
%>
                <form name="modiinfo" method="POST" action="usermodpassok11.asp" onSubmit="return CheckForm();">
                  <tr> 
                    <td width="27%" height="30" bgcolor="#FFFFFF">用户帐号：</td>
                    <td width="51%" bgcolor="#FFFFFF">&nbsp; <%=rs("uid")%></td>
                    <td width="22%" bgcolor="#FFFFFF">&nbsp; </td>
                  </tr>
                  <tr bgcolor="#FFFFFF"> 
                    <td height="31" colspan="3"><div align="center">请填写新的密码&nbsp; 
                        <input type="hidden" name="pwdold" value="<%=pwdold%>">
                        <input type="hidden" name="pwdcon" value=<%=rs("pwd")%>>
                        <input type="hidden" name="uid" value=<%=rs("uid")%>>
                      </div></td>
                  </tr>
                  <br>
                  <br>
                  <tr> 
                    <td height="28" bgcolor="#FFFFFF">密码： </td>
                    <td height="28" bgcolor="#FFFFFF"> &nbsp; <input type="password" name="pwd" maxlength="21" class="form" size="21"></td>
                    <td height="28" bgcolor="#FFFFFF">&nbsp;</td>
                  </tr>
                  <tr> 
                    <td width="27%" height="30" bgcolor="#FFFFFF">重复密码：</td>
                    <td bgcolor="#FFFFFF">&nbsp; <input type="password" name="PW_Again" maxlength="21" class="form" size="21"></td>
                    <td width="22%" bgcolor="#FFFFFF">&nbsp; </td>
                  </tr>
                  <tr> 
                    <td width="27%" height="30" bgcolor="#FFFFFF">&nbsp;</td>
                    <td bgcolor="#FFFFFF">&nbsp;</td>
                    <td width="22%" bgcolor="#FFFFFF">&nbsp; </td>
                  </tr>
                  <tr bgcolor="#FFFFFF"> 
                    <td height="30" colspan="3"><div align="center"> 
                        <input type="submit" name="ok" value="修 改">
                        &nbsp;&nbsp; 
                        <input type="reset" name="Reset" value="清 除">
                      </div></td>
                  </tr>
                </form>
              </table></td>
          </tr>
          <%
rs.close
set rs=nothing
%>
          <tr> 
            <td colspan="2" height="1"></td>
          </tr>
        </table>
        <br>
        <br>
        <br>
      </div></td>
  </tr>
</table>
</body>
</html>
