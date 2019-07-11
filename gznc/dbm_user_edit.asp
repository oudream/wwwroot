<script language="JavaScript">
function checkform()
{
	if (form_administrator.allname.value.length == 0) {
		alert("请填写用户名! ");
		form_administrator.allname.focus();
		return false;
		}
	if(form_administrator.email.value.length > 1) {
	if(! isMail(form_administrator.email.value)) {
		alert('请填入正确的E-mail地址!');
		form_administrator.email.focus();
		return false;
		}}
	if (form_administrator.uid.value.length == 0) {
		alert("请填写用户账号! ");
		form_administrator.uid.focus();
		return false;
		}
	if (form_administrator.pwd.value.length == 0) {
		alert(" 请填写密码! ");
		form_administrator.pwd.focus();
		return false;
		}
	if (form_administrator.pwd.value != form_administrator.cpwd.value) {
		alert(" 两次输入密码不相符，请从新输入 ");
		form_administrator.pwd.focus();
		return false;
		}
	return true;
}

function isLetter(name)
{
var letter=" _0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"
	if(name.length == 0)
		return false;
	for(i = 0; i < name.length; i++) {
		if(letter.indexOf(name.charAt(i)) == -1)
			return false;
	}
	return true;
}

function isEnglish(name)
{
	for(i = 0; i < name.length; i++) {
		if(name.charCodeAt(i) > 128)
			return false;
	}
	return true;
}

function isMail(name)
{
	if(! isEnglish(name))
		return false;
	i = name.indexOf("@");
	j = name.lastIndexOf("@");
	if(i == -1)
		return false;
	if(i != j)
		return false;
	if(i == name.length)
		return false;
	return true;
}

</script>
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="dbm_css.css" type=text/css rel=stylesheet>
</head>
<body text="#000000" leftmargin="8" topmargin="8" marginwidth="0" marginheight="0">


<!--#include file="dbm_adminconn.asp" -->


<%
user_id=request("user_id")
if request("option")="edit" then
allname=trim(request("allname"))
allname=replace(allname,"'","''")
addr=trim(request("addr"))
addr=replace(addr,"'","''")
tel2=trim(request("tel2"))
tel2=replace(tel2,"'","''")
tel3=trim(request("tel3"))
tel3=replace(tel3,"'","''")
email=trim(request("email"))
email=replace(email,"'","''")
qq=trim(request("qq"))
qq=replace(qq,"'","''")
msn=trim(request("msn"))
msn=replace(msn,"'","''")
uid=trim(request("uid"))
uid=replace(uid,"'","''")
pwd=trim(request("pwd"))
pwd=replace(pwd,"'","''")
memo=trim(request("memo"))
memo=replace(memo,"'","''")



usql="select * from dbm_user where user_id<>"&user_id&" and ( uid='"&uid&"' or allname='"&allname&"' )"
set urs=server.createobject("ADODB.Recordset")
urs.open usql,conn,1,1
if not ( urs.eof or urs.bof ) then
response.Write("<SCRIPT LANGUAGE=JavaScript> alert('出错！此用户有同名存在，请用其它名字输入 ');</SCRIPT>")
urs.close
set urs=nothing
else
urs.close
set urs=nothing


if addr="" then addr=" "
if tel2="" then tel2=" "
if tel3="" then tel3=" "
if qq="" then qq=" "
if msn="" then msn=" "
if email="" then email=" "
if memo="" then memo=" "
if allname<>"" then
conn.Execute("update dbm_user set allname='"&allname&"',addr='"&addr&"',tel2='"&tel2&"',tel3='"&tel3&"',email='"&email&"',qq='"&qq&"',msn='"&msn&"',uid='"&uid&"',pwd='"&pwd&"',memo='"&memo&"' where user_id=" & user_id)
response.Write("<SCRIPT LANGUAGE=JavaScript> alert('编辑成功');</SCRIPT>")
end if
end if
end if
%>
<%
sql="select * from dbm_user where user_id=" & user_id
set rs=server.createobject("ADODB.Recordset")
rs.open sql,conn,1,1
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
<FORM action="dbm_user_edit.asp" method="post" name="form_administrator" id="form_administrator" onSubmit="return checkform();">
  <TABLE width=100% 
                  border=1 cellPadding=3 cellSpacing=0 bordercolor="#C0C0C0">
    <TBODY>
      <TR align="center"> 
        <TD colspan="2"><strong><font color="#FF0000" size="+1">管理员编辑</font></strong> </TD>
      </TR>
      <TR> 
        <TD width=46% align="right"> 姓名</TD>
        <TD width=396><INPUT maxLength=100 size=30 id="allname" name="allname" value="<%=rs("allname")%>">
          *</TD>
      </TR>
      <TR> 
        <TD width=46% align="right">地址</TD>
        <TD width=396><INPUT maxLength=100 size=30 id="addr" name="addr" value="<%=rs("addr")%>">
        </TD>
      </TR>
      <TR> 
        <TD width=46% align="right">电话1</TD>
        <TD width=396><INPUT maxLength=100 size=30 id="tel2" name="tel2" value="<%=rs("tel2")%>">
        </TD>
      </TR>
      <TR> 
        <TD width=46% align="right"> 电话2</TD>
        <TD width=396><INPUT maxLength=100 size=30 id="tel3" name="tel3" value="<%=rs("tel3")%>">
        </TD>
      </TR>
        <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							              email
<!------------------------------------------------------------------------------------------------------------------------------------ -->
      <TR> 
        <TD width=46% align="right"> E-mail</TD>
        <TD width=396><INPUT maxLength=40 size=30 id="email" name="email" value="<%=rs("email")%>">
        </TD>
      </TR>
        <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							              email
<!------------------------------------------------------------------------------------------------------------------------------------ -->
      <TR> 
        <TD width=46% align="right"> OICQ / ICQ</TD>
        <TD width=396><INPUT maxLength=40 size=30 id="qq" name="qq" value="<%=rs("qq")%>">
        </TD>
      </TR>
        <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							              email
<!------------------------------------------------------------------------------------------------------------------------------------ -->
      <TR> 
        <TD width=46% align="right"> MSN</TD>
        <TD width=396><INPUT maxLength=40 size=30 id="msn" name="msn" value="<%=rs("msn")%>">
        </TD>
      </TR>
      <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							              USERID
<!------------------------------------------------------------------------------------------------------------------------------------ -->
      <TR> 
        <TD width="46%" align="right"> 用户账号</TD>
        <TD><INPUT maxLength=20 size=30 id="uid" name="uid"  value="<%=rs("uid")%>">
          *</TD>
      </TR>
      <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							         PASSWORD
<!------------------------------------------------------------------------------------------------------------------------------------ -->
      <TR> 
        <TD width="46%" align="right"> 密码</TD>
        <TD><INPUT maxLength=20 size=30 id="pwd" name="pwd" value="<%=rs("pwd")%>">
          *</TD>
      </TR>
      <TR> 
        <TD width="46%" align="right"> 重复密码</TD>
        <TD><INPUT maxLength=20 size=30 id="cpwd" name="cpwd" value="<%=rs("pwd")%>">
          *</TD>
      </TR>
      <TR> 
        <TD height=45 align="right"> 备注</TD>
        <TD><textarea name="memo" id="memo" cols="60" rows="10"><%=rs("memo")%>
</textarea></TD>
      </TR>
      <TR> 
        <TD height=32 colspan="2">&nbsp;</TD>
      </TR>
      <TR> 
        <TD height="50" colSpan=2> <P><INPUT type=submit id="option" name="option" value="edit">
            &nbsp;&nbsp; 
            <INPUT type="button" id="goback" name="goback" value="go back" onClick="history.go( -1 );"></P>
          </TD>
      </TR>
    </TBODY>
  </TABLE>
  <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							HIDDEN FIELD
<!------------------------------------------------------------------------------------------------------------------------------------ -->
<INPUT type="hidden" id="user_id" name="user_id" value=<%=user_id%>>
</FORM>
<%
rs.close
set rs=nothing
%>
<P></P>
<P></P>
<br>
<br>
</body>
</html>
