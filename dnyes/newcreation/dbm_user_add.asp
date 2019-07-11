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
if request("option")="add" then
allname=trim(request("allname"))
allname=replace(allname,"'","''")
adminlevel=request("adminlevel")
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



usql="select * from dbm_user where uid='"&uid&"' or allname='"&allname&"'"
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
if email="" then email=" "
if qq="" then qq=" "
if msn="" then msn=" "
if memo="" then memo=" "
if allname<>"" then
sql="insert into dbm_user (allname,adminlevel,addr,tel2,tel3,email,qq,msn,uid,pwd,memo,inserttime) values ('"&allname&"',"&adminlevel&",'"&addr&"','"&tel2&"','"&tel3&"','"&email&"','"&qq&"','"&msn&"','"&uid&"','"&pwd&"','"&memo&"','"&now()&"')"
conn.Execute(sql)
response.Write("<SCRIPT LANGUAGE=JavaScript> alert('添加成功');</SCRIPT>")
end if
end if
end if
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
<FORM action="dbm_user_add.asp" method="post" name="form_administrator" id="form_administrator" onSubmit="return checkform();">
  <TABLE width=100% 
                  border=1 cellPadding=3 cellSpacing=0 bordercolor="#C0C0C0">
    <TBODY>
      <TR align="center"> 
        <TD colspan="2"><strong><font color="#FF0000" size="+1">管理员添加</font></strong> </TD>
      </TR>
      <TR> 
        <TD width=46% align="right"> 姓名</TD>
        <TD width=396><INPUT maxLength=100 size=30 id="allname" name="allname">
          *</TD>
      </TR>
      <TR> 
        <TD width=46% align="right">管理员级别</TD>
        <TD width=396><select name="adminlevel" id="adminlevel">
            <option value="1">高级管理员</option>
            <option value="2" selected>一般管理员</option>
          </select>
          *</TD>
      </TR>
      <TR> 
        <TD width=46% align="right">地址</TD>
        <TD width=396><INPUT maxLength=100 size=30 id="addr" name="addr"> </TD>
      </TR>
      <TR> 
        <TD width=46% align="right">电话1</TD>
        <TD width=396><INPUT maxLength=100 size=30 id="tel2" name="tel2"> </TD>
      </TR>
      <TR> 
        <TD width=46% align="right"> 电话2</TD>
        <TD width=396><INPUT maxLength=100 size=30 id="tel3" name="tel3"> </TD>
      </TR>
      <TR> 
        <TD width=46% align="right"> E-mail</TD>
        <TD width=396><INPUT maxLength=40 size=30 id="email" name="email"> </TD>
      </TR>
      <TR> 
        <TD width=46% align="right"> OICQ / ICQ</TD>
        <TD width=396><INPUT maxLength=40 size=30 id="qq" name="qq"> </TD>
      </TR>
      <TR> 
        <TD width=46% align="right"> MSN</TD>
        <TD width=396><INPUT maxLength=40 size=30 id="msn" name="msn"> </TD>
      </TR>
      <TR> 
        <TD width="46%" align="right"> 用户账号</TD>
        <TD><INPUT maxLength=20 size=30 id="uid" name="uid">
          *</TD>
      </TR>
      <TR> 
        <TD width="46%" align="right"> 密码</TD>
        <TD><INPUT maxLength=20 size=30 id="pwd" name="pwd">
          *</TD>
      </TR>
      <TR> 
        <TD width="46%" align="right"> 重复密码</TD>
        <TD><INPUT maxLength=20 size=30 id="cpwd" name="cpwd">
          *</TD>
      </TR>
      <TR> 
        <TD width="46%" align="right"> 备注</TD>
        <TD><textarea name="memo" id="memo" cols="60" rows="10"></textarea></TD>
      </TR>
      <TR> 
        <TD height=32 colspan="2">&nbsp;</TD>
      </TR>
      <TR> 
        <TD height="50" colSpan=2> <P> 
            <INPUT type=submit id="option" name="option" value="add">
            &nbsp;&nbsp; 
            <INPUT type="button" id="goback" name="goback" value="go back" onClick="history.go( -1 );">
          </P></TD>
      </TR>
    </TBODY>
  </TABLE>
  <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							HIDDEN FIELD
<!------------------------------------------------------------------------------------------------------------------------------------ -->
</FORM>
<P></P>
<P></P>
<br>
<br>
</body>
</html>
