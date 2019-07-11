<script language="JavaScript">
function checkform()
{
	if (form_administrator.uname.value.length == 0) {
		alert("Please enter a valid Name. ");
		form_administrator.uname.focus();
		return false;
		}
	if(form_administrator.email.value.length > 1) {
	if(! isMail(form_administrator.email.value)) {
		alert('Please enter a valid EMAIL address');
		form_administrator.email.focus();
		return false;
		}}
	if (form_administrator.uid.value.length == 0) {
		alert("Please enter a valid Userid. ");
		form_administrator.uid.focus();
		return false;
		}
	if(! isLetter(form_administrator.uid.value)) {
		alert('The "Userid" field makes up by "_" and "0~9" adn "a~z" and "A~Z" and blank space');
		form_administrator.uid.focus();
		return false;
		}
	if (form_administrator.pwd.value.length == 0) {
		alert(" Enter the name and type of competition. ");
		form_administrator.pwd.focus();
		return false;
		}
	if(! isLetter(form_administrator.pwd.value)) {
		alert('The "Password" field makes up by "_" and "0~9" adn "a~z" and "A~Z" and blank space');
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
<LINK href="css.css" type=text/css rel=stylesheet>
</head>
<body text="#000000" leftmargin="8" topmargin="8" marginwidth="0" marginheight="0">

<!--#include file="adminconn.asp" -->
<%
if request("option")="edit" then
uname=trim(request("uname"))
uname=replace(uname,"'","''")
addr=trim(request("addr"))
addr=replace(addr,"'","''")
tel2=trim(request("tel2"))
tel2=replace(tel2,"'","''")
tel3=trim(request("tel3"))
tel3=replace(tel3,"'","''")
fax=trim(request("fax"))
fax=replace(fax,"'","''")
email=trim(request("email"))
email=replace(email,"'","''")
qq=trim(request("qq"))
qq=replace(qq,"'","''")
msn=trim(request("msn"))
msn=replace(msn,"'","''")
uid=trim(request("uid"))
pwd=trim(request("pwd"))
memo=trim(request("memo"))
memo=replace(memo,"'","''")



usql="select * from art_user where id<>"&session("user_id")&" and ( uid='"&uid&"' or uname='"&uname&"' )"
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
if fax="" then fax=" "
if email="" then email=" "
if memo="" then memo=" "
if uname<>"" then
conn.Execute("update art_user set uname='"&uname&"',addr='"&addr&"',tel2='"&tel2&"',tel3='"&tel3&"',fax='"&fax&"',email='"&email&"',qq='"&qq&"',msn='"&msn&"',uid='"&uid&"',pwd='"&pwd&"',memo='"&memo&"' where id=" & session("user_id"))
response.Write("<SCRIPT LANGUAGE=JavaScript> alert('编辑成功');</SCRIPT>")
end if
end if
end if
%>
<%
sql="select * from art_user where id=" & session("user_id")
set rs=server.createobject("ADODB.Recordset")
rs.open sql,conn,1,1
%>
<H2><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000066>个人资料编缉</FONT></H2>
<!----------------------------------------------------------------------------------------------------------------------
<!--                                                   ERROR MSGS
<!-------------------------------------------------------------------------------------------------------------------- -->
<FORM action="user_edit.asp" method="post" name="form_administrator" id="form_administrator" onSubmit="return checkform();">
  <TABLE height=307 cellSpacing=0 cellPadding=0 width=550 
                  border=0>
    <TBODY>
      <TR> 
        <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							              NAME
<!------------------------------------------------------------------------------------------------------------------------------------ -->
        <TD width=120 height=45> 姓名</TD>
        <TD width=396><INPUT maxLength=100 size=30 id="uname" name="uname" value="<%=rs("uname")%>">
          *</TD>
      </TR>
      <TR> 
        <TD width=120 height=45>地址</TD>
        <TD width=396><INPUT maxLength=100 size=30 id="addr" name="addr" value="<%=rs("addr")%>">
          *</TD>
      </TR>
      <TR> 
        <TD width=120 height=45>电话1</TD>
        <TD width=396><INPUT maxLength=100 size=30 id="tel2" name="tel2" value="<%=rs("tel2")%>">
          *</TD>
      </TR>
      <TR> 
        <TD width=120 height=45> 电话2</TD>
        <TD width=396><INPUT maxLength=100 size=30 id="tel3" name="tel3" value="<%=rs("tel3")%>">
          *</TD>
      </TR>
      <TR> 
        <TD width=120 height=45>传真</TD>
        <TD width=396><INPUT maxLength=100 size=30 id="fax" name="fax" value="<%=rs("fax")%>">
          *</TD>
      </TR>
        <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							              email
<!------------------------------------------------------------------------------------------------------------------------------------ -->
      <TR> 
        <TD width=120 height=45> 邮箱</TD>
        <TD width=396><INPUT maxLength=40 size=30 id="email" name="email" value="<%=rs("email")%>">
          *</TD>
      </TR>
        <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							              email
<!------------------------------------------------------------------------------------------------------------------------------------ -->
      <TR> 
        <TD width=120 height=45> QQ</TD>
        <TD width=396><INPUT maxLength=40 size=30 id="qq" name="qq" value="<%=rs("qq")%>">
          *</TD>
      </TR>
        <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							              email
<!------------------------------------------------------------------------------------------------------------------------------------ -->
      <TR> 
        <TD width=120 height=45> MSN</TD>
        <TD width=396><INPUT maxLength=40 size=30 id="msn" name="msn" value="<%=rs("msn")%>">
          *</TD>
      </TR>
      <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							              USERID
<!------------------------------------------------------------------------------------------------------------------------------------ -->
      <TR> 
        <TD height=45> Userid</TD>
        <TD><INPUT maxLength=20 size=30 id="uid" name="uid"  value="<%=rs("uid")%>">
          *</TD>
      </TR>
      <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							         PASSWORD
<!------------------------------------------------------------------------------------------------------------------------------------ -->
      <TR> 
        <TD height=45> Password</TD>
        <TD><INPUT maxLength=20 size=30 id="pwd" name="pwd" value="<%=rs("pwd")%>">
          *</TD>
      </TR>
      <TR> 
        <TD height=45> 备注</TD>
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
