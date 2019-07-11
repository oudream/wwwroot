<script language="JavaScript">

function checkform()
{
	if (form_museum.mname.value.length == 0) {
		alert("博物馆名称不能为空!请重新输入!");
		form_museum.mname.focus();
		return false;
		}
	
	if(form_museum.email.value.length > 1)
	{
		if(! isMail(form_museum.email.value)) {
		alert('请输入有效的E-mail地址!');
		form_museum.email.focus();
		return false;
		}
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
	if(!isEnglish(name))
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
if request("option")="添加" then

mname=trim(request("mname"))
mname=replace(mname,"'","''")
addr=trim(request("addr"))
addr=replace(addr,"'","''")
zip=trim(request("zip"))
zip=replace(zip,"'","''")
tel=trim(request("tel"))
tel=replace(tel,"'","''")
fax=trim(request("fax"))
fax=replace(fax,"'","''")
email=trim(request("email"))
email=replace(email,"'","''")
webaddr=trim(request("webaddr"))
webaddr=replace(webaddr,"'","''")
ic_name=trim(request("ic_name"))
ic_name=replace(ic_name,"'","''")
ic_mp=trim(request("ic_mp"))
ic_mp=replace(ic_mp,"'","''")
ic_wp=trim(request("ic_wp"))
ic_wp=replace(ic_wp,"'","''")
intro=trim(request("intro"))
intro=replace(intro,"'","''")

if addr="" then addr=" " 
if zip="" then zip=" "
if tel="" then tel=" "
if fax="" then fax=" "
if email="" then email=" "
if webaddr="" then webaddr=" "
if ic_name="" then ic_name=" "
if ic_mp="" then ic_mp=" "
if ic_wp="" then ec_wp=" "
if intro="" then intro=" "

sql="insert into art_museum (artname,addr,zip,tel,fax,email,website_addr,principal,mp_principal,wp_principal,briefintro) values ('"&mname&"','"&addr&"','"&zip&"','"&tel&"','"&fax&"','"&email&"','"&webaddr&"','"&ic_name&"','"&ic_mp&"','"&ic_wp&"','"&intro&"')"
conn.Execute(sql)
response.Write("<SCRIPT LANGUAGE=JavaScript> alert('添加成功');</SCRIPT>")
End if
%>
<H2><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000066>博物馆资料编缉</FONT></H2>
<!----------------------------------------------------------------------------------------------------------------------
<!--                                                   ERROR MSGS
<!-------------------------------------------------------------------------------------------------------------------- -->
<FORM action="art_museum_add.asp" method="post" name="form_museum" id="form_museum" onSubmit="return checkform();">
  <TABLE height=307 cellSpacing=0 cellPadding=0 width=550 
                  border=0>
    <TBODY>
      <TR> 
        <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							              NAME
<!------------------------------------------------------------------------------------------------------------------------------------ -->
        <TD width=120 height=45> 名称</TD>
        <TD width=396><INPUT maxLength=100 size=30 id="mname" name="mname">
          *</TD>
      </TR>
      <TR> 
        <TD width=120 height=45>地址</TD>
        <TD width=396><INPUT maxLength=100 size=30 id="addr" name="addr">
          *</TD>
      </TR>
      <TR> 
        <TD width=120 height=45>邮编</TD>
        <TD width=396><INPUT maxLength=100 size=30 id="zip" name="zip">
          *</TD>
      </TR>
      <TR> 
        <TD width=120 height=45> 电话</TD>
        <TD width=396><INPUT maxLength=100 size=30 id="tel" name="tel">
          *</TD>
      </TR>
      <TR> 
        <TD width=120 height=45>传真</TD>
        <TD width=396><INPUT maxLength=100 size=30 id="fax" name="fax">
          *</TD>
      </TR>
      <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							              email
<!------------------------------------------------------------------------------------------------------------------------------------ -->
      <TR> 
        <TD width=120 height=45> 电子邮箱</TD>
        <TD width=396><INPUT maxLength=40 size=30 id="email" name="email">
          *</TD>
      </TR>
      <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							              email
<!------------------------------------------------------------------------------------------------------------------------------------ -->
      <TR> 
        <TD width=120 height=45> 网址</TD>
        <TD width=396><INPUT maxLength=40 size=30 id="webaddr" name="webaddr">
          *</TD>
      </TR>
      <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							              email
<!------------------------------------------------------------------------------------------------------------------------------------ -->
      <TR> 
        <TD width=120 rowspan="3"> 负责人</TD>
        <TD width=396 height="45"> 姓名： 
          <INPUT maxLength=40 size=30 id="ic_name" name="ic_name">
          *</TD>
      </TR>
      <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							              USERID
<!------------------------------------------------------------------------------------------------------------------------------------ -->
      <TR> 
        <TD height="45">手机号码：
          <INPUT maxLength=20 size=30 id="ic_mp" name="ic_mp">
          *</TD>
      </TR>
      <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							         PASSWORD
<!------------------------------------------------------------------------------------------------------------------------------------ -->
      <TR> 
        <TD height="45">办公电话：
          <INPUT maxLength=20 size=30 id="ic_wp" name="ic_wp">
          *</TD>
      </TR>
      <TR> 
        <TD height=45> 简介</TD>
        <TD><textarea name="intro" id="intro" cols="60" rows="10"></textarea></TD>
      </TR>
      <TR> 
        <TD height=32 colspan="2">&nbsp;</TD>
      </TR>
      <TR> 
        <TD height="50" colSpan=2> <P> 
            <INPUT type=submit id="option" name="option" value="添加">
            &nbsp;&nbsp; 
            <INPUT type="button" id="goback" name="goback" value="返回" onClick="history.go( -1 );">
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
