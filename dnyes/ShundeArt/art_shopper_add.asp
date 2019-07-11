<script language="JavaScript">

function checkform()
{
	if (form_shopper.allname.value.length == 0) {
		alert("姓名不能为空!请重新输入!");
		form_shopper.allname.focus();
		return false;
		}
	
	if(form_shopper.email.value.length > 1)
	{
		if(! isMail(form_shopper.email.value)) {
		alert('请输入有效的E-mail地址!');
		form_shopper.email.focus();
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
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=site.css rel=stylesheet>
<title></title>
</head>
<body topmargin="0">
<!--#include file="art_adminconn.asp" -->


<%
if request("option")="添加" then

allname=trim(request("allname"))
allname=replace(allname,"'","''")
addr=trim(request("addr"))
addr=replace(addr,"'","''")
zip=trim(request("zip"))
zip=replace(zip,"'","''")
tel=trim(request("tel"))
tel=replace(tel,"'","''")
sex=trim(request("sex"))
sex=replace(sex,"'","''")
email=trim(request("email"))
email=replace(email,"'","''")
nation=trim(request("nation"))
nation=replace(nation,"'","''")
original=trim(request("original"))
original=replace(original,"'","''")
birth=trim(request("birth"))
birth=replace(birth,"'","''")
workplace=trim(request("workplace"))
workplace=replace(workplace,"'","''")
duty=trim(request("duty"))
duty=replace(duty,"'","''")
hobby=trim(request("hobby"))
hobby=replace(hobby,"'","''")
memo=trim(request("memo"))
memo=replace(memo,"'","''")

usql="select * from art_shopper where allname='" & allname & "'"
set urs=server.createobject("ADODB.Recordset")
urs.open usql,conn,1,1
if not ( urs.eof or urs.bof ) then
response.Write("<SCRIPT LANGUAGE=JavaScript> alert('出错！此作品有同名存在，请用其它名字输入');</SCRIPT>")
urs.close
set urs=nothing
else
urs.close
set urs=nothing

if addr="" then addr=" " 
if zip="" then zip=" "
if tel="" then tel=" "
if sex="" then sex=" "
if email="" then email=" "
if nation="" then nation=" "
if original="" then original=" "
if birth="" then birth=" "
if workplace="" then workplace=" "
if duty="" then duty=" "
if hobby="" then hobby=" "
if memo="" then memo=" "

sql="insert into art_shopper (allname,addr,zip,tel,sex,email,nation,original,birth,workplace,duty,hobby,memo) values ('"&allname&"','"&addr&"','"&zip&"','"&tel&"','"&sex&"','"&email&"','"&nation&"','"&original&"','"&birth&"','"&workplace&"','"&duty&"','"&hobby&"','"&memo&"')"
conn.Execute(sql)
response.Write("<SCRIPT LANGUAGE=JavaScript> alert('添加成功');</SCRIPT>")
End if
End if
%>
<table width="75%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
<TABLE width=100% 
                  border=1 cellPadding=3 cellSpacing=0 bordercolor="#CCCCCC">
 <FORM action="art_shopper_add.asp" method="post" name="form_shopper" id="form_shopper" onSubmit="return checkform();">
   <TBODY>
      <TR align="center"> 
        <TD height="30" colspan="2"><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#FF0000><strong>个人客户买家资料编缉</strong></font></TD>
      </TR>
      <TR> 
        <TD width=46% align="right">姓名：</TD>
        <TD><input maxlength=100 size=30 id="allname" name="allname">
          *</TD>
      </TR>
      <TR> 
        <TD align="right">性别：</TD>
        <TD><INPUT maxLength=100 size=30 id="sex" name="sex">
          *</TD>
      </TR>
      <TR> 
        <TD align="right">民族：</TD>
        <TD><INPUT maxLength=100 size=30 id="nation" name="nation">
          *</TD>
      </TR>
      <TR> 
        <TD align="right">原藉：</TD>
        <TD><INPUT maxLength=100 size=30 id="original" name="original">
          *</TD>
      </TR>
      <TR> 
        <TD align="right">出生年月：</TD>
        <TD><INPUT maxLength=100 size=30 id="birth" name="birth">
          *</TD>
      </TR>
      <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							              email
<!------------------------------------------------------------------------------------------------------------------------------------ -->
      <TR> 
        <TD align="right">工作单位：</TD>
        <TD><INPUT maxLength=40 size=30 id="workplace" name="workplace">
          *</TD>
      </TR>
      <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							              email
<!------------------------------------------------------------------------------------------------------------------------------------ -->
      <TR> 
        <TD align="right">职务：</TD>
        <TD><INPUT maxLength=40 size=30 id="duty" name="duty">
          *</TD>
      </TR>
      <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							              email
<!------------------------------------------------------------------------------------------------------------------------------------ -->
      <TR> 
        <TD align="right">兴趣爱好：</TD>
        <TD> <INPUT maxLength=40 size=30 id="hobby" name="hobby">
          *</TD>
      </TR>
      <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							              USERID
<!----------------------------------------------------------------------------------------------------------------------------------- -->
      <TR> 
        <TD align="right">住址：</TD>
        <TD> <INPUT maxLength=20 size=30 id="addr" name="addr">
          *</TD>
      </TR>
      <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							         PASSWORD
<!------------------------------------------------------------------------------------------------------------------------------------ -->
      <TR> 
        <TD align="right">邮编：</TD>
        <TD> <INPUT maxLength=20 size=30 id="zip" name="zip">
          *</TD>
      </TR>
      <TR> 
        <TD align="right">电话：</TD>
        <TD> <INPUT maxLength=20 size=30 id="tel" name="tel">
          *</TD>
      </TR>
      <TR> 
        <TD align="right">E-mail：</TD>
        <TD> <INPUT maxLength=20 size=30 id="email" name="email">
          *</TD>
      </TR>
      <TR> 
        <TD height=32 align="right"> 备注：</TD>
        <TD> <textarea name="memo" id="memo" cols="60" rows="10"></textarea> 
        </TD>
      </TR>
    
    <TR> 
      <TD height="50" colSpan=2 align="center"> <INPUT type=submit id="option" name="option" value="添加"> 
        &nbsp;&nbsp; <INPUT type="button" id="goback" name="goback" value="返回" onClick="history.go( -1 );"> 
      </TD>
    </TR>
</FORM>
  </TABLE>
<br>
<br>
</body>
</html>
