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
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=site.css rel=stylesheet>
<title></title>
</head>
<body topmargin="0">
<!--#include file="art_adminconn.asp" -->






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
memo=trim(request("memo"))
memo=replace(memo,"'","''")

usql="select * from art_museum where artname='" & mname & "'"
set urs=server.createobject("ADODB.Recordset")
urs.open usql,conn,1,1
if not ( urs.eof or urs.bof ) then
response.Write("<SCRIPT LANGUAGE=JavaScript> alert('出错！此博物馆有同名存在，请用其它名字输入');</SCRIPT>")
urs.close
set urs=nothing
else
urs.close
set urs=nothing

if addr="" then addr=" " 
if zip="" then zip=" "
if tel="" then tel=" "
if fax="" then fax=" "
if email="" then email=" "
if webaddr="" then webaddr=" "
if ic_name="" then ic_name=" "
if ic_mp="" then ic_mp=" "
if ic_wp="" then ic_wp=" "
if intro="" then intro=" "
if memo="" then memo=" "

sql="insert into art_museum (artname,addr,zip,tel,fax,email,website_addr,principal,mp_principal,wp_principal,briefintro,memo) values ('"&mname&"','"&addr&"','"&zip&"','"&tel&"','"&fax&"','"&email&"','"&webaddr&"','"&ic_name&"','"&ic_mp&"','"&ic_wp&"','"&intro&"','"&memo&"')"
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
<FORM action="art_museum_add.asp" method="post" name="form_museum" id="form_museum" onSubmit="return checkform();">
    <TBODY>
      <TR align="center"> 
        <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							              NAME
<!------------------------------------------------------------------------------------------------------------------------------------ -->
        <TD colspan="4"><font color="#FF0000" size="+1"><strong>博物馆添加</strong></font></TD>
      </TR>
      <TR> 
        <TD colspan="2" align="right" width="46%"> 名称：</TD>
        <TD colspan="2"><input maxlength=100 size=30 id="mname" name="mname">
          *</TD>
      </TR>
      <TR> 
        <TD colspan="2" align="right">地址：</TD>
        <TD colspan="2"><INPUT maxLength=100 size=30 id="addr" name="addr">
          *</TD>
      </TR>
      <TR> 
        <TD colspan="2" align="right">邮编：</TD>
        <TD colspan="2"><INPUT maxLength=100 size=30 id="zip" name="zip">
          *</TD>
      </TR>
      <TR> 
        <TD colspan="2" align="right"> 电话：</TD>
        <TD colspan="2"><INPUT maxLength=100 size=30 id="tel" name="tel">
          *</TD>
      </TR>
      <TR> 
        <TD colspan="2" align="right">传真：</TD>
        <TD colspan="2"><INPUT maxLength=100 size=30 id="fax" name="fax">
          *</TD>
      </TR>
      <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							              email
<!------------------------------------------------------------------------------------------------------------------------------------ -->
      <TR> 
        <TD colspan="2" align="right"> 电子邮箱：</TD>
        <TD colspan="2"><INPUT maxLength=40 size=30 id="email" name="email">
          *</TD>
      </TR>
      <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							              email
<!------------------------------------------------------------------------------------------------------------------------------------ -->
      <TR> 
        <TD colspan="2" align="right"> 网址：</TD>
        <TD colspan="2"><INPUT maxLength=40 size=30 id="webaddr" name="webaddr">
          *</TD>
      </TR>
      <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							              email
<!------------------------------------------------------------------------------------------------------------------------------------ -->
      <TR> 
        <TD rowspan="3" align="right"> 负责人：</TD>
        <TD align="right">姓名：</TD>
        <TD colspan="2"> <input maxlength=40 size=30 id="ic_name" name="ic_name">
          *</TD>
      </TR>
      <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							              USERID
<!------------------------------------------------------------------------------------------------------------------------------------ -->
      <TR> 
        <TD align="right">手机号码： </TD>
        <TD colspan="2"><input maxlength=20 size=30 id="ic_mp2" name="ic_mp">
          *</TD>
      </TR>
      <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							         PASSWORD
<!------------------------------------------------------------------------------------------------------------------------------------ -->
      <TR> 
        <TD align="right">办公电话： </TD>
        <TD colspan="2"> <input maxlength=20 size=30 id="ic_wp2" name="ic_wp">
          *</TD>
      </TR>
      <TR> 
        <TD colspan="2" align="right"> 简介：</TD>
        <TD colspan="2"><textarea name="intro" id="intro" cols="60" rows="10"></textarea></TD>
      </TR>
      <TR> 
        <TD colspan="2" align="right"> 备注：</TD>
        <TD> <textarea name="memo" id="memo" cols="60" rows="10"></textarea> 
        </TD>
      </TR>
      <TR> 
        <TD height=32 colspan="4">&nbsp;</TD>
      </TR>
      <TR> 
        <TD height="50" colSpan=4 align="center"> <INPUT type=submit id="option" name="option" value="添加"> 
          &nbsp;&nbsp; <INPUT type="button" id="goback" name="goback" value="返回" onClick="history.go( -1 );"> 
        </TD>
      </TR>
    </TBODY>
</FORM>
  </TABLE>
</body>
</html>
