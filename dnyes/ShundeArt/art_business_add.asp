<script language="JavaScript">

function checkform()
{
	if (form_administrator.allname.value.length == 0) {
		alert("请输入全称");
		form_administrator.allname.focus();
		return false;
		}
	if(form_administrator.email.value.length > 0){
	if(! isMail(form_administrator.email.value)) {
		alert('请输入正确的邮件地址');
		form_administrator.email.focus();
		return false;
		}}
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
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=site.css rel=stylesheet>
<title></title>
</head>
<body topmargin="0">
<!--#include file="art_adminconn.asp" -->






<%
if request("option")="add" then

allname=trim(request("allname"))
allname=replace(allname,"'","''")
addr=trim(request("addr"))
addr=replace(addr,"'","''")
tel=trim(request("tel"))
tel=replace(tel,"'","''")
fax=trim(request("fax"))
fax=replace(fax,"'","''")
zip=trim(request("zip"))
zip=replace(zip,"'","''")
email=trim(request("email"))
email=replace(email,"'","''")
web_addr=trim(request("web_addr"))
web_addr=replace(web_addr,"'","''")
legal_representative=trim(request("legal_representative"))
legal_representative=replace(legal_representative,"'","''")
brief=trim(request("brief"))
brief=replace(brief,"'","''")
memo=trim(request("memo"))
memo=replace(memo,"'","''")
usql="select * from art_business where allname='"&allname&"'"
set urs=server.createobject("ADODB.Recordset")
urs.open usql,conn,1,1
if not ( urs.eof or urs.bof ) then
response.Write("<SCRIPT LANGUAGE=JavaScript> alert('出错！此企业有同名存在，请用其它名字输入');</SCRIPT>")
urs.close
set urs=nothing
else
urs.close
set urs=nothing
if addr="" then addr=" "
if tel="" then tel=" "
if fax="" then fax=" "
if zip="" then zip=" "
if email="" then email=" "
if web_addr="" then web_addr=" "
if legal_representative="" then legal_representative=" "
if brief="" then brief=" "
if memo="" then memo=" "
sql="insert into art_business (allname,addr,tel,fax,zip,email,web_addr,legal_representative,brief,memo,inserttime) values ('"&allname&"','"&addr&"','"&tel&"','"&fax&"','"&zip&"','"&email&"','"&web_addr&"','"&legal_representative&"','"&brief&"','"&memo&"','"&now()&"')"
conn.Execute(sql)
response.Write("<SCRIPT LANGUAGE=JavaScript> alert('添加成功');</SCRIPT>")
end if
end if
%>
<table width="75%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
<TABLE width=100% 
                  border=1 cellPadding=3 cellSpacing=0 bordercolor="#CCCCCC">
<FORM action="art_business_add.asp" method="post" name="form_administrator" id="form_administrator" onSubmit="return checkform();">
    <TBODY>
      <TR align="center"> 
        <TD height=30 colspan="2"><font color="#FF0000" size="+1"><strong> 企业添加 
          </strong></font></TD>
      </TR>
      <TR> 
        <TD width="46%" align="right">姓名：</TD>
        <TD> 
          <input maxlength=100 size=30 id="allname" name="allname">
          * </TD>
      </TR>
      <TR> 
        <TD align="right"> 地址：</TD>
        <TD> 
          <INPUT maxLength=100 size=30 id="addr" name="addr">
          </TD>
      </TR>
      <TR> 
        <TD align="right"> 电话：</TD>
        <TD> 
          <INPUT maxLength=100 size=30 id="tel" name="tel">
          </TD>
      </TR>
      <TR> 
        <TD align="right"> 传真：</TD>
        <TD> 
          <INPUT maxLength=100 size=30 id="fax" name="fax">
          </TD>
      </TR>
      <TR> 
        <TD align="right"> 邮编：</TD>
        <TD> 
          <INPUT maxLength=100 size=30 id="zip" name="zip">
          </TD>
      </TR>
      <TR> 
        <TD align="right"> E-mail：</TD>
        <TD> 
          <INPUT maxLength=40 size=30 id="email" name="email">
          </TD>
      </TR>
      <TR> 
        <TD align="right"> 网址：</TD>
        <TD> 
          <INPUT maxLength=40 size=30 id="web_addr" name="web_addr">
          </TD>
      </TR>
      <TR> 
        <TD align="right"> 法定代理人：</TD>
        <TD> 
          <INPUT maxLength=40 size=30 id="legal_representative" name="legal_representative">
          </TD>
      </TR>
      <TR> 
        <TD align="right"> 简介：</TD>
        <TD> 
          <textarea name="brief" id="brief" cols="60" rows="10"></textarea>
          </TD>
      </TR>
      <TR> 
        <TD align="right"> 备注：</TD>
        <TD> 
          <textarea name="memo" id="memo" cols="60" rows="10"></textarea>
          </TD>
      </TR>
      <TR align="center"> 
        <TD height="50" colSpan=2>  
            <INPUT type=submit id="option" name="option" value="add">
            </TD>
      </TR>
    </TBODY>
</FORM>
  </TABLE>
</body>
</html>
