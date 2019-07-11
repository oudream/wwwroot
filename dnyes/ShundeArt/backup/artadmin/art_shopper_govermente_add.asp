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
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="css.css" type=text/css rel=stylesheet>
</head>
<body text="#000000" leftmargin="8" topmargin="8" marginwidth="0" marginheight="0">
<!--#include file="adminconn.asp" -->
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
contacter=trim(request("contacter"))
contacter=replace(contacter,"'","''")
memo=trim(request("memo"))
memo=replace(memo,"'","''")
usql="select * from art_shopper_govermente where allname='"&allname&"'"
set urs=server.createobject("ADODB.Recordset")
urs.open usql,conn,1,1
if not ( urs.eof or urs.bof ) then
response.Write("<SCRIPT LANGUAGE=JavaScript> alert('出错！此美术家有同名存在，请用其它名字输入');</SCRIPT>")
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
if contacter="" then contacter=" "
if memo="" then memo=" "
sql="insert into art_shopper_govermente (allname,addr,tel,fax,zip,email,web_addr,contacter,memo,inserttime) values ('"&allname&"','"&addr&"','"&tel&"','"&fax&"','"&zip&"','"&email&"','"&web_addr&"','"&contacter&"','"&memo&"','"&now()&"')"
conn.Execute(sql)
response.Write("<SCRIPT LANGUAGE=JavaScript> alert('添加成功');</SCRIPT>")
end if
end if
%>
<H2><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000066>美术家添加</FONT></H2>
<!----------------------------------------------------------------------------------------------------------------------
<!--                                                   ERROR MSGS
<!-------------------------------------------------------------------------------------------------------------------- -->
<FORM action="art_shopper_govermente_add.asp" method="post" name="form_administrator" id="form_administrator" onSubmit="return checkform();">
  <TABLE height=307 cellSpacing=0 cellPadding=0 width=550 
                  border=0>
    <TBODY>
      <TR> 
        <TD width=120 height=45> <DIV align=center><font color="#000000" face="Verdana, Arial, Helvetica, sans-serif">姓名</font></DIV></TD>
        <TD width=396><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 > 
          <INPUT maxLength=100 size=30 id="allname" name="allname">
          * </FONT></TD>
      </TR>
      <TR> 
        <TD width=120 height=45> <DIV align=center><font color="#000000" face="Verdana, Arial, Helvetica, sans-serif">地址</font></DIV></TD>
        <TD width=396><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 > 
          <INPUT maxLength=100 size=30 id="addr" name="addr">
          </FONT></TD>
      </TR>
      <TR> 
        <TD width=120 height=45> <DIV align=center><font color="#000000"  face="Verdana, Arial, Helvetica, sans-serif">电话</font></DIV></TD>
        <TD width=396><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 > 
          <INPUT maxLength=100 size=30 id="tel" name="tel">
          </FONT></TD>
      </TR>
      <TR> 
        <TD width=120 height=45> <DIV align=center><font color="#000000"  face="Verdana, Arial, Helvetica, sans-serif">传真</font></DIV></TD>
        <TD width=396><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 > 
          <INPUT maxLength=100 size=30 id="fax" name="fax">
          </FONT></TD>
      </TR>
      <TR> 
        <TD width=120 height=45> <DIV align=center><font color="#000000"  face="Verdana, Arial, Helvetica, sans-serif">邮编</font></DIV></TD>
        <TD width=396><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 > 
          <INPUT maxLength=100 size=30 id="zip" name="zip">
          </FONT></TD>
      </TR>
      <TR> 
        <TD width=120 height=45> <DIV align=center><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 >E-mail</FONT></DIV></TD>
        <TD width=396><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 > 
          <INPUT maxLength=40 size=30 id="email" name="email">
          </FONT></TD>
      </TR>
      <TR> 
        <TD width=120 height=45> <DIV align=center><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 >网址</FONT></DIV></TD>
        <TD width=396><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 > 
          <INPUT maxLength=40 size=30 id="web_addr" name="web_addr">
          </FONT></TD>
      </TR>
      <TR> 
        <TD width=120 height=45> <DIV align=center><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 >联系人</FONT></DIV></TD>
        <TD width=396><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 > 
          <INPUT maxLength=40 size=30 id="concacter" name="contacter">
          </FONT></TD>
      </TR>
      <TR> 
        <TD height=32> <DIV align=center><font color="#000000" face="Verdana, Arial, Helvetica, sans-serif">地方历史及文化</font></DIV></TD>
        <TD><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 > 
          <textarea name="memo" id="memo" cols="60" rows="10"></textarea>
          </FONT></TD>
      </TR>
      <TR> 
        <TD height="50" colSpan=2> <P><STRONG><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000066 size=2> 
            <INPUT type=submit id="option" name="option" value="add">
            </FONT></STRONG></P></TD>
      </TR>
    </TBODY>
  </TABLE>
<INPUT maxLength=100 size=30 id="contacter" name="contacter" type="hidden">
</FORM>
<br>
<br>
</body>
</html>
