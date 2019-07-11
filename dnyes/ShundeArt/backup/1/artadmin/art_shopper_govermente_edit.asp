<script language="JavaScript">

function checkform()
{
	if (form_administrator.allname.value.length == 0) {
		alert("请输入全名");
		form_administrator.allname.focus();
		return false;
		}
	if(form_administrator.email.value.length > 1){
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
if request("option")="edit" then

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

usql="select * from art_shopper_govermente where allname='"&allname&"' and art_shopper_govermente_id<>" & request("art_shopper_govermente_id")
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
conn.Execute("update art_shopper_govermente set allname='"&allname&"',addr='"&addr&"',tel='"&tel&"',fax='"&fax&"',zip='"&zip&"',email='"&email&"',web_addr='"&web_addr&"',contacter='"&contacter&"',memo='"&memo&"' where art_shopper_govermente_id=" & request("art_shopper_govermente_id"))
response.Write("<SCRIPT LANGUAGE=JavaScript> alert('编辑成功');</SCRIPT>")
end if
end if
%>
<%
sql="select * from art_shopper_govermente where art_shopper_govermente_id="&request("art_shopper_govermente_id") 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1 

if rs.eof or rs.bof then 
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' None data ! ');</SCRIPT>")
response.End()
end if
%>
<H2><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000066>美术家编缉</FONT></H2>
<!----------------------------------------------------------------------------------------------------------------------
<!--                                                   ERROR MSGS
<!-------------------------------------------------------------------------------------------------------------------- -->
<FORM action="art_shopper_govermente_edit.asp" method="post" name="form_administrator" id="form_administrator" onSubmit="return checkform();">
  <TABLE height=307 cellSpacing=0 cellPadding=0 width=550 
                  border=0>
    <TBODY>
      <TR> 
        <TD width=120 height=45> <DIV align=center><font color="#000000" face="Verdana, Arial, Helvetica, sans-serif">全称</font></DIV></TD>
        <TD width=396><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 > 
          <INPUT maxLength=100 size=30 id="allname" name="allname" value="<%=rs("allname")%>">
          * </FONT></TD>
      </TR>
      <TR> 
        <TD width=120 height=45> <DIV align=center><font color="#000000" face="Verdana, Arial, Helvetica, sans-serif">地址</font></DIV></TD>
        <TD width=396><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 > 
          <INPUT maxLength=100 size=30 id="addr" name="addr" value="<%=rs("addr")%>">
          </FONT></TD>
      </TR>
      <TR> 
        <TD width=120 height=45> <DIV align=center><font color="#000000"  face="Verdana, Arial, Helvetica, sans-serif">电话</font></DIV></TD>
        <TD width=396><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 > 
          <INPUT maxLength=100 size=30 id="tel" name="tel" value="<%=rs("tel")%>">
          </FONT></TD>
      </TR>
      <TR> 
        <TD width=120 height=45> <DIV align=center><font color="#000000"  face="Verdana, Arial, Helvetica, sans-serif">传真</font></DIV></TD>
        <TD width=396><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 > 
          <INPUT maxLength=100 size=30 id="fax" name="fax" value="<%=rs("fax")%>">
          </FONT></TD>
      </TR>
      <TR> 
        <TD width=120 height=45> <DIV align=center><font color="#000000" face="Verdana, Arial, Helvetica, sans-serif">邮编</font></DIV></TD>
        <TD width=396><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 > 
          <INPUT maxLength=100 size=30 id="zip" name="zip" value="<%=rs("zip")%>">
          </FONT></TD>
      </TR>
      <TR> 
        <TD width=120 height=45> <DIV align=center><font color="#000000" face="Verdana, Arial, Helvetica, sans-serif">邮箱</font></DIV></TD>
        <TD width=396><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 > 
          <INPUT maxLength=40 size=30 id="email" name="email" value="<%=rs("email")%>">
          </FONT></TD>
      </TR>
      <TR> 
        <TD width=120 height=45> <DIV align=center><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 >网址</FONT></DIV></TD>
        <TD width=396><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 > 
          <INPUT maxLength=40 size=30 id="web_addr" name="web_addr" value="<%=rs("web_addr")%>">
          </FONT></TD>
      </TR>
      <TR> 
        <TD width=120 height=45> <DIV align=center><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 >联系人</FONT></DIV></TD>
        <TD width=396><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 > 
          <INPUT maxLength=40 size=30 id="contacter" name="contacter" value="<%=rs("contacter")%>">
          </FONT></TD>
      </TR>
      <TR> 
        <TD width="120" height=32> <DIV align=center><font color="#000000" face="Verdana, Arial, Helvetica, sans-serif">备注</font></DIV></TD>
        <TD><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 > 
          <textarea name="memo" id="memo" cols="60" rows="10"><%=rs("memo")%></textarea>
          </FONT></TD>
      </TR>
      <TR> 
        <TD height="50" colSpan=2> <P><STRONG><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000066 size=2> 
            <INPUT type=submit id="option" name="option" value="edit">
            &nbsp;&nbsp;&nbsp; 
            <INPUT type="button" id="goback" name="goback" value="go back" onClick="history.go( -1 );">
            <input type="hidden" id="art_shopper_govermente_id" name="art_shopper_govermente_id" value="<%=request("art_shopper_govermente_id")%>">
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
