<script language="JavaScript">

function checkform()
{
	if (form_administrator.allname.value.length == 0) {
		alert("������ȫ��");
		form_administrator.allname.focus();
		return false;
		}
	if(form_administrator.email.value.length > 1){
	if(! isMail(form_administrator.email.value)) {
		alert('��������ȷ���ʼ���ַ');
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
brief=trim(request("brief"))
brief=replace(brief,"'","''")
memo=trim(request("memo"))
memo=replace(memo,"'","''")

usql="select * from art_shopper_govermente where allname='"&allname&"' and art_shopper_govermente_id<>" & request("art_shopper_govermente_id")
set urs=server.createobject("ADODB.Recordset")
urs.open usql,conn,1,1
if not ( urs.eof or urs.bof ) then
response.Write("<SCRIPT LANGUAGE=JavaScript> alert('�����˵ط����������ͬ�����ڣ�����������������');</SCRIPT>")
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
if brief="" then brief=" "
if memo="" then memo=" "
conn.Execute("update art_shopper_govermente set allname='"&allname&"',addr='"&addr&"',tel='"&tel&"',fax='"&fax&"',zip='"&zip&"',email='"&email&"',web_addr='"&web_addr&"',contacter='"&contacter&"',brief='"&brief&"',memo='"&memo&"' where art_shopper_govermente_id=" & request("art_shopper_govermente_id"))
response.Write("<SCRIPT LANGUAGE=JavaScript> alert('�༭�ɹ�');</SCRIPT>")
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
<table width="75%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
<TABLE width=100% 
                  border=1 cellPadding=3 cellSpacing=0 bordercolor="#CCCCCC">
<FORM action="art_shopper_govermente_edit.asp" method="post" name="form_administrator" id="form_administrator" onSubmit="return checkform();">
    <TBODY>
      <TR align="center"> 
        <TD colspan="2"><font color="#FF0000" size="+1"><strong>�ط�������ұ༩</strong></font></TD>
      </TR>
      <TR> 
        <TD width="46%" align="right">ȫ�ƣ�</TD>
        <TD> 
          <input maxlength=100 size=30 id="allname" name="allname" value="<%=rs("allname")%>">
          * </TD>
      </TR>
      <TR> 
        <TD align="right"> ��ַ��</TD>
        <TD> 
          <INPUT maxLength=100 size=30 id="addr" name="addr" value="<%=rs("addr")%>">
          </TD>
      </TR>
      <TR> 
        <TD align="right"> �绰��</TD>
        <TD> 
          <INPUT maxLength=100 size=30 id="tel" name="tel" value="<%=rs("tel")%>">
          </TD>
      </TR>
      <TR> 
        <TD align="right"> ���棺</TD>
        <TD> 
          <INPUT maxLength=100 size=30 id="fax" name="fax" value="<%=rs("fax")%>">
          </TD>
      </TR>
      <TR> 
        <TD align="right"> �ʱࣺ</TD>
        <TD> 
          <INPUT maxLength=100 size=30 id="zip" name="zip" value="<%=rs("zip")%>">
          </TD>
      </TR>
      <TR> 
        <TD align="right"> ���䣺</TD>
        <TD> 
          <INPUT maxLength=40 size=30 id="email" name="email" value="<%=rs("email")%>">
          </TD>
      </TR>
      <TR> 
        <TD align="right"> ��ַ��</TD>
        <TD> 
          <INPUT maxLength=40 size=30 id="web_addr" name="web_addr" value="<%=rs("web_addr")%>">
          </TD>
      </TR>
      <TR> 
        <TD align="right"> ��ϵ�ˣ�</TD>
        <TD> 
          <INPUT maxLength=40 size=30 id="contacter" name="contacter" value="<%=rs("contacter")%>">
          </TD>
      </TR>
      <TR> 
        <TD align="right"> �ط���ʷ���Ļ���</TD>
        <TD> 
          <textarea name="brief" id="brief" cols="60" rows="10"><%=rs("brief")%></textarea>
          </TD>
      </TR>
      <TR> 
        <TD align="right"> ��ע��</TD>
        <TD> 
          <textarea name="memo" id="memo" cols="60" rows="10"><%=rs("memo")%></textarea>
          </TD>
      </TR>
      <TR align="center"> 
        <TD height="50" colSpan=2> 
            <INPUT type=submit id="option" name="option" value="edit">
            &nbsp;&nbsp;&nbsp; 
            <INPUT type="button" id="goback" name="goback" value="go back" onClick="history.go( -1 );">
            <input type="hidden" id="art_shopper_govermente_id" name="art_shopper_govermente_id" value="<%=request("art_shopper_govermente_id")%>">
            </TD>
      </TR>
    </TBODY>
</FORM>
  </TABLE>
</body>
</html>
