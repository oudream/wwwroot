<script language="JavaScript">

function checkform()
{
	if (form_shopper.allname.value.length == 0) {
		alert("��������Ϊ��!����������!");
		form_shopper.allname.focus();
		return false;
		}
	
	if(form_shopper.email.value.length > 1)
	{
		if(! isMail(form_shopper.email.value)) {
		alert('��������Ч��E-mail��ַ!');
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
if request("option")="���" then

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
response.Write("<SCRIPT LANGUAGE=JavaScript> alert('��������Ʒ��ͬ�����ڣ�����������������');</SCRIPT>")
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
response.Write("<SCRIPT LANGUAGE=JavaScript> alert('��ӳɹ�');</SCRIPT>")
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
                  color=#FF0000><strong>���˿ͻ�������ϱ༩</strong></font></TD>
      </TR>
      <TR> 
        <TD width=46% align="right">������</TD>
        <TD><input maxlength=100 size=30 id="allname" name="allname">
          *</TD>
      </TR>
      <TR> 
        <TD align="right">�Ա�</TD>
        <TD><INPUT maxLength=100 size=30 id="sex" name="sex">
          *</TD>
      </TR>
      <TR> 
        <TD align="right">���壺</TD>
        <TD><INPUT maxLength=100 size=30 id="nation" name="nation">
          *</TD>
      </TR>
      <TR> 
        <TD align="right">ԭ�壺</TD>
        <TD><INPUT maxLength=100 size=30 id="original" name="original">
          *</TD>
      </TR>
      <TR> 
        <TD align="right">�������£�</TD>
        <TD><INPUT maxLength=100 size=30 id="birth" name="birth">
          *</TD>
      </TR>
      <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							              email
<!------------------------------------------------------------------------------------------------------------------------------------ -->
      <TR> 
        <TD align="right">������λ��</TD>
        <TD><INPUT maxLength=40 size=30 id="workplace" name="workplace">
          *</TD>
      </TR>
      <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							              email
<!------------------------------------------------------------------------------------------------------------------------------------ -->
      <TR> 
        <TD align="right">ְ��</TD>
        <TD><INPUT maxLength=40 size=30 id="duty" name="duty">
          *</TD>
      </TR>
      <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							              email
<!------------------------------------------------------------------------------------------------------------------------------------ -->
      <TR> 
        <TD align="right">��Ȥ���ã�</TD>
        <TD> <INPUT maxLength=40 size=30 id="hobby" name="hobby">
          *</TD>
      </TR>
      <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							              USERID
<!----------------------------------------------------------------------------------------------------------------------------------- -->
      <TR> 
        <TD align="right">סַ��</TD>
        <TD> <INPUT maxLength=20 size=30 id="addr" name="addr">
          *</TD>
      </TR>
      <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							         PASSWORD
<!------------------------------------------------------------------------------------------------------------------------------------ -->
      <TR> 
        <TD align="right">�ʱࣺ</TD>
        <TD> <INPUT maxLength=20 size=30 id="zip" name="zip">
          *</TD>
      </TR>
      <TR> 
        <TD align="right">�绰��</TD>
        <TD> <INPUT maxLength=20 size=30 id="tel" name="tel">
          *</TD>
      </TR>
      <TR> 
        <TD align="right">E-mail��</TD>
        <TD> <INPUT maxLength=20 size=30 id="email" name="email">
          *</TD>
      </TR>
      <TR> 
        <TD height=32 align="right"> ��ע��</TD>
        <TD> <textarea name="memo" id="memo" cols="60" rows="10"></textarea> 
        </TD>
      </TR>
    
    <TR> 
      <TD height="50" colSpan=2 align="center"> <INPUT type=submit id="option" name="option" value="���"> 
        &nbsp;&nbsp; <INPUT type="button" id="goback" name="goback" value="����" onClick="history.go( -1 );"> 
      </TD>
    </TR>
</FORM>
  </TABLE>
<br>
<br>
</body>
</html>
