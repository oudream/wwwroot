<script language="JavaScript">

function checkform()
{
	if (form_administrator.allname.value.length == 0) {
		alert("������ȫ��");
		form_administrator.allname.focus();
		return false;
		}
	if(form_administrator.email.value.length > 0){
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
if request("option")="add" then

allname=trim(request("allname"))
allname=replace(allname,"'","''")
simname=trim(request("simname"))
simname=replace(simname,"'","''")

sex=trim(request("sex"))
sex=replace(sex,"'","''")

nation=trim(request("nation"))
nation=replace(nation,"'","''")

native=trim(request("native"))
native=replace(native,"'","''")

born_time=trim(request("born_time"))
born_time=replace(born_time,"'","''")

graduation=trim(request("graduation"))
graduation=replace(graduation,"'","''")

special=trim(request("special"))
special=replace(special,"'","''")

job_now=trim(request("job_now"))
job_now=replace(job_now,"'","''")

job_old=trim(request("job_old"))
job_old=replace(job_old,"'","''")
            
job_rank=trim(request("job_rank"))
job_rank=replace(job_rank,"'","''")

expert=trim(request("expert"))
expert=replace(expert,"'","''")

awarded=trim(request("awarded"))
awarded=replace(awarded,"'","''")

standest=trim(request("standest"))
standest=replace(standest,"'","''")

interest=trim(request("interest"))
interest=replace(interest,"'","''")

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
contacter=trim(request("contacter"))
contacter=replace(contacter,"'","''")
memo=trim(request("memo"))
memo=replace(memo,"'","''")
usql="select * from art_artist where allname='"&allname&"'"
set urs=server.createobject("ADODB.Recordset")
urs.open usql,conn,1,1
if not ( urs.eof or urs.bof ) then
response.Write("<SCRIPT LANGUAGE=JavaScript> alert('��������������ͬ�����ڣ�����������������');</SCRIPT>")
urs.close
set urs=nothing
else
urs.close
set urs=nothing

if simname="" then simname=" "

if sex="" then sex=" "

if nation="" then nation=" "

if native="" then native=" "

if born_time="" then born_time=" "

if graduation="" then graduation=" "

if special="" then special=" "

if job_now="" then job_now=" "

if job_old="" then job_old=" "

if job_rank="" then job_rank=" "

if expert="" then expert=" "

if awarded="" then awarded=" "

if standest="" then standest=" "

if interest="" then interest=" "

if addr="" then addr=" "
if tel="" then tel=" "
if fax="" then fax=" "
if zip="" then zip=" "
if email="" then email=" "
if contacter="" then contacter=" "
if memo="" then memo=" "
sql="insert into art_artist (allname,simname,sex, nation, native, born_time, graduation, special, job_now, job_old, job_rank, expert, awarded, standest, interest, addr,tel,fax,zip,email,contacter,memo,inserttime) values ('"&allname&"','"&simname&"','"&sex&"','"&nation&"','"&native&"','"&born_time&"','"&graduation&"','"&special&"','"&job_now&"','"&job_old&"','"&job_rank&"','"&expert&"','"&awarded&"','"&standest&"','"&interest&"','"&addr&"','"&tel&"','"&fax&"','"&zip&"','"&email&"','"&contacter&"','"&memo&"','"&now()&"')"
conn.Execute(sql)
response.Write("<SCRIPT LANGUAGE=JavaScript> alert('��ӳɹ�');</SCRIPT>")
end if
end if
%>
<!----------------------------------------------------------------------------------------------------------------------
<!--                                                   ERROR MSGS
<!-------------------------------------------------------------------------------------------------------------------- -->
<table width="75%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
<TABLE width=100% 
                  border=1 cellPadding=3 cellSpacing=0 bordercolor="#C0C0C0">
  <FORM action="art_artist_add.asp" method="post" name="form_administrator" id="form_administrator" onSubmit="return checkform();">
    <TBODY>
      <TR align="center"> 
        <TD height="30" colspan="2"><strong><font color="#FF0000" size="+1">���������</font></strong> 
        </TD>
      </TR>
      <TR> 
        <TD width="46%" align="right"> ������</TD>
        <TD> <INPUT maxLength=100 size=30 id="allname" name="allname">
          * </TD>
      </TR>
      <TR> 
        <TD align="right"> �źţ�</TD>
        <TD> <INPUT maxLength=100 size=30 id="simname" name="simname"> </TD>
      </TR>
      <TR> 
        <TD align="right"> �Ա�</TD>
        <TD> <select name="sex" id="sex">
            <option value="��" selected>��</option>
            <option value="Ů">Ů</option>
          </select> </TD>
      </TR>
      <TR> 
        <TD align="right"> ���壺</TD>
        <TD> <INPUT maxLength=100 size=30 id="nation" name="nation"> </TD>
      </TR>
      <TR> 
        <TD align="right"> ԭ����</TD>
        <TD> <INPUT maxLength=100 size=30 id="native" name="native"> </TD>
      </TR>
      <TR> 
        <TD align="right"> �������£�</TD>
        <TD> <INPUT maxLength=100 size=30 id="born_time" name="born_time">
          ��.1950-12-12</TD>
      </TR>
      <TR> 
        <TD align="right"> ��ҵԺУ��</TD>
        <TD> <INPUT maxLength=100 size=30 id="graduation" name="graduation"> </TD>
      </TR>
      <TR> 
        <TD align="right"> ��ѧרҵ��</TD>
        <TD> <INPUT maxLength=100 size=30 id="special" name="special"> </TD>
      </TR>
      <TR> 
        <TD align="right"> ����ְ��</TD>
        <TD> <INPUT maxLength=100 size=30 id="job_now" name="job_now"> </TD>
      </TR>
      <TR> 
        <TD align="right"> ����ְ��</TD>
        <TD> <INPUT maxLength=100 size=30 id="job_old" name="job_old"> </TD>
      </TR>
      <TR> 
        <TD align="right"> ְ�ƣ�</TD>
        <TD> <INPUT maxLength=100 size=30 id="job_rank" name="job_rank"> </TD>
      </TR>
      <TR> 
        <TD align="right"> �ó���</TD>
        <TD> <INPUT maxLength=100 size=30 id="expert" name="expert"> </TD>
      </TR>
      <TR> 
        <TD align="right"> �����</TD>
        <TD> <INPUT maxLength=100 size=30 id="awarded" name="awarded"> </TD>
      </TR>
      <TR> 
        <TD align="right"> ������Ʒ��</TD>
        <TD> <INPUT maxLength=100 size=30 id="standest" name="standest"> </TD>
      </TR>
      <TR> 
        <TD align="right"> ��Ȥ���ã�</TD>
        <TD> <INPUT maxLength=100 size=30 id="interest" name="interest"> </TD>
      </TR>
      <TR> 
        <TD align="right"> ��ַ��</TD>
        <TD> <INPUT maxLength=100 size=30 id="addr" name="addr"> </TD>
      </TR>
      <TR> 
        <TD align="right"> �绰��</TD>
        <TD> <INPUT maxLength=100 size=30 id="tel" name="tel"> </TD>
      </TR>
      <TR> 
        <TD align="right"> ���棺</TD>
        <TD> <INPUT maxLength=100 size=30 id="fax" name="fax"> </TD>
      </TR>
      <TR> 
        <TD align="right"> �ʱࣺ</TD>
        <TD> <INPUT maxLength=100 size=30 id="zip" name="zip"> </TD>
      </TR>
      <TR> 
        <TD align="right"> E-mail��</TD>
        <TD> <INPUT maxLength=40 size=30 id="email" name="email"> </TD>
      </TR>
      <TR> 
        <TD height=32 align="right"> ��ע��</TD>
        <TD> <textarea name="memo" id="memo" cols="60" rows="10"></textarea> 
        </TD>
      </TR>
      <TR align="center"> 
        <TD height="50" colSpan=2> <INPUT type=submit id="options" name="options" value="���"> 
        </TD>
      </TR>
    </TBODY>
    <INPUT type="hidden" id="option" name="option" value="add">
    <INPUT maxLength=100 size=30 id="contacter" name="contacter" type="hidden">
  </FORM>
</TABLE>
<br>
<br>
</body>
</html>
