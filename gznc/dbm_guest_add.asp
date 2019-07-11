<script language="JavaScript">

function checkform()
{
	if (form_administrator.allname.value.length == 0) {
		alert("请填写名称");
		form_administrator.allname.focus();
		return false;
		}
	if(form_administrator.email.value.length > 0){
	if(! isMail(form_administrator.email.value)) {
		alert('请输入正确的邮件地址');
		form_administrator.email.focus();
		return false;
		}}
	if (form_administrator.gid.value.length == 0) {
		alert("请填写客户账号! ");
		form_administrator.gid.focus();
		return false;
		}
	if (form_administrator.pwd.value.length == 0) {
		alert(" 请填写密码! ");
		form_administrator.pwd.focus();
		return false;
		}
	if (form_administrator.pwd.value != form_administrator.cpwd.value) {
		alert(" 两次输入密码不相符，请从新输入 ");
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
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">

<title></title>
<LINK href="dbm_css.css" type=text/css rel=stylesheet>
</head>
<body topmargin="0">
<!--#include file="dbm_adminconn.asp" -->

<%
if request("option")="add" then

allname=trim(request("allname"))
allname=replace(allname,"'","''")
simname=trim(request("simname"))
simname=replace(simname,"'","''")

states=trim(request("states"))
states=replace(states,"'","''")

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
gid=trim(request("gid"))
gid=replace(gid,"'","''")
pwd=trim(request("pwd"))
pwd=replace(pwd,"'","''")
memo=trim(request("memo"))
memo=replace(memo,"'","''")
usql="select * from dbm_guest where allname='"&allname&"'"
set urs=server.createobject("ADODB.Recordset")
urs.open usql,conn,1,1
if not ( urs.eof or urs.bof ) then
response.Write("<SCRIPT LANGUAGE=JavaScript> alert('出错！此客户有同名存在，请用其它名字输入');</SCRIPT>")
urs.close
set urs=nothing
else
urs.close
set urs=nothing

if simname="" then simname=" "

if states="" then states=" "

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
sql="insert into dbm_guest (allname,simname,states, nation, native, born_time, graduation, special, job_now, job_old, job_rank, expert, awarded, standest, interest, addr,tel,fax,zip,email,contacter,gid,pwd,memo,inserttime) values ('"&allname&"','"&simname&"','"&states&"','"&nation&"','"&native&"','"&born_time&"','"&graduation&"','"&special&"','"&job_now&"','"&job_old&"','"&job_rank&"','"&expert&"','"&awarded&"','"&standest&"','"&interest&"','"&addr&"','"&tel&"','"&fax&"','"&zip&"','"&email&"','"&contacter&"','"&gid&"','"&pwd&"','"&memo&"','"&now()&"')"
conn.Execute(sql)

  If not Err Then
  	zrootpath = Server.MapPath(".")
    Response.Redirect("dbm_folder_option.asp?Action=NewFolder&FName="+zrootpath+"\sosoon\"+allname+"&zurl=dbm_guest_add.asp")
  End If		

response.Write("<SCRIPT LANGUAGE=JavaScript> alert('添加成功');</SCRIPT>")
end if
end if
%>
<!----------------------------------------------------------------------------------------------------------------------
<!--                                                   ERROR MSGS
<!-------------------------------------------------------------------------------------------------------------------- -->
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
  <FORM action="dbm_guest_add.asp" method="post" name="form_administrator" id="form_administrator" onSubmit="return checkform();">
<TABLE width=100% border=1 cellPadding=3 cellSpacing=0 bordercolor="#C0C0C0">
    <TBODY>
      <TR align="center"> 
        <TD height="30" colspan="2"><strong><font color="#FF0000" size="+1">客户添加</font></strong> 
        </TD>
      </TR>
      <TR> 
        <TD width="46%" align="right"> 名称(公司或个人)：</TD>
        <TD> <INPUT maxLength=100 size=30 id="allname" name="allname">
          * </TD>
      </TR>
      <TR> 
        <TD align="right"> 是否新项目客户：</TD>
        <TD> <select name="states" id="states">
            <option value="新项目" selected>新项目</option>
            <option value="旧项目">旧项目</option>
          </select> </TD>
      </TR>
      <TR> 
        <TD align="right"> 地址：</TD>
        <TD> <INPUT maxLength=100 size=30 id="addr" name="addr"> </TD>
      </TR>
      <TR> 
        <TD align="right"> 电话：</TD>
        <TD> <INPUT maxLength=100 size=30 id="tel" name="tel"> </TD>
      </TR>
      <TR> 
        <TD align="right"> 传真：</TD>
        <TD> <INPUT maxLength=100 size=30 id="fax" name="fax"> </TD>
      </TR>
      <TR> 
        <TD height="25" align="right"> 邮编：</TD>
        <TD> <INPUT maxLength=100 size=30 id="zip" name="zip"> </TD>
      </TR>
      <TR> 
        <TD align="right"> E-mail：</TD>
        <TD> <INPUT maxLength=40 size=30 id="email" name="email"> </TD>
      </TR>
      <TR> 
        <TD align="right"> 客户账号：</TD>
        <TD> <INPUT maxLength=50 size=30 id="gid" name="gid"> *  </TD>
      </TR>
      <TR> 
        <TD height="27" align="right"> 密码：</TD>
        <TD> <INPUT maxLength=50 size=30 id="pwd" name="pwd"> *  </TD>
      </TR>
      <TR> 
        <TD align="right"> 重复密码：</TD>
        <TD> <INPUT maxLength=50 size=30 id="cpwd" name="cpwd"> *  </TD>
      </TR>
      <TR> 
        <TD height=32 align="right"> 备注：</TD>
        <TD> <textarea name="memo" id="memo" cols="60" rows="10"></textarea> 
        </TD>
      </TR>
      <TR align="center"> 
        <TD height="50" colSpan=2><INPUT type="submit" id="option" name="option" value="add">
        </TD>
      </TR>
    </TBODY>
    <INPUT maxLength=100 size=30 id="contacter" name="contacter" type="hidden">
</TABLE>
  </FORM>
<br>
<br>
</body>
</html>
