<script language="JavaScript">

function checkform()
{
	if (form_administrator.allname.value.length == 0) {
		alert("请填写名称");
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
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">

<title></title>
</head>
<body topmargin="0">
<!--#include file="dbm_adminconn.asp" -->






<%
if request("option")="编辑" then

allname=trim(request("allname"))
allname=replace(allname,"'","''")

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
if allname<>"" then
usql="select * from dbm_guest where ( allname='"&allname&"' ) and guest_id<>" & request("guest_id")
set urs=server.createobject("ADODB.Recordset")
urs.open usql,conn,1,1
if not ( urs.eof or urs.bof ) then
response.Write("<SCRIPT LANGUAGE=JavaScript> alert('出错！此客户有同名存在，请用其它名字输入');</SCRIPT>")
urs.close
set urs=nothing
else
urs.close
set urs=nothing

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
conn.Execute("update dbm_guest set allname='"&allname&"',states='"&states&"',nation='"&nation&"',native='"&native&"',born_time='"&born_time&"',graduation='"&graduation&"',special='"&special&"',job_now='"&job_now&"',job_old='"&job_old&"',job_rank='"&job_rank&"',expert='"&expert&"',awarded='"&awarded&"',standest='"&standest&"',interest='"&interest&"',addr='"&addr&"',tel='"&tel&"',fax='"&fax&"',zip='"&zip&"',email='"&email&"',contacter='"&contacter&"',gid='"&gid&"',pwd='"&pwd&"',memo='"&memo&"' where guest_id=" & request("guest_id"))
response.Write("<SCRIPT LANGUAGE=JavaScript> alert('编辑成功');</SCRIPT>")
end if
end if
end if
%>
<%
sql="select * from dbm_guest where guest_id="&request("guest_id") 
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
  <FORM action="dbm_guest_edit.asp" method="post" name="form_administrator" id="form_administrator" onSubmit="return checkform();">
    <TBODY>
      <TR align="center"> 
        <TD height=30 colspan="3"><font color="#FF0000" size="+1"><strong>客户编缉</strong></font></TD>
      </TR>
      <TR> 
        <TD width="43%" align="right"> 名称(公司或个人)：</TD>
        <TD width="22%"> <%=rs("allname")%>
          <input type="hidden" id="allname" name="allname" value="<%=rs("allname")%>">
        </TD>
        <TD width="35%">* 此项至关重要，不能修改</TD>
      </TR>
      <TR> 
        <TD align="right"> 是否新项目客户：</TD>
        <TD colspan="2"> <select name="states" id="states">
            <option value="<%=rs("states")%>" selected><%=rs("states")%></option>
            <option value="new">新项目</option>
            <option value="old">旧项目</option>
          </select> </TD>
      </TR>
      <TR> 
        <TD align="right"> 地址：</TD>
        <TD colspan="2"> <INPUT maxLength=100 size=30 id="addr" name="addr" value="<%=rs("addr")%>"> 
        </TD>
      </TR>
      <TR> 
        <TD align="right"> 电话：</TD>
        <TD colspan="2"> <INPUT maxLength=100 size=30 id="tel" name="tel" value="<%=rs("tel")%>"> 
        </TD>
      </TR>
      <TR> 
        <TD align="right"> 传真：</TD>
        <TD colspan="2"> <INPUT maxLength=100 size=30 id="fax" name="fax" value="<%=rs("fax")%>"> 
        </TD>
      </TR>
      <TR> 
        <TD align="right"> 邮编：</TD>
        <TD colspan="2"> <INPUT maxLength=100 size=30 id="zip" name="zip" value="<%=rs("zip")%>"> 
        </TD>
      </TR>
      <TR> 
        <TD align="right"> E-mail：</TD>
        <TD colspan="2"> <INPUT maxLength=40 size=30 id="email" name="email" value="<%=rs("email")%>"> 
        </TD>
      </TR>
      <TR> 
        <TD align="right"> 客户账号：</TD>
        <TD colspan="2"> <INPUT maxLength=50 size=30 id="gid" name="gid" value="<%=rs("gid")%>">
          * </TD>
      </TR>
      <TR> 
        <TD align="right"> 密码：</TD>
        <TD colspan="2"> <INPUT maxLength=50 size=30 id="pwd" name="pwd" value="<%=rs("pwd")%>">
          * </TD>
      </TR>
      <TR> 
        <TD align="right"> 重复密码：</TD>
        <TD colspan="2"> <INPUT maxLength=50 size=30 id="cpwd" name="cpwd" value="<%=rs("pwd")%>">
          * </TD>
      </TR>
      <TR> 
        <TD align="right"> 备注：</TD>
        <TD colspan="2"> <textarea name="memo" id="memo" cols="60" rows="10"><%=rs("memo")%></textarea> 
        </TD>
      </TR>
      <TR> 
        <TD height="50" colSpan=3> <INPUT type=submit id="option" name="option" value="编辑"> 
          &nbsp;&nbsp;&nbsp; <INPUT type="button" id="goback" name="goback" value="返回" onClick="history.go( -1 );"> 
          <input type="hidden" id="guest_id" name="guest_id" value="<%=request("guest_id")%>"> 
        </TD>
      </TR>
    </TBODY>
    <INPUT maxLength=100 size=30 id="contacter" name="contacter" type="hidden">
  </FORM>
</TABLE>
</body>
</html>
