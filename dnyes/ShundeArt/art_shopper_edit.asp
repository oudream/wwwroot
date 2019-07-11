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

<!--#include file="art_adminconn.asp" -->

<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=site.css rel=stylesheet>
<title></title>
</head>
<body topmargin="0">

<%
if request("option")="修改" then

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

usql="select * from art_shopper where allname='" & allname & "' and art_shopper_id<>" & request("art_shopper_id") 
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

conn.Execute("update art_shopper set allname='"&allname&"',addr='"&addr&"',tel='"&tel&"',sex='"&sex&"',zip='"&zip&"',email='"&email&"',nation='"&nation&"',original='"&original&"',birth='"&birth&"',duty='"&duty&"',workplace='"&workplace&"',hobby='"&hobby&"',memo='"&memo&"' where art_shopper_id=" & request("art_shopper_id"))
response.Write("<SCRIPT LANGUAGE=JavaScript> alert('修改成功!');</SCRIPT>")
End if
End if
%>
<%
sql="select * from art_shopper where art_shopper_id="&request("art_shopper_id") 
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
<FORM action="art_shopper_edit.asp" method="post" name="form_shopper" id="form_shopper" onSubmit="return checkform();">
    <TBODY>
      <TR align="center"> 
        <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							              NAME
<!------------------------------------------------------------------------------------------------------------------------------------ -->
        <TD height=30 colspan="2"><strong><font color="#FF0000">画廊资料</font></strong></TD>
      </TR>
      <TR> 
        <TD align="right">姓名：</TD>
        <TD><input maxlength=100 size=30 id="allname" name="allname" value="<%=rs("allname")%>">
          *</TD>
      </TR>
      <TR> 
        <TD width=46% align="right">性别：</TD>
        <TD><INPUT maxLength=100 size=30 id="sex" name="sex" value="<%=rs("sex")%>">
          *</TD>
      </TR>
      <TR> 
        <TD width=46% align="right">民族：</TD>
        <TD><INPUT maxLength=100 size=30 id="nation" name="nation" value="<%=rs("nation")%>">
          *</TD>
      </TR>
      <TR> 
        <TD width=46% align="right">原藉：</TD>
        <TD><INPUT maxLength=100 size=30 id="original" name="original" value="<%=rs("original")%>">
          *</TD>
      </TR>
      <TR> 
        <TD width=46% align="right">出生年月：</TD>
        <TD><INPUT maxLength=100 size=30 id="birth" name="birth" value="<%=rs("birth")%>">
          *</TD>
      </TR>
      <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							              email
<!------------------------------------------------------------------------------------------------------------------------------------ -->
      <TR> 
        <TD width=46% align="right">工作单位：</TD>
        <TD><INPUT maxLength=40 size=30 id="workplace" name="workplace" value="<%=rs("workplace")%>">
          *</TD>
      </TR>
      <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							              email
<!------------------------------------------------------------------------------------------------------------------------------------ -->
      <TR> 
        <TD width=46% align="right">职务：</TD>
        <TD><INPUT maxLength=40 size=30 id="duty" name="duty" value="<%=rs("duty")%>">
          *</TD>
      </TR>
      <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							              email
<!------------------------------------------------------------------------------------------------------------------------------------ -->
      <TR> 
        <TD width=46% align="right">兴趣爱好：</TD>
        <TD> <INPUT maxLength=40 size=30 id="hobby" name="hobby" value="<%=rs("hobby")%>">
          *</TD>
      </TR>
      <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							              USERID
<!----------------------------------------------------------------------------------------------------------------------------------- -->
      <TR> 
        <TD width=46% align="right">住址：</TD>
        <TD> <INPUT maxLength=20 size=30 id="addr" name="addr" value="<%=rs("addr")%>">
          *</TD>
      </TR>
      <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							         PASSWORD
<!------------------------------------------------------------------------------------------------------------------------------------ -->
      <TR> 
        <TD width=46% align="right">邮编：</TD>
        <TD> <INPUT maxLength=20 size=30 id="zip" name="zip" value="<%=rs("zip")%>">
          *</TD>
      </TR>
      <TR> 
        <TD width=46% align="right">电话：</TD>
        <TD> <INPUT maxLength=20 size=30 id="tel" name="tel" value="<%=rs("tel")%>">
          *</TD>
      </TR>
      <TR> 
        <TD width=46% align="right">E-mail：</TD>
        <TD> <INPUT maxLength=20 size=30 id="email" name="email" value="<%=rs("email")%>">
          *</TD>
      </TR>
      <TR> 
        <TD height=32 align="right"> 备注：</TD>
        <TD> <textarea name="memo" id="memo" cols="60" rows="10"><%=rs("memo")%></textarea> 
        </TD>
      </TR>
      <TR align="center"> 
        <TD height="50" colSpan=2> <INPUT type=submit id="option" name="option" value="修改"> 
          &nbsp;&nbsp; <INPUT type="button" id="goback" name="goback" value="返回" onClick="history.go( -1 );"> 
        </TD>
      </TR>
    </TBODY>
            <INPUT type="hidden" id="art_shopper_id" name="art_shopper_id" value="<%=request("art_shopper_id")%>">
</FORM>
  </TABLE>
</body>
</html>