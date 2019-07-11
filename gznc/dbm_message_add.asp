<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">

<title></title>
<LINK href="dbm_css.css" type=text/css rel=stylesheet>
</head>
<body topmargin="0">
<script language="JavaScript">

function checkform()
{
	if (form_administrator.guest_id.value.length == 0) {
		alert("请选择发送对象");
		form_administrator.guest_id.focus();
		return false;
		}
	if (form_administrator.subject.value.length == 0) {
		alert("请输入类型名称");
		form_administrator.subject.focus();
		return false;
		}
	if (form_administrator.memo.value.length == 0) {
		alert("请选取类型所属");
		form_administrator.memo.focus();
		return false;
		}
	return true;
}

</script>

<!--#include file="dbm_adminconn.asp" -->






<%
if request("option")="add" then
guest_id=request("guest_id")
subject=request("subject")
subject=replace(subject,"'","''")
memo=request("memo")
memo=replace(memo,"'","''")
if memo="" then memo=" "

if session("user_adminlevel")=9 then
sql="insert into dbm_message (subject,user_id,guest_id,memo,inserttime) values ('"&subject&"',0,"&guest_id&",'"&memo&"','"&now()&"')"
else
sql="insert into dbm_message (subject,user_id,guest_id,memo,inserttime) values ('"&subject&"',"&session("user_id")&","&guest_id&",'"&memo&"','"&now()&"')"
end if
conn.Execute(sql)
response.Write("<SCRIPT LANGUAGE=JavaScript> alert('添加成功');</SCRIPT>")
end if
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
  
<TABLE width="100%" border=1 cellPadding=3 cellSpacing=0 bordercolor="#CCCCCC" >
  <FORM action="dbm_message_add.asp" method=post name="form_administrator" id="form_administrator" onSubmit="return checkform();">
    <TBODY>
      <TR align="center"> 
        <TD colspan="2"><font color="#FF0000"><strong>短信息添加</strong></font></TD>
      </TR>
<%
if session("user_adminlevel")=9 then
%>		
<input type="hidden" name="guest_id" id="guest_id" value=<%=session("user_id")%>>
<%
else
%>
      <TR> 
        <TD align="right">发送对象：</TD>
        <TD>
		<select name="guest_id" class="display_dropdown" id="guest_id">
                    <option selected value="">请选择要发信息的客户...</option>
                    <%
if session("user_adminlevel")=1 then
sql="select * from dbm_guest order by allname" 
else
sql="select * from dbm_guest where guest_id in ( select guest_id from dbm_ug where user_id="&session("user_id")&" )"
end if
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.Open sql,conn,1,1 
do while not rs.eof
%>
                    <option value=<%=rs("guest_id")%>><%=rs("allname")%> </option>
                    <%
rs.movenext
loop
rs.close
set rs=nothing
%>
                  </select>
				  </TD>
      </TR>
<%
end if
%>				  
      <TR> 
        <TD width="46%" align="right">主题：</TD>
        <TD bgColor=#ffffff><input name="subject" type="text" id="subject" size="30" maxlength="50">
          *</TD>
      </TR>
      <TR> 
        <TD align="right">短信息：</TD>
        <TD bgColor=#ffffff> <textarea name="memo" id="memo" cols="60" rows="10"></textarea></TD>
      </TR>
      <TR align="center"> 
        <TD colspan="2"><input type=submit id="option" name="option" value="add"> 
          &nbsp;&nbsp;&nbsp; <INPUT type="button" id="goback" name="goback" value="go back" onClick="history.go( -1 );"> 
        </TD>
      </TR>
    </TBODY>
  </FORM>
</TABLE>
