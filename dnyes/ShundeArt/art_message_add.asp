<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=site.css rel=stylesheet>
<title></title>
</head>
<body topmargin="0">
<script language="JavaScript">

function checkform()
{
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

<!--#include file="art_adminconn.asp" -->






<%
if request("option")="add" then
subject=request("subject")
subject=replace(subject,"'","''")
memo=request("memo")
memo=replace(memo,"'","''")
if memo="" then memo=" "

sql="insert into art_message (subject,user_id,memo,inserttime) values ('"&subject&"',"&session("user_id")&",'"&memo&"','"&now()&"')"
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
<FORM action="art_message_add.asp" method=post name="form_administrator" id="form_administrator" onSubmit="return checkform();">
    <TBODY>
      <TR align="center"> 
        <TD colspan="2"><font color="#FF0000"><strong>留言添加</strong></font></TD>
      </TR>
      <TR>
        <TD colspan="2">&nbsp;</TD>
      </TR>
      <TR> 
        <TD width="46%" align="right">主题：</TD>
        <TD bgColor=#ffffff><input name="subject" type="text" id="subject" size="30" maxlength="50">
          *</TD>
      </TR>
      <TR> 
        <TD align="right">留言：</TD>
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
