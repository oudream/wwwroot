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
	if (form_administrator.note_time.value.length == 0) {
		alert("请输入题摘");
		form_administrator.note_time.focus();
		return false;
		}
	if (form_administrator.memo.value.length == 0) {
		alert("请输入纪事");
		form_administrator.memo.focus();
		return false;
		}
	return true;
}

</script>

<!--#include file="art_adminconn.asp" -->






<%
if request("option")="add" then
note_time=request("note_time")
note_time=replace(note_time,"'","''")
memo=request("memo")
memo=replace(memo,"'","''")
if memo="" then memo=" "

sql="insert into art_note_old (note_time,user_id,memo,inserttime) values ('"&note_time&"',"&session("user_id")&",'"&memo&"','"&now()&"')"
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
<FORM action="art_note_old_add.asp" method=post name="form_administrator" id="form_administrator" onSubmit="return checkform();">
    <TBODY>
      <TR align="center"> 
        <TD height="30" colspan="2"><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#FF0000><strong>美协纪事添加</strong></font></TD>
      </TR>
      <TR>
        <TD colspan="2">&nbsp;</TD>
      </TR>
      <TR> 
        <TD width="46%" align="right">时间：</TD>
        <TD bgColor=#ffffff><input name="note_time" type="text" id="note_time" size="30" maxlength="50">
          *</TD>
      </TR>
      <TR> 
        <TD align="right">纪事：</TD>
        <TD bgColor=#ffffff> <textarea name="memo" id="memo" cols="60" rows="10"></textarea>
          *</TD>
      </TR>
      <TR align="center"> 
        <TD height="50" colspan="2"><input type=submit id="option" name="option" value="add"> 
          &nbsp;&nbsp;&nbsp; <INPUT type="button" id="goback" name="goback" value="go back" onClick="history.go( -1 );"> 
        </TD>
      </TR>
    </TBODY>
</FORM>
  </TABLE>
