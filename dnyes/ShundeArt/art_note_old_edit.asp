<script language="JavaScript">

function checkform()
{
	if (form_administrator.note_time.value.length == 0) {
		alert("请输入备忘名称");
		form_administrator.uname.focus();
		return false;
		}
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
if trim(request("option"))="del" then
conn.execute "delete * from art_note_old where id=" & request("zid")
response.Redirect("art_note_old_view.asp?option_id_del=succ")
end if
%>
<%
if request("option")="edit" then

note_time=request("note_time")
note_time=replace(note_time,"'","''")
memo=request("memo")
memo=replace(memo,"'","''")
if memo="" then memo=" "

conn.Execute("update art_note_old set note_time='"&note_time&"', memo='"&memo&"' where id=" & request("zid"))
response.Write("<SCRIPT LANGUAGE=JavaScript> alert('编辑成功');</SCRIPT>")
end if
%>
<%
sql="select * from art_note_old where id="&request("zid") 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1 

if rs.eof or rs.bof then 
response.Write("<SCRIPT LANGUAGE=JavaScript> alert('没有所需要的');</SCRIPT>")
response.End()
end if
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
  
<TABLE width=100% 
                  border=1 cellPadding=3 cellSpacing=0 bordercolor="#CCCCCC">
  <FORM action="art_note_old_edit.asp" method="post" name="form_administrator" id="form_administrator" onSubmit="return checkform();">
    <TBODY>
      <TR align="center"> 
        <TD height=30 colspan="2"><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#FF0000><strong>备忘编缉</strong></font></TD>
      </TR>
      <TR align="center">
        <TD colspan="2">&nbsp;</TD>
      </TR>
      <TR> 
        <TD height=23 align="right"><font color="#000000" face="Verdana, Arial, Helvetica, sans-serif">备忘名称：</font></TD>
        <TD width=396><font 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 > 
          <input maxlength=100 size=30 id="note_time" name="note_time" value="<%=rs("note_time")%>">
          * </font></TD>
      </TR>
      <TR> 
        <TD width="150" height=32 align="right"> <font color="#000000" face="Verdana, Arial, Helvetica, sans-serif">备忘：</font></TD>
        <TD><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 > 
          <textarea name="memo" id="memo" cols="60" rows="10"><%=rs("memo")%></textarea>
          </FONT></TD>
      </TR>
      <TR align="center"> 
        <TD height="50" colSpan=2> <P><STRONG><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000066 size=2> 
            <INPUT type=submit id="option" name="option" value="edit">
            &nbsp;&nbsp;&nbsp; 
            <input type="submit" id="option" name="option" value="del" onClick="return confirm('确定删除此项吗？')">
            &nbsp;&nbsp;&nbsp; 
            <INPUT type="button" id="goback" name="goback" value="go back" onClick="history.go( -1 );">
            <input type="hidden" id="zid" name="zid" value="<%=request("zid")%>">
            </FONT></STRONG></P></TD>
      </TR>
    </TBODY>
  </FORM>
</TABLE>
<br>
<br>
</body>
</html>
