<script language="JavaScript">

function checkform()
{
	if (form_administrator.subject.value.length == 0) {
		alert("��������������");
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
conn.execute "delete * from art_message where id=" & request("zid")
response.Redirect("message_view.asp?option_id_del=succ")
end if
%>
<%
if request("option")="edit" then

subject=request("subject")
subject=replace(subject,"'","''")
memo=request("memo")
memo=replace(memo,"'","''")
if memo="" then memo=" "

conn.Execute("update art_message set subject='"&subject&"', memo='"&memo&"' where id=" & request("zid"))
response.Write("<SCRIPT LANGUAGE=JavaScript> alert('�༭�ɹ�');</SCRIPT>")
end if
%>
<%
sql="select * from art_message where id="&request("zid") 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1 

if rs.eof or rs.bof then 
response.Write("<SCRIPT LANGUAGE=JavaScript> alert('û������Ҫ��');</SCRIPT>")
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
<FORM action="art_message_edit.asp" method="post" name="form_administrator" id="form_administrator" onSubmit="return checkform();">
    <TBODY>
      <TR align="center"> 
        <TD height="30" colspan="2"><strong><font color="#FF0000"> ���Ա༩</DIV></font></strong></TD>
      </TR>
      <TR>
        <TD height=11 colspan="2">&nbsp;</TD>
      </TR>
      <TR> 
        <TD width="46%" align="right">�������ƣ�</TD>
        <TD> 
          <input maxlength=100 size=30 id="subject" name="subject" value="<%=rs("subject")%>">
          * </TD>
      </TR>
      <TR> 
        <TD align="right"> ���ԣ�</DIV></TD>
        <TD> 
          <textarea name="memo" id="memo" cols="60" rows="10"><%=rs("memo")%></textarea>
          </TD>
      </TR>
      <TR align="center"> 
        <TD height="50" colSpan=2>  
            <INPUT type=submit id="option" name="option" value="edit">
            &nbsp;&nbsp;&nbsp; 
            <input type="submit" id="option" name="option" value="del" onClick="return confirm('ȷ��ɾ��������')">
            &nbsp;&nbsp;&nbsp; 
            <INPUT type="button" id="goback" name="goback" value="go back" onClick="history.go( -1 );">
            <input type="hidden" id="zid" name="zid" value="<%=request("zid")%>">
            </TD>
      </TR>
    </TBODY>
</FORM>
  </TABLE>
<br>
<br>
</body>
</html>
