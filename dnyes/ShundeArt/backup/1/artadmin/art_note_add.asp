<script language="JavaScript">

function checkform()
{
	if (form_administrator.subject.value.length == 0) {
		alert("��������ժ");
		form_administrator.subject.focus();
		return false;
		}
	if (form_administrator.memo.value.length == 0) {
		alert("��ѡȡ��������");
		form_administrator.memo.focus();
		return false;
		}
	return true;
}

</script>

<!--#include file="adminconn.asp" -->
<%
if request("option")="add" then
subject=request("subject")
subject=replace(subject,"'","''")
memo=request("memo")
memo=replace(memo,"'","''")
if memo="" then memo=" "

sql="insert into art_note (subject,user_id,memo,inserttime) values ('"&subject&"',"&session("user_id")&",'"&memo&"','"&now()&"')"
conn.Execute(sql)
response.Write("<SCRIPT LANGUAGE=JavaScript> alert('��ӳɹ�');</SCRIPT>")
end if
%>
<H2><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000066>���ű������</FONT></H2>
<FORM action="art_note_add.asp" method=post name="form_administrator" id="form_administrator" onSubmit="return checkform();">
                        
  <TABLE width="550" border=0 cellPadding=5 cellSpacing=1 >
    <TBODY>
      <TR> 
        <TD width="150">��ժ</TD>
        <TD width="71%" bgColor=#ffffff><input name="subject" type="text" id="subject" size="30" maxlength="50">
          *</TD>
      </TR>
      <TR> 
        <TD width="150">����</TD>
        <TD width="71%" bgColor=#ffffff> <textarea name="memo" id="memo" cols="60" rows="10"></textarea></TD>
      </TR>
      <TR> 
        <TD colspan="2"><input type=submit id="option" name="option" value="add">            &nbsp;&nbsp;&nbsp; 
            <INPUT type="button" id="goback" name="goback" value="go back" onClick="history.go( -1 );">
 
        </TD>
      </TR>
    </TBODY>
  </TABLE>
</FORM>
