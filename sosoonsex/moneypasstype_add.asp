<script language="JavaScript">

function checkform()
{
	if (form_administrator.mptname.value.length == 0) {
		alert("��������������");
		form_administrator.mptname.focus();
		return false;
		}
	if (form_administrator.sign.value.length == 0) {
		alert("��ѡȡ��������");
		form_administrator.sign.focus();
		return false;
		}
	return true;
}

</script>

<!--#include file="adminconn.asp" -->
<%
if request("option")="add" then
mptname=request("mptname")
mptname=replace(mptname,"'","''")
sign=request("sign")
sign=replace(sign,"'","''")
memo=request("memo")
memo=replace(memo,"'","''")
if memo="" then memo=" "
usql="select * from moneypasstype where mptname='"&mptname&"'"
set urs=server.createobject("ADODB.Recordset")
urs.open usql,conn,1,1
if not ( urs.eof or urs.bof ) then
response.Write("<SCRIPT LANGUAGE=JavaScript> alert('������������ͬ�����ڣ�����������������');</SCRIPT>")
urs.close
set urs=nothing
else
urs.close
set urs=nothing
sql="insert into moneypasstype (mptname,sign,memo) values ('"&mptname&"','"&sign&"','"&memo&"')"
conn.Execute(sql)
response.Write("<SCRIPT LANGUAGE=JavaScript> alert('��ӳɹ�');</SCRIPT>")
end if
end if
%>
<H2><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000066>��Ǯ��ͨ�������</FONT></H2>
<FORM action="moneypasstype_add.asp" method=post name="form_administrator" id="form_administrator" onSubmit="return checkform();">
                        
  <TABLE width="550" border=0 cellPadding=5 cellSpacing=1 >
    <TBODY>
      <TR> 
        <TD width="150">��������</TD>
        <TD width="71%" bgColor=#ffffff><input name="mptname" type="text" id="mptname" size="30" maxlength="50">
          *</TD>
      </TR>
      <TR> 
        <TD width="150">��������</TD>
        <TD width="71%" bgColor=#ffffff>
		<select name="sign" id="sign">
            <option value="">��ѡ�񡭡�</option>
            <option value="plus">����</option>
            <option value="minus">֧��</option>
          </select>
          *</TD>
      </TR>
      <TR> 
        <TD width="150">��ע</TD>
        <TD width="71%" bgColor=#ffffff> <textarea name="memo" id="memo" cols="60" rows="10"></textarea></TD>
      </TR>
      <TR> 
        <TD colspan="2"><input type=submit id="option" name="option" value="add"> 
        </TD>
      </TR>
    </TBODY>
  </TABLE>
</FORM>
