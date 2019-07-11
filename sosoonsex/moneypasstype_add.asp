<script language="JavaScript">

function checkform()
{
	if (form_administrator.mptname.value.length == 0) {
		alert("请输入类型名称");
		form_administrator.mptname.focus();
		return false;
		}
	if (form_administrator.sign.value.length == 0) {
		alert("请选取类型所属");
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
response.Write("<SCRIPT LANGUAGE=JavaScript> alert('出错！此类型有同名存在，请用其它名字输入');</SCRIPT>")
urs.close
set urs=nothing
else
urs.close
set urs=nothing
sql="insert into moneypasstype (mptname,sign,memo) values ('"&mptname&"','"&sign&"','"&memo&"')"
conn.Execute(sql)
response.Write("<SCRIPT LANGUAGE=JavaScript> alert('添加成功');</SCRIPT>")
end if
end if
%>
<H2><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000066>金钱流通类型添加</FONT></H2>
<FORM action="moneypasstype_add.asp" method=post name="form_administrator" id="form_administrator" onSubmit="return checkform();">
                        
  <TABLE width="550" border=0 cellPadding=5 cellSpacing=1 >
    <TBODY>
      <TR> 
        <TD width="150">类型名称</TD>
        <TD width="71%" bgColor=#ffffff><input name="mptname" type="text" id="mptname" size="30" maxlength="50">
          *</TD>
      </TR>
      <TR> 
        <TD width="150">类型所属</TD>
        <TD width="71%" bgColor=#ffffff>
		<select name="sign" id="sign">
            <option value="">请选择……</option>
            <option value="plus">收入</option>
            <option value="minus">支出</option>
          </select>
          *</TD>
      </TR>
      <TR> 
        <TD width="150">备注</TD>
        <TD width="71%" bgColor=#ffffff> <textarea name="memo" id="memo" cols="60" rows="10"></textarea></TD>
      </TR>
      <TR> 
        <TD colspan="2"><input type=submit id="option" name="option" value="add"> 
        </TD>
      </TR>
    </TBODY>
  </TABLE>
</FORM>
