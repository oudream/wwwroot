<script language="JavaScript">

function checkform()
{
	if (form_administrator.sname.value.length == 0) {
		alert("请输入种类名称");
		form_administrator.sname.focus();
		return false;
		}
	return true;
}

</script>

<!--#include file="adminconn.asp" -->
<%
if request("option")="add" then
sname=trim(request("sname"))
sname=replace(sname,"'","''")
memo=trim(request("memo"))
memo=replace(memo,"'","''")
memo=trim(request("memo"))
memo=replace(memo,"'","''")
if memo="" then memo=" "
usql="select * from sort where sname='"&sname&"'"
set urs=server.createobject("ADODB.Recordset")
urs.open usql,conn,1,1
if not ( urs.eof or urs.bof ) then
response.Write("<SCRIPT LANGUAGE=JavaScript> alert('出错！此种类有同名存在，请用其它名字输入');</SCRIPT>")
urs.close
set urs=nothing
else
urs.close
set urs=nothing
sql="insert into sort (sname,memo) values ('"&sname&"','"&memo&"')"
conn.Execute(sql)
response.Write("<SCRIPT LANGUAGE=JavaScript> alert('添加成功');</SCRIPT>")
end if
end if
%>
<H2><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000066>种类添加</FONT></H2>
<FORM action="sort_add.asp" method=post name="form_administrator" id="form_administrator" onSubmit="return checkform();">
                        
  <TABLE width="550" border=0 cellPadding=5 cellSpacing=1 >
    <TBODY>
      <TR> 
        <TD width="150">种类名称</TD>
        <TD width="71%" bgColor=#ffffff><input name="sname" type="text" id="sname" size="30" maxlength="50">
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
