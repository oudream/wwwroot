<script language="JavaScript">

function checkform()
{
	if (form_administrator.bname.value.length == 0) {
		alert("请输入品牌名称");
		form_administrator.bname.focus();
		return false;
		}
	return true;
}

</script>

<!--#include file="adminconn.asp" -->
<%
if request("option")="add" then
bname=trim(request("bname"))
bname=replace(bname,"'","''")
memo=trim(request("memo"))
memo=replace(memo,"'","''")
if memo="" then memo=" "

usql="select * from brand where bname='"&bname&"'"
set urs=server.createobject("ADODB.Recordset")
urs.open usql,conn,1,1
if not ( urs.eof or urs.bof ) then
response.Write("<SCRIPT LANGUAGE=JavaScript> alert('出错！此品牌有同名存在，请用其它名字输入 ');</SCRIPT>")
urs.close
set urs=nothing
else
urs.close
set urs=nothing

sql="insert into brand (bname,memo) values ('"&bname&"','"&memo&"')"
conn.Execute(sql)
response.Write("<SCRIPT LANGUAGE=JavaScript> alert('添加成功');</SCRIPT>")
end if
end if
%>
<H2><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000066>种类添加</FONT></H2>
<FORM action="brand_add.asp" method=post name="form_administrator" id="form_administrator" onSubmit="return checkform();">
                        
  <TABLE width="550" border=0 cellPadding=5 cellSpacing=1>
    <TBODY>
      <TR> 
        <TD width="150">品牌名称</TD>
        <TD width="71%" bgColor=#ffffff><input name="bname" type="text" id="bname" size="30" maxlength="50"></TD>
      </TR>
      <TR> 
        <TD width="150">备注</TD>
        <TD width="71%" bgColor=#ffffff> <textarea name="memo" id="memo" cols="60" rows="10"></textarea></TD>
      </TR>
      <TR> 
        <TD colspan="2"> <input type="submit" id="option" name="option" value="add"></TD>
      </TR>
    </TBODY>
  </TABLE>
</FORM>
