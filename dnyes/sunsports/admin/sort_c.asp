<script language="JavaScript">

function checkform()
{
	if (form_administrator.name.value.length == 0) {
		alert("Please enter a valid Datetime. ");
		form_administrator.name.focus();
		return false;
		}
	return true;
}

</script>

<!--#include file="adminconn.asp" -->
<%
if request("option")="add" then
zname=trim(request("name"))
zname=replace(zname,"'","''")
sql="insert into sort (name) values ('"&zname&"')"
conn.Execute(sql)
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' OK . ');</SCRIPT>")
end if
%>
<FORM action="sort_c.asp" method=post name="form_administrator" id="form_administrator" onSubmit="return checkform();">
                        
  <TABLE width="511" border=0 align="center" cellPadding=5 
                        cellSpacing=1 bgColor=#CCCCCC>
    <TBODY>
      <TR bgColor=#ffffff> 
        <TD width="29%" align="center" borderColor=#c9c9c9 bgColor=#f5f5f5>Sort 
          Name £º</TD>
        <TD width="71%" bgColor=#ffffff><input name="name" type="text" id="name" size="30" maxlength="50"></TD>
      </TR>
      <TR bgColor=#ffffff> 
        <TD bgColor=#f5f5f5 borderColor=#c9c9c9>&nbsp;</TD>
        <TD bgColor=#ffffff> <input type="hidden" id="option" name="option" value="add"> 
          <input type=submit id="options" name="options" value="OK AND SENT &gt;&gt;"> 
        </TD>
      </TR>
    </TBODY>
  </TABLE>
</FORM>
