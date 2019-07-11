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
payment_id=request("payment_id")
if request("option")="add" then
zname=trim(request("name"))
zname=replace(zname,"'","''")
content=trim(request("content"))
content=replace(content,"'","''")
if content="" then content=" "
sql="update payment set name='"&zname&"',content='"&content&"' where payment_id=" & payment_id
conn.Execute(sql)
response.Write("<SCRIPT LANGUAGE=JavaScript> alert('  Edit the payment is complete . ');</SCRIPT>")
end if
%>
<%
sql="select * from payment where payment_id=" & payment_id
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.Open sql,conn,1,1 
%>
<FORM action="payment_e.asp" method=post name="form_administrator" id="form_administrator" onSubmit="return checkform();">
                        
  <TABLE width="511" border=0 align="center" cellPadding=5 
                        cellSpacing=1 bgColor=#CCCCCC>
    <TBODY>
      <TR bgColor=#ffffff> 
        <TD width="29%" align="center" borderColor=#c9c9c9 bgColor=#f5f5f5>payment 
          Name £º</TD>
        <TD width="71%" bgColor=#ffffff><input name="name" type="text" id="name" size="30" maxlength="50" value="<%=rs("name")%>"></TD>
      </TR>
      <TR bgColor=#ffffff> 
        <TD width="29%" align="center" borderColor=#c9c9c9 bgColor=#f5f5f5>Content 
          £º</TD>
        <TD width="71%" bgColor=#ffffff><textarea name="content" cols="50" rows="10" id="content"><%=rs("content")%></textarea></TD>
      </TR>
      <TR bgColor=#ffffff> 
        <TD bgColor=#f5f5f5 borderColor=#c9c9c9>&nbsp;</TD>
        <TD bgColor=#ffffff> 
		<input type="hidden" id="option" name="option" value="add"> 
		<input type="hidden" id="payment_id" name="payment_id" value=<%=payment_id%>> 
        <input type=submit id="options" name="options" value="OK AND SENT &gt;&gt;"> 
        </TD>
      </TR>
    </TBODY>
  </TABLE>
</FORM>
<%
rs.close
set rs=nothing
%>
