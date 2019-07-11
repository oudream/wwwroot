<script language="JavaScript">

function checkform()
{
	if (form_administrator.name.value.length == 0) {
		alert("Please enter a valid Datetime. ");
		form_administrator.name.focus();
		return false;
		}
	if(! isLetter(form_administrator.price.value)) {
		alert('The "Price" field makes up by "_" and "0~9" and "a~z" ');
		form_administrator.price.focus();
		return false;
		}
	return true;
}

function isLetter(name)
{
var letter=" _.0123456789"
	for(i = 0; i < name.length; i++) {
		if(letter.indexOf(name.charAt(i)) == -1)
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

price=request("price")
if price="" then price=0

content=trim(request("content"))
content=replace(content,"'","''")
if content="" then content=" "
sql="insert into payout (name,price,content) values ('"&zname&"',"&price&",'"&content&"')"
conn.Execute(sql)
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Your Content is send , it will be show at homepage . ');</SCRIPT>")
end if
%>
<FORM action="payout_c.asp" method=post name="form_administrator" id="form_administrator" onSubmit="return checkform();">
                        
  <TABLE width="511" border=0 align="center" cellPadding=5 
                        cellSpacing=1 bgColor=#CCCCCC>
    <TBODY>
      <TR bgColor=#ffffff> 
        <TD width="29%" align="center" borderColor=#c9c9c9 bgColor=#f5f5f5>payout 
          Name £º</TD>
        <TD width="71%" bgColor=#ffffff><input name="name" type="text" id="name" size="30" maxlength="50"></TD>
      </TR>
      <TR bgColor=#ffffff> 
        <TD width="29%" align="center" borderColor=#c9c9c9 bgColor=#f5f5f5>payout 
          Price £º</TD>
        <TD width="71%" bgColor=#ffffff><input name="price" type="text" id="price" size="10" maxlength="10">
          ( Default 0 )</TD>
      </TR>
      <TR bgColor=#ffffff> 
        <TD width="29%" align="center" borderColor=#c9c9c9 bgColor=#f5f5f5>Content 
          £º</TD>
        <TD width="71%" bgColor=#ffffff><textarea name="content" cols="50" rows="10" id="content"></textarea></TD>
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
