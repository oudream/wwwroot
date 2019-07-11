<script language="JavaScript">

function checkform()
{
	if (form_administrator.sort_id.value.length == 0) {
		alert("Please Select a valid Sort Name. ");
		form_administrator.sort_id.focus();
		return false;
		}
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
subs_id=request("subs_id")

if request("option")="edit" then

zname=trim(request("name"))
zname=replace(zname,"'","''")

pic=trim(request("pic"))
pic=replace(pic,"'","''")
if pic ="" then pic=" "

photo=trim(request("photo"))
photo=replace(photo,"'","''")
if photo ="" then photo=" "

sql="update subs set pic='"&pic&"',name='"&zname&"',photo='"&photo&"' where subs_id=" & subs_id
conn.Execute(sql)
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Edit the subs is complete .');</SCRIPT>")
end if
%>

<%
ssql="select * from subs where subs_id="&request("subs_id") 
Set srs=Server.CreateObject("ADODB.RecordSet") 
srs.Open ssql,conn,1,1 
if srs.bof or srs.eof then response.Redirect("error.asp")
%>

<FORM action="subs_e.asp" method=post name="form_administrator" id="form_administrator" onSubmit="return checkform();">
                        
  <TABLE width="569" border=0 align="center" cellPadding=5 
                        cellSpacing=1 bgColor=#CCCCCC>
    <TBODY>
      <TR bgColor=#ffffff> 
        <TD colspan="2" align="center" borderColor=#c9c9c9 bgColor=#f5f5f5><strong><font size="5">subs 
          edit wizard</font></strong></TD>
      </TR>
      <TR bgColor=#ffffff> 
        <TD width="34%" align="center" borderColor=#c9c9c9 bgColor=#f5f5f5><strong>subs 
          Name :</strong></TD>
        <TD width="66%" bgColor=#ffffff><input name="name" type="text" id="name" size="30" maxlength="50" value="<%=srs("name")%>"> 
          <font color="#FF0000"> * </font></TD>
      </TR>
      <TR bgColor=#ffffff> 
        <TD width="34%" align="center" borderColor=#c9c9c9 bgColor=#f5f5f5><strong>subs 
          Code :</strong></TD>
        <TD width="66%" bgColor=#ffffff><input name="pic" type="text" id="pic" size="30" maxlength="50" value="<%=srs("pic")%>">
          <font color="#FF0000"> * </font></TD>
      </TR>
      <TR bgColor=#ffffff> 
        <TD width="34%" align="center" borderColor=#c9c9c9 bgColor=#f5f5f5><strong>subs 
          Price :</strong></TD>
        <TD width="66%" bgColor=#ffffff><input name="photo" type="text" id="photo" size="30" maxlength="50" value=<%=srs("photo")%>>
          <font color="#FF0000">*</font></TD>
      </TR>
      <TR bgColor=#ffffff> 
        <TD bgColor=#f5f5f5 borderColor=#c9c9c9>&nbsp;</TD>
        <TD bgColor=#ffffff> <input type="hidden" id="subs_id" name="subs_id" value=<%=request("subs_id")%>> 
          <input type="hidden" id="option" name="option" value="edit"> <input type=submit id="options" name="options" value="OK AND SENT &gt;&gt;"> 
        </TD>
      </TR>
    </TBODY>
  </TABLE>
</FORM>
<%
srs.close
set srs=nothing
%>
