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
if request("option")="edit" then

sort_id=request("sort_id")

zname=trim(request("name"))
zname=replace(zname,"'","''")

code=trim(request("code"))
code=replace(code,"'","''")
if code ="" then code=" "

price=trim(request("price"))
if price ="" then price=0

zsize=trim(request("size"))
zsize=replace(zsize,"'","''")
if zsize ="" then zsize=" "

putdate=trim(request("putdate"))
putdate=replace(putdate,"'","''")
if putdate ="" then putdate=" "

packaging=trim(request("packaging"))
packaging=replace(packaging,"'","''")
if packaging ="" then packaging=" "

madein=trim(request("madein"))
madein=replace(madein,"'","''")
if madein ="" then madein=" "

brand=trim(request("brand"))
brand=replace(brand,"'","''")
if brand ="" then brand=" "

pic=trim(request("pic"))
pic=replace(pic,"'","''")
if pic ="" then pic=" "

photo=trim(request("photo"))
photo=replace(photo,"'","''")
if photo ="" then photo=" "

content=trim(request("content"))
content=replace(content,"'","''")
if content ="" then content=" "

sql="update subs set sort_id="&sort_id&",name='"&zname&"',code='"&code&"',price="&price&",size='"&zsize&"',putdate='"&putdate&"',packaging='"&packaging&"',madein='"&madein&"',brand='"&brand&"',pic='"&pic&"',photo='"&photo&"',content='"&content&"' where subs_id=" & request("subs_id")
conn.Execute(sql)
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Edit the subs is complete .');</SCRIPT>")
end if
%>

<%
ssql="select * from subs where subs_id="&request("subs_id") 
Set srs=Server.CreateObject("ADODB.RecordSet") 
srs.Open ssql,conn,1,1 
if srs.bof or srs.eof then response.Redirect("error.asp?err=1")
%>

<FORM action="subs_e.asp" method=post name="form_administrator" id="form_administrator" onSubmit="return checkform();">
                        
  <TABLE width="569" border=0 align="center" cellPadding=5 
                        cellSpacing=1 bgColor=#CCCCCC>
    <TBODY>
      <TR bgColor=#ffffff> 
        <TD width="34%" align="center" borderColor=#c9c9c9 bgColor=#f5f5f5>Sort 
          Name £º</TD>
        <TD width="66%" bgColor=#ffffff> 
		<select name="sort_id">
<%
sql="select * from sort where sort_id="&srs("sort_id") 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.Open sql,conn,1,1 
if rs.bof or rs.eof then response.Redirect("error.asp?err=1")
do while not rs.eof
%>
		<option value=<%=rs("sort_id")%> selected><%=rs("name")%></option>
<%
rs.movenext
loop
rs.close
set rs=nothing
%>
<%
sql="select * from sort where sort_id<>"&srs("sort_id") 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.Open sql,conn,1,1 
if rs.bof or rs.eof then response.Redirect("error.asp?err=1")
do while not rs.eof
%>
		<option value=<%=rs("sort_id")%> selected><%=rs("name")%></option>
<%
rs.movenext
loop
rs.close
set rs=nothing
%>
          </select>
          <font color="#FF0000"> * </font></TD>
      </TR>
      <TR bgColor=#ffffff> 
        <TD width="34%" align="center" borderColor=#c9c9c9 bgColor=#f5f5f5>subs 
          Name £º</TD>
        <TD width="66%" bgColor=#ffffff><input name="name" type="text" id="name" size="30" maxlength="50" value="<%=srs("name")%>">
          <font color="#FF0000"> * </font></TD>
      </TR>
      <TR bgColor=#ffffff> 
        <TD width="34%" align="center" borderColor=#c9c9c9 bgColor=#f5f5f5>subs 
          Code £º</TD>
        <TD width="66%" bgColor=#ffffff><input name="code" type="text" id="code" size="30" maxlength="50" value="<%=srs("code")%>">
        </TD>
      </TR>
      <TR bgColor=#ffffff> 
        <TD width="34%" align="center" borderColor=#c9c9c9 bgColor=#f5f5f5>subs 
          Price £º</TD>
        <TD width="66%" bgColor=#ffffff><input name="price" type="text" id="price" size="30" maxlength="50" value=<%=srs("price")%>></TD>
      </TR>
      <TR bgColor=#ffffff> 
        <TD width="34%" align="center" borderColor=#c9c9c9 bgColor=#f5f5f5>subs 
          Size £º</TD>
        <TD width="66%" bgColor=#ffffff><input name="size" type="text" id="size" size="30" maxlength="50" value="<%=srs("size")%>"></TD>
      </TR>
      <TR bgColor=#ffffff> 
        <TD width="34%" align="center" borderColor=#c9c9c9 bgColor=#f5f5f5>subs 
          Putdate £º</TD>
        <TD width="66%" bgColor=#ffffff><input name="putdate" type="text" id="putdate" size="30" maxlength="50" value="<%=srs("putdate")%>"></TD>
      </TR>
      <TR bgColor=#ffffff> 
        <TD width="34%" align="center" borderColor=#c9c9c9 bgColor=#f5f5f5>subs 
          Pack £º</TD>
        <TD width="66%" bgColor=#ffffff><input name="packaging" type="text" id="packaging" size="30" maxlength="50" value="<%=srs("packaging")%>"></TD>
      </TR>
      <TR bgColor=#ffffff> 
        <TD width="34%" align="center" borderColor=#c9c9c9 bgColor=#f5f5f5>subs 
          Made From£º</TD>
        <TD width="66%" bgColor=#ffffff><input name="madein" type="text" id="madein" size="30" maxlength="50" value="<%=srs("madein")%>"></TD>
      </TR>
      <TR bgColor=#ffffff> 
        <TD width="34%" align="center" borderColor=#c9c9c9 bgColor=#f5f5f5>subs 
          Brand £º</TD>
        <TD width="66%" bgColor=#ffffff><input name="brand" type="text" id="brand" size="30" maxlength="50" value="<%=srs("brand")%>"></TD>
      </TR>
      <TR bgColor=#ffffff> 
        <TD width="34%" align="center" borderColor=#c9c9c9 bgColor=#f5f5f5>subs 
          Pic £º</TD>
        <TD width="66%" bgColor=#ffffff><input name="pic" type="text" id="pic" size="30" maxlength="50" value="<%=srs("pic")%>"></TD>
      </TR>
      <TR bgColor=#ffffff> 
        <TD width="34%" align="center" borderColor=#c9c9c9 bgColor=#f5f5f5>subs 
          Photo £º</TD>
        <TD width="66%" bgColor=#ffffff><input name="photo" type="text" id="photo" size="30" maxlength="50" value="<%=srs("photo")%>"></TD>
      </TR>
      <TR bgColor=#ffffff> 
        <TD width="34%" align="center" borderColor=#c9c9c9 bgColor=#f5f5f5>subs 
          Content £º</TD>
        <TD width="66%" bgColor=#ffffff><textarea name="content" cols="50" rows="10" id="content"> <%=srs("content")%> </textarea></TD>
      </TR>
      <TR bgColor=#ffffff> 
        <TD bgColor=#f5f5f5 borderColor=#c9c9c9>&nbsp;</TD>
        <TD bgColor=#ffffff> 
		<input type="hidden" id="subs_id" name="subs_id" value=<%=request("subs_id")%>> 
		<input type="hidden" id="option" name="option" value="edit"> 
        <input type=submit id="options" name="options" value="OK AND SENT &gt;&gt;"> 
        </TD>
      </TR>
    </TBODY>
  </TABLE>
</FORM>
<%
srs.close
set srs=nothing
%>
