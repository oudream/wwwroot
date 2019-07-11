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
if request("option")="add" then

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

sql="insert into subs (sort_id,name,code,price,size,putdate,packaging,madein,brand,pic,photo,content,inserttime) values ("&sort_id&",'"&zname&"','"&code&"',"&price&",'"&zsize&"','"&putdate&"','"&packaging&"','"&madein&"','"&brand&"','"&pic&"','"&photo&"','"&content&"','"&now()&"')"
conn.Execute(sql)
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Your Content is send , it will be show at homepage . ');</SCRIPT>")
end if
%>
<FORM action="subs_c.asp" method=post name="form_administrator" id="form_administrator" onSubmit="return checkform();">
                        
  <TABLE width="569" border=0 align="center" cellPadding=5 
                        cellSpacing=1 bgColor=#CCCCCC>
    <TBODY>
      <TR bgColor=#ffffff> 
        <TD width="34%" align="center" borderColor=#c9c9c9 bgColor=#f5f5f5>Sort 
          Name £º</TD>
        <TD width="66%" bgColor=#ffffff> 
		<select name="sort_id">
		<option value="" selected>Select ...</option>
<%
sql="select * from sort order by sort_id desc" 
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
        <TD width="66%" bgColor=#ffffff><input name="name" type="text" id="name" size="30" maxlength="50">
          <font color="#FF0000"> * </font></TD>
      </TR>
      <TR bgColor=#ffffff> 
        <TD width="34%" align="center" borderColor=#c9c9c9 bgColor=#f5f5f5>subs 
          Code £º</TD>
        <TD width="66%" bgColor=#ffffff><input name="code" type="text" id="code" size="30" maxlength="50">
        </TD>
      </TR>
      <TR bgColor=#ffffff> 
        <TD width="34%" align="center" borderColor=#c9c9c9 bgColor=#f5f5f5>subs 
          Price £º</TD>
        <TD width="66%" bgColor=#ffffff><input name="price" type="text" id="price" size="30" maxlength="50"></TD>
      </TR>
      <TR bgColor=#ffffff> 
        <TD width="34%" align="center" borderColor=#c9c9c9 bgColor=#f5f5f5>subs 
          Size £º</TD>
        <TD width="66%" bgColor=#ffffff><input name="size" type="text" id="size" size="30" maxlength="50"></TD>
      </TR>
      <TR bgColor=#ffffff> 
        <TD width="34%" align="center" borderColor=#c9c9c9 bgColor=#f5f5f5>subs 
          Putdate £º</TD>
        <TD width="66%" bgColor=#ffffff><input name="putdate" type="text" id="putdate" size="30" maxlength="50"></TD>
      </TR>
      <TR bgColor=#ffffff> 
        <TD width="34%" align="center" borderColor=#c9c9c9 bgColor=#f5f5f5>subs 
          Pack £º</TD>
        <TD width="66%" bgColor=#ffffff><input name="packaging" type="text" id="packaging" size="30" maxlength="50"></TD>
      </TR>
      <TR bgColor=#ffffff> 
        <TD width="34%" align="center" borderColor=#c9c9c9 bgColor=#f5f5f5>subs 
          Made From£º</TD>
        <TD width="66%" bgColor=#ffffff><input name="madein" type="text" id="madein" size="30" maxlength="50"></TD>
      </TR>
      <TR bgColor=#ffffff> 
        <TD width="34%" align="center" borderColor=#c9c9c9 bgColor=#f5f5f5>subs 
          Brand £º</TD>
        <TD width="66%" bgColor=#ffffff><input name="brand" type="text" id="brand" size="30" maxlength="50"></TD>
      </TR>
      <TR bgColor=#ffffff> 
        <TD width="34%" align="center" borderColor=#c9c9c9 bgColor=#f5f5f5>subs 
          Pic £º</TD>
        <TD width="66%" bgColor=#ffffff><input name="pic" type="text" id="pic" size="30" maxlength="50" value="pic/"></TD>
      </TR>
      <TR bgColor=#ffffff> 
        <TD width="34%" align="center" borderColor=#c9c9c9 bgColor=#f5f5f5>subs 
          Photo £º</TD>
        <TD width="66%" bgColor=#ffffff><input name="photo" type="text" id="photo" size="30" maxlength="50" value="photo/"></TD>
      </TR>
      <TR bgColor=#ffffff> 
        <TD width="34%" align="center" borderColor=#c9c9c9 bgColor=#f5f5f5>subs 
          Content £º</TD>
        <TD width="66%" bgColor=#ffffff><textarea name="content" cols="50" rows="10" id="content"></textarea></TD>
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
