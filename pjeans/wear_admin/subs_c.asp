<script language="JavaScript">

String.prototype.Trim  = function(){return this.replace(/^\s+|\s+$/g,"");}
String.prototype.Ltrim = function(){return this.replace(/^\s+/g, "");}
String.prototype.Rtrim = function(){return this.replace(/\s+$/g, "");}

function checkform()
{
	if (form_administrator.name.value.length == 0) {
		alert("Please enter a valid Datetime. ");
		form_administrator.name.focus();
		return false;
		}
	if (form_administrator.pic.value.Trim() == form_administrator.photo.value.Trim()) {
		alert(" Two file name have not to be equal . ");
		form_administrator.pic.value = '';
		form_administrator.photo.value = '';
		form_administrator.pic.focus();
		return false;
		}
	return true;
}

</script>

<!--#include file="adminconn.asp" -->
<%
sort_id=request("sort_id")

if request("option")="add" then

zname=trim(request("name"))
zname=replace(zname,"'","''")

pic=trim(request("pic"))
pic=replace(pic,"'","''")
if pic ="" then pic=" "

photo=trim(request("photo"))
photo=replace(photo,"'","''")
if photo ="" then photo=" "

sql="insert into subs (sort_id,name,pic,photo,inserttime) values ("&sort_id&",'"&zname&"','"&pic&"','"&photo&"','"&now()&"')"
conn.Execute(sql)
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Your Content is send , it will be show at homepage . ');</SCRIPT>")
end if
%>
<FORM action="subs_c.asp" method=post name="form_administrator" id="form_administrator" onSubmit="return checkform();">
 
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td height="70" align="center" valign="bottom"><h2><font color="#000066">
        <%
if sort_id<>"" then
Select Case sort_id
           Case 11    response.Write("ADD MEN¡¯S PRODUCTS")
           Case 22    response.Write("ADD WOMEN¡¯S PRODUCTS")
           Case 33    response.Write("ADD CHILDREN¡¯S PRODUCTS")
End Select
end if
		%>
        </font></h2></td>
    </tr>
  </table>
  <TABLE width="569" border=1 align="center" cellPadding=2 
                        cellSpacing=0 bordercolor="#CCCCCC">
    <TBODY>
      <TR> 
        <TD width="34%" align="center">subs Name £º</TD>
        <TD width="66%" bgColor=#ffffff><input name="name" type="text" id="name" size="30" maxlength="50"> 
          <font color="#FF0000"> * </font></TD>
      </TR>
      <TR> 
        <TD width="34%" align="center">subs Pic £º</TD>
        <TD width="66%" bgColor=#ffffff><input name="pic" type="text" id="pic" size="30" maxlength="50"> 
          <font color="#FF0000"> * </font></TD>
      </TR>
      <TR> 
        <TD width="34%" align="center">subs Photo £º</TD>
        <TD width="66%" bgColor=#ffffff><input name="photo" type="text" id="photo" size="30" maxlength="50"> 
          <font color="#FF0000"> *</font></TD>
      </TR>
      <TR align="center"> 
        <TD colspan="2"> <input type="hidden" id="sort_id" name="sort_id" value=<%=sort_id%>> 
          <input type="hidden" id="option" name="option" value="add"> <input type=submit id="options" name="options" value="OK AND SENT &gt;&gt;"> 
        </TD>
      </TR>
      <TR> 
        <TD height="100" colspan="2" align="left"><strong>Notice : the picture 
          file must put in folder "<font color="#FF0000"> 
          <%
select case sort_id
case 11 response.Write("MEN")
case 22 response.Write("WOMEN")
case 33 response.Write("CHILDREN")
end select
%>
          </font>"</strong></TD>
      </TR>
    </TBODY>
  </TABLE>
</FORM>
