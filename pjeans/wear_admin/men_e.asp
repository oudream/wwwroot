<script language="JavaScript">

function checkform()
{
	if (form_administrator.wear_right.value.length == 0) {
		alert("Please Select a valid -right picture- Name. ");
		form_administrator.wear_right.focus();
		return false;
		}
	if(! isEnglish(form_administrator.wear_right.value)) {
		alert(' Sorry ! the -right picture- Name must be made up by ENGLISH');
		form_administrator.wear_right.focus();
		return false;
		}
	if (form_administrator.wear_one.value.length == 0) {
		alert("Please Select a valid -one picture- Name. ");
		form_administrator.wear_one.focus();
		return false;
		}
	if(! isEnglish(form_administrator.wear_one.value)) {
		alert(' Sorry ! the -one picture- Name must be made up by ENGLISH');
		form_administrator.wear_one.focus();
		return false;
		}
	if (form_administrator.wear_two.value.length == 0) {
		alert("Please Select a valid -two picture- Name. ");
		form_administrator.wear_two.focus();
		return false;
		}
	if(! isEnglish(form_administrator.wear_two.value)) {
		alert(' Sorry ! the -two picture- Name must be made up by ENGLISH');
		form_administrator.wear_two.focus();
		return false;
		}
	if (form_administrator.wear_tree.value.length == 0) {
		alert("Please Select a valid -tree picture- Name. ");
		form_administrator.wear_tree.focus();
		return false;
		}
	if(! isEnglish(form_administrator.wear_tree.value)) {
		alert(' Sorry ! the -tree picture- Name must be made up by ENGLISH');
		form_administrator.wear_tree.focus();
		return false;
		}
	return true;
}

function isEnglish(name)
{
	for(i = 0; i < name.length; i++) {
		if(name.charCodeAt(i) > 128)
			return false;
	}
	return true;
}


</script>

<!--#include file="adminconn.asp" -->
<%
wear_num=request("wear_num")

if request("option")="edit" then

select case wear_num
case 11 sort_id_0=1110
		sort_id_1=1111
		sort_id_1=1112
		sort_id_1=1113
case 22 sort_id_0=2210
		sort_id_1=2211
		sort_id_1=2212
		sort_id_1=2213
case 33 sort_id_0=3310
		sort_id_1=3311
		sort_id_1=3312
		sort_id_1=3313
end select

wear_right=trim(request("wear_right")
wear_one=trim(request("wear_one")
wear_two=trim(request("wear_two")
wear_tree=trim(request("wear_tree")

conn.Execute("update subs set pic='"&wear_right&"' where sort_id=" & sort_id_0)
conn.Execute("update subs set pic='"&wear_one&"' where sort_id=" & sort_id_1)
conn.Execute("update subs set pic='"&wear_two&"' where sort_id=" & sort_id_2)
conn.Execute("update subs set pic='"&wear_tree&"' where sort_id=" & sort_id_3)
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Edit the subs is complete .');</SCRIPT>")
end if
%>

<%
dim wear_array(3)
i=0
if wear_num=11 then sql="select * from subs where sort_id=1110 or sort_id=1111 or sort_id=1112 or sort_id=1113 order by sort_id asc" 
if wear_num=22 then sql="select * from subs where sort_id=2210 or sort_id=2211 or sort_id=2212 or sort_id=2213 order by sort_id asc" 
if wear_num=33 then sql="select * from subs where sort_id=3310 or sort_id=3311 or sort_id=3312 or sort_id=3313 order by sort_id asc" 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.Open ssql,conn,1,1 
if rs.bof or rs.eof then response.Redirect("error.asp")
do while not rs.eof
wear_array(i)=rs("pic")
i=i+1
rs.movenext
loop
rs.close
set rs=nothing
%>

<FORM action="subs_e.asp" method=post name="form_administrator" id="form_administrator" onSubmit="return checkform();">
                        
  <TABLE width="569" border=0 align="center" cellPadding=5 
                        cellSpacing=1 bgColor=#CCCCCC>
    <TBODY>
      <TR align="center" bordercolor="#FFFFFF" bgColor=#FFFFFF> 
        <TD colspan="2"><font color="#FF0000" size="5"><strong>
<%
if wear_num=11 then response.Write("MEN PAGE EDIT WIZARD")
if wear_num=22 then response.Write("WOMEN PAGE EDIT WIZARD")
if wear_num=33 then response.Write("CHILDREN PAGE EDIT WIZARD")
%>
          </strong> </font> </TD>
      </TR>
      <TR> 
        <TD width="34%" align="left" borderColor=#c9c9c9 bgColor=#f5f5f5><strong>top 
          right ( picture ) :</strong></TD>
        <TD width="66%" bgColor=#ffffff><input name="wear_right" type="text" id="wear_right" size="30" maxlength="50" value="<%=wear_array(0)%>"> 
          <font color="#FF0000"> * </font></TD>
      </TR>
      <TR> 
        <TD width="34%" align="left" borderColor=#c9c9c9 bgColor=#f5f5f5> <strong>bottom 
          one ( picture ) :</strong></TD>
        <TD width="66%" bgColor=#ffffff><input name="wear_one" type="text" id="wear_one" size="30" maxlength="50" value="<%=wear_array(1)%>"> 
          <font color="#FF0000"> * </font> </TD>
      </TR>
      <TR> 
        <TD width="34%" align="left" borderColor=#c9c9c9 bgColor=#f5f5f5> <strong>bottom 
          two ( picture ) :</strong></TD>
        <TD width="66%" bgColor=#ffffff><input name="wear_two" type="text" id="wear_two" size="30" maxlength="50" value=<%=wear_array(2)%>> 
          <font color="#FF0000"> * </font> </TD>
      </TR>
      <TR> 
        <TD width="34%" align="left" borderColor=#c9c9c9 bgColor=#f5f5f5><strong>bottom 
          tree ( picture ) :</strong></TD>
        <TD width="66%" bgColor=#ffffff><input name="wear_tree" type="text" id="wear_tree" size="30" maxlength="50" value="<%=wear_array(3)%>"> 
          <font color="#FF0000"> * </font> </TD>
      </TR>
      <TR> 
        <TD height="50" align="left" borderColor=#c9c9c9 bgColor=#f5f5f5>&nbsp;</TD>
        <TD bgColor=#ffffff> <input type="hidden" id="wear_num" name="wear_num" value="<%=wear_num%>"> 
          <input type=submit id="options" name="options" value="OK AND SENT &gt;&gt;"> 
        </TD>
      </TR>
    </TBODY>
  </TABLE>
</FORM>
