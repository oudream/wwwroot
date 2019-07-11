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
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="css.css" type=text/css rel=stylesheet>
</head>
<body text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">

<!--#include file="adminconn.asp" -->
<%
wear_num=request("wear_num")

if request("option")="edit" then

select case wear_num
case 11 sort_id_0=1110
		sort_id_1=1111
		sort_id_2=1112
		sort_id_3=1113
case 22 sort_id_0=2210
		sort_id_1=2211
		sort_id_2=2212
		sort_id_3=2213
case 33 sort_id_0=3310
		sort_id_1=3311
		sort_id_2=3312
		sort_id_3=3313
end select

wear_right=trim(request("wear_right"))
wear_one=trim(request("wear_one"))
wear_two=trim(request("wear_two"))
wear_tree=trim(request("wear_tree"))

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
rs.Open sql,conn,1,1 
if rs.bof or rs.eof then response.Redirect("error.asp")
do while not rs.eof
wear_array(i)=rs("pic")
i=i+1
rs.movenext
loop
rs.close
set rs=nothing
%>

<FORM action="wear_page_e.asp" method=post name="form_administrator" id="form_administrator" onSubmit="return checkform();">
  <br>
  <br>
  <table width="569" border="1" align="center" cellpadding="2" cellspacing="0" bordercolor="CCCCCC">
    <tr align="center"> 
      <td height="40" colspan="2"><font color="#FF0000" size="5"><strong> 
        <%
if wear_num=11 then response.Write("MEN PAGE EDIT WIZARD")
if wear_num=22 then response.Write("WOMEN PAGE EDIT WIZARD")
if wear_num=33 then response.Write("CHILDREN PAGE EDIT WIZARD")
%>
        </strong></font></td>
    </tr>
    <tr> 
      <td><strong>top right ( picture ) :</strong></td>
      <td><input name="wear_right" type="text" id="wear_right" size="30" maxlength="50" value="<%=wear_array(0)%>"> 
        <font color="#FF0000"> * </font></td>
    </tr>
    <tr> 
      <td> <strong>bottom one ( picture ) :</strong></td>
      <td><input name="wear_one" type="text" id="wear_one2" size="30" maxlength="50" value="<%=wear_array(1)%>"> 
        <font color="#FF0000"> * </font></td>
    </tr>
    <tr> 
      <td> <strong>bottom two ( picture ) :</strong></td>
      <td><input name="wear_two" type="text" id="wear_two2" size="30" maxlength="50" value=<%=wear_array(2)%>> 
        <font color="#FF0000"> * </font></td>
    </tr>
    <tr> 
      <td><strong>bottom tree ( picture ) :</strong></td>
      <td><input name="wear_tree" type="text" id="wear_tree2" size="30" maxlength="50" value="<%=wear_array(3)%>"> 
        <font color="#FF0000"> * </font></td>
    </tr>
    <tr align="center"> 
      <td colspan="2"><input type="hidden" id="wear_num3" name="wear_num" value="<%=wear_num%>"> 
        <input type="hidden" id="option3" name="option" value="edit"> <input type=submit id="Submit3" name="Submit" value="OK AND SENT &gt;&gt;"></td>
    </tr>
    <tr>
      <td colspan="2"><strong>Notice : the picture file must put in folder "<font color="#FF0000"> 
        <%
select case wear_num
case 11 response.Write("MEN")
case 22 response.Write("WOMEN")
case 33 response.Write("CHILDREN")
end select
%>
        </font>"</strong></td>
    </tr>
  </table>
  
</FORM>
</body>
</html>
