<SCRIPT LANGUAGE="JavaScript">
<!--

function checkform()
{
	if(zform.pid.value == 0) {
		alert(' Please select Userid ');
		zform.pid.focus();
		return false;
		}
	if(zform.fid.value == 0) {
		alert(' Please select Field ');
		zform.fid.focus();
		return false;
		}
	if(! isNumber(zform.mmonth.value)) {
		alert('Use integers only for "Match Month" field');
		zform.mmonth.focus();
		return false;
		}
	if (zform.mmonth.value > 12 && zform.mmonth.value < 1) {
		alert('Please enter a valid "Match Month"');
		zform.mmonth.focus();
		return false;
		}
	if(! isNumber(zform.mday.value)) {
		alert('Use integers only for "Match Day" field');
		zform.mday.focus();
		return false;
		}
	if (zform.mday.value > 31 && zform.mday.value < 1) {
		alert('Please enter a valid "Match Day"');
		zform.mday.focus();
		return false;
		}
	if(! isNumber(zform.myear.value)) {
		alert('Use integers only for "Match Year" field');
		zform.myear.focus();
		return false;
		}
	if (zform.myear.value > 2099 && zform.myear.value < 2000) {
		alert('Please enter a valid "Match Year"');
		zform.myear.focus();
		return false;
		}
	if(! isNumber(zform.quantum.value)) {
		alert('Use integers only for "Quantum" field');
		zform.quantum.focus();
		return false;
		}
	return true;
}

function isNumber(name)
{
	if ( name.length==0 )
		return false;
	for(i = 0; i < name.length; i++) {
		if(name.charAt(i) < "0" || name.charAt(i) > "9")
		return false;
	}
	return true;
}
//-->

</SCRIPT>

<!--#include file="adminconn.asp" -->
<%
id=request("id")
if id="" then
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Please select the booking of field you want to edit at first ! ');</SCRIPT>")
response.End()
end if
%>
<%
if request("option")="edit" then

pid=request("pid")
zfield=split(request("fid"),"zooz")
fid=clng(zfield(0))
fname=zfield(1)
mmonth=request("mmonth")
mday=request("mday")
myear=request("myear")
quantum=request("quantum")
zdescription=request("description")
zdescription=replace(zdescription,"'","''")
if zdescription="" then zdescription=" "

sql="update main set pid="&pid&",fid="&fid&",fname='"&fname&"',ddate='"&mmonth&"-"&mday&"-"&myear&"',month="&mmonth&",day="&mday&",year="&myear&",quantum="&quantum&",description='"&zdescription&"' where id=" & id
conn.Execute(sql)
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Edit Field Booking is complete. ');</SCRIPT>")
end if
%>
<%
sql="select * from main where id=" & id
set rs=server.createobject("ADODB.Recordset")
rs.open sql,conn,1,1
%> 
<table width="100%" border="0" cellspacing="2" cellpadding="2">
  <tr>
    <td>&nbsp;</td>
    <td><table width="100%" border="0" cellspacing="2" cellpadding="2">
        <tr> 
          <td height="10">&nbsp;</td>
        </tr>
        <tr> 
          <td height="50" align="center" valign="middle"><strong>Field Booking</strong></td>
        </tr>
      </table></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td align="center"> 
      <form id="zform" action="field_book_e.asp" onSubmit="return checkform();" method="post">
	    <table width="100%" border="0" cellspacing="2" cellpadding="2">
          <tr align="center" valign="middle"> 
            <td height="10" colspan="3"></td>
          </tr>
          <tr> 
            <td width="21%">Match Userid :&nbsp;</td>
            <td width="75%"> <select class="display_dropdown" id="pid" name="pid">
<%
psql="select * from player where pid="&rs("pid")
Set prs=Server.CreateObject("ADODB.RecordSet") 
prs.Open psql,conn,1,1 
if not ( prs.bof or prs.eof ) then
%>
                <option value="<%=prs("pid")%>"><%=prs("uid")%> </option>
<%
end if
prs.close
set prs=nothing
%>
<%
psql="select * from player where uid<>''"
Set prs=Server.CreateObject("ADODB.RecordSet") 
prs.Open psql,conn,1,1 
if not ( prs.bof or prs.eof ) then
do while not prs.eof
%>
                <option value="<%=prs("pid")%>"><%=prs("uid")%> </option>
<%
prs.movenext
loop
end if
prs.close
set prs=nothing
%>
              </select></td>
            <td width="4%">&nbsp;</td>
          </tr>
          <tr> 
            <td width="21%">Match Field :&nbsp;</td>
            <td width="75%"> <select class="display_dropdown" id="fid" name="fid">
                <option selected value="<%=cstr(rs("fid")) & "zooz" & rs("fname")%>"><%=rs("fname")%> </option>
                <%
fsql="select fid,name from field"
Set frs=Server.CreateObject("ADODB.RecordSet") 
frs.Open fsql,conn,1,1 
if not ( frs.bof or frs.eof ) then
do while not frs.eof
%>
                <option value="<%=cstr(frs("fid")) & "zooz" & frs("name")%>"><%=frs("name")%> </option>
                <%
frs.movenext
loop
end if
frs.close
set frs=nothing
%>
              </select></td>
            <td width="4%">&nbsp;</td>
          </tr>
          <tr> 
            <td width="21%">Match Date :&nbsp;</td>
            <td width="75%"><input name="mmonth" id="mmonth" type="text" size="5" maxlength="2" value=<%=rs("month")%>>
              -Month- 
              <input name="mday" id="mday" type="text" size="5" maxlength="2" value=<%=rs("day")%>>
              -Day- 
              <input name="myear" id="myear" type="text" size="5" maxlength="4"  value=<%=rs("year")%>>
              -Year</td>
            <td width="4%">&nbsp;</td>
          </tr>
          <tr> 
            <td width="21%">Quantum :&nbsp;</td>
            <td width="75%"><input name="quantum" id="quantum" type="text" size="5" maxlength="2" value=<%=rs("quantum")%>></td>
            <td width="4%">&nbsp;</td>
          </tr>
          <tr> 
            <td>Description : </td>
            <td><textarea name="description" cols="60" rows="10"><%=rs("description")%></textarea></td>
            <td>&nbsp;</td>
          </tr>
          <tr> 
            <td>&nbsp;</td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
          </tr>
          <tr> 
            <td colspan="3">
        <input type="hidden" id="id" name="id" value=<%=id%>> 
        <input type="submit" id="option" name="option" value="edit"> 
        <input type="submit" id="option" name="option" value="del" onClick="return confirm(' Are you sure to del the tournament ?')"> 
        <input type="button" name="goback" value="go back" onClick="history.go( -1 );return true;"> 
              &nbsp;</td>
          </tr>
        </table>
</form>
</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
</table>
<%
rs.close
set rs=nothing
%>