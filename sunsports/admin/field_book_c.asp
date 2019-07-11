<SCRIPT LANGUAGE="JavaScript">
<!--
function gettarg(targ,selObj,restore){ //v3.0
  eval(targ+".location='field_book_c.asp?fid="+selObj.options[selObj.selectedIndex].value+"'");
  if (restore) selObj.selectedIndex=0;
}

function checkform()
{
	if(zform.username.value == 0) {
		alert(' Please enter a valid "Booked By"');
		zform.username.focus();
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
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="css.css" type=text/css rel=stylesheet>
</head>
<body text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!--#include file="adminconn.asp" -->
<%
fid=request("fid")
if fid="" then fid=0

if request("option")="add" then

fsql="select fid,name from field where fid="&fid
Set frs=Server.CreateObject("ADODB.RecordSet") 
frs.Open fsql,conn,1,1 
fid=frs("fid")
fname=frs("name")
frs.close
set frs=nothing

username=request("username")
username=replace(username,"'","''")
mmonth=request("mmonth")
mday=request("mday")
myear=request("myear")
quantum=request("quantum")
zdescription=request("description")
zdescription=replace(zdescription,"'","''")
if zdescription="" then zdescription=" "


if fid<>"" and quantum<>"" and mmonth<>"" and mday<>"" and myear<>"" then
msql="select * from main where fid="&fid&" and quantum="&quantum&" and month="&mmonth&" and day="&mday&" and year="&myear
set mrs=server.createobject("ADODB.Recordset")
mrs.open msql,conn,1,1
if not ( mrs.eof or mrs.bof ) then
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' The Field have been booked , please book other field. ');</SCRIPT>")
mrs.close
set mrs=nothing
else
mrs.close
set mrs=nothing


sql="insert into main (username,fid,fname,ddate,month,day,year,quantum,description,inserttime) values ('"&username&"',"&fid&",'"&fname&"','"&mmonth&"-"&mday&"-"&myear&"',"&mmonth&","&mday&","&myear&","&quantum&",'"&zdescription&"','"&now()&"')"
conn.Execute(sql)
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Add Field Booking is complete. ');</SCRIPT>")
end if
end if
end if
%>
<table width="100%" border="0" cellspacing="2" cellpadding="2">
  <tr>
    <td>&nbsp;</td>
    <td><table width="100%" border="0" cellspacing="2" cellpadding="2">
        <tr> 
          <td height="10">&nbsp;</td>
        </tr>
        <tr> 
          <td height="50" align="center" valign="middle"><h2><font color="#003366"><strong>Field 
              Booking</strong></font></h2></td>
        </tr>
      </table></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td align="center"> 
      <form id="zform" action="field_book_c.asp" onSubmit="return checkform();" method="post">
	    <table width="100%" border="0" cellspacing="2" cellpadding="2">
          <tr align="center" valign="middle"> 
            <td height="10" colspan="3">&nbsp;</td>
          </tr>
<tr> 
            <td width="21%"> Field :&nbsp;</td>
            <td width="75%"> <select class="display_dropdown" id="fid" name="fid" onChange="gettarg('self',this,0)">
                <option selected value="">SELECT FIELD ... </option>
                <option>--------------------</option>
<%
fsql="select fid,name from field where fid="&fid
Set frs=Server.CreateObject("ADODB.RecordSet") 
frs.Open fsql,conn,1,1 
if not ( frs.bof or frs.eof ) then
%>
                <option value=<%=frs("fid")%> selected><%=frs("name")%></option>
<%
end if
frs.close
set frs=nothing
%>
<%
fsql="select fid,name from field where fid<>"&fid
Set frs=Server.CreateObject("ADODB.RecordSet") 
frs.Open fsql,conn,1,1 
if not ( frs.bof or frs.eof ) then
do while not frs.eof
%>
                <option value=<%=frs("fid")%>><%=frs("name")%> </option>
                <%
frs.movenext
loop
end if
frs.close
set frs=nothing
%>
              </select></td>
            <td width="4%">&nbsp;</td>
          </tr>          <tr> 
            <td width="21%">Booked By :&nbsp;</td>
            <td width="75%"> 
              <input name="username" type="text" id="username" size="30" maxlength="100" class="form"></td>
            <td width="4%">&nbsp;</td>
          </tr>
          
          <tr> 
            <td width="21%"> Date :&nbsp;</td>
            <td width="75%">
          <select name="mmonth" id="mmonth" class="display_dropdown">
            <option value="1" selected>Select The Month ...</option>
            <option>------------------</option>
            <option value="1">1</option>
            <option value="2">2</option>
            <option value="3">3</option>
            <option value="4">4</option>
            <option value="5">5</option>
            <option value="6">6</option>
            <option value="7">7</option>
            <option value="8">8</option>
            <option value="9">9</option>
            <option value="10">10</option>
            <option value="11">11</option>
            <option value="12">12</option>
          </select>
          <select name="mday" id="mday" class="display_dropdown">
            <option value="1" selected>Select The Day ...</option>
            <option>------------------</option>
            <option value="1">1</option>
            <option value="2">2</option>
            <option value="3">3</option>
            <option value="4">4</option>
            <option value="5">5</option>
            <option value="6">6</option>
            <option value="7">7</option>
            <option value="8">8</option>
            <option value="9">9</option>
            <option value="10">10</option>
            <option value="11">11</option>
            <option value="12">12</option>
            <option value="13">13</option>
            <option value="14">14</option>
            <option value="15">15</option>
            <option value="16">16</option>
            <option value="17">17</option>
            <option value="18">18</option>
            <option value="19">19</option>
            <option value="20">20</option>
            <option value="21">21</option>
            <option value="22">22</option>
            <option value="23">23</option>
            <option value="24">24</option>
            <option value="25">25</option>
            <option value="26">26</option>
            <option value="27">27</option>
            <option value="28">28</option>
            <option value="29">29</option>
            <option value="30">30</option>
            <option value="31">31</option>
          </select>
          <select name="myear" id="myear" class="display_dropdown">
            <option value="2004" selected>Select The Year ...</option>
            <option>------------------</option>
            <option value="2001">2001</option>
            <option value="2002">2002</option>
            <option value="2003">2003</option>
            <option value="2004" selected>2004</option>
            <option value="2005">2005</option>
            <option value="2006">2006</option>
            <option value="2007">2007</option>
            <option value="2008">2008</option>
            <option value="2009">2009</option>
            <option value="2010">2010</option>
          </select></td>
            <td width="4%">&nbsp;</td>
          </tr>
          <tr> 
            <td width="21%">Quantum :&nbsp;</td>
            <td width="75%">
			<select id="quantum" name="quantum" class="display_dropdown">
<%
if fid=5 then
%>			
			<option value=1>8:00 -- 10:00</option>
			<option value=2>10:00 -- 12:00</option>
			<option value=3>12:00 -- 14:00</option>
			<option value=4>14:00 -- 16:00</option>
			<option value=5>16:00 -- 18:00</option>
			<option value=6>19:00 -- 21:00</option>
			<option value=7>21:00 -- 23:00</option>
<%
else
%>			
			<option value=1>8:00 -- 10:00</option>
			<option value=2>10:00 -- 12:00</option>
			<option value=3>12:00 -- 14:00</option>
			<option value=4>14:00 -- 16:00</option>
			<option value=5>16:00 -- 18:00</option>
<%
end if
%>
			</select>
			</td>
            <td width="4%">&nbsp;</td>
          </tr>
          <tr> 
            <td>Description : </td>
            <td><textarea name="description" cols="60" rows="10"></textarea></td>
            <td>&nbsp;</td>
          </tr>
          <tr> 
            <td>&nbsp;</td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
          </tr>
          <tr> 
            <td colspan="3">
			  <input type="submit" id="option" name="option" value="add"> 
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
