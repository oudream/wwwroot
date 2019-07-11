<SCRIPT LANGUAGE="JavaScript">
<!--

function checkform()
{
	if(zform.fid.value == 0) {
		alert(' Please select Field ');
		zform.fid.focus();
		return false;
		}
	if(! isNumber(zform.myear.value)) {
		alert('Use integers only for "Match Year" field');
		zform.myear.focus();
		return false;
		}
	if (zform.myear.value > 2099 || zform.myear.value < 2000) {
		alert(' The Year must in : 2000 --2099 ' );
		zform.myear.focus();
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
if trim(request("option"))="del" then
conn.execute "delete * from main where id=" & request("id")
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Del field booking is complete. ');</SCRIPT>")
end if
%>
<table width="100%" border="0" cellspacing="2" cellpadding="2">
  <tr>
    <td>&nbsp;</td>
    <td height="45" align="center"><h2><font color="#000066">Field booking Edit/Del</font></h2></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td> 
      <form id="zform" action="field_book_v.asp" onSubmit="return checkform();" method="post">
        <table width="100%" border="0" cellspacing="2" cellpadding="2">
          <tr align="center" valign="middle"> 
            <td height="10" colspan="5"></td>
          </tr>
          <tr> 
            <td width="12%">Match Field :&nbsp;</td>
            <td width="17%"> <select class="display_dropdown" id="fid" name="fid">
                <option selected value="">SELECT ... </option>
                <%
fsql="select fid,name from field"
Set frs=Server.CreateObject("ADODB.RecordSet") 
frs.Open fsql,conn,1,1 
if not ( frs.bof or frs.eof ) then
do while not frs.eof
%>
                <option value="<%=frs("fid")%>"><%=frs("name")%> </option>
                <%
frs.movenext
loop
end if
frs.close
set frs=nothing
%>
              </select></td>
            <td width="13%">&nbsp;</td>
            <td width="41%">&nbsp;</td>
            <td width="17%">&nbsp;</td>
          </tr>
          <tr> 
            <td colspan="5">Match Date :&nbsp; <input name="mmonth" id="mmonth" type="text" size="5" maxlength="2" class="form">
              -Month- 
              <input name="mday" id="mday" type="text" size="5" maxlength="2" class="form">
              -Day- 
              <input name="myear" id="myear" type="text" size="5" maxlength="4" value="<%=year(date())%>" class="form">
              -Year
              <input type="submit" name="Submit" value="Display"></td>
          </tr>
        </table>
      </form></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>
<%
fid=request("fid")
if fid="" then fid=1
mmonth=request("mmonth")
mday=request("mday")
myear=request("myear")
if myear="" then myear=2004
%>
	<table width="100%" border="0" align="center" cellpadding="2" cellspacing="2">
        <tr> 
          <td> 
<%
msql="select * from main where fid="& fid & " and year=" & myear
if mmonth<>"" then msql = msql & " and month=" & mmonth
if mday<>"" then msql = msql & " and day=" & mday
msql = msql & " order by year , month, day ,quantum " 
Set mrs=Server.CreateObject("ADODB.RecordSet") 
mrs.Open msql,conn,1,1 
%> 
            <table width="100%" border="1" align="center" cellpadding="2" cellspacing="0" bordercolor="#CCCCCC">
              <%
if mrs.eof or mrs.bof then
response.Write("<tr bgcolor=#FFFFFF><td height=25 colspan=11> The field have not be booked . <a href=field_book_c.asp> <b> Add Field Booking </b></a></td></tr></table>")
response.End()
mrs.close
set mrs=nothing
else
%>
              <tr bgcolor="#FFFFFF"> 
                <td height="25" colspan="8"><b> Field ( <font color="#FF0000"><%=mrs("fname")%></font> ) Booking's Status</b></td>
              </tr>
              <tr bgcolor="#FFFFFF"> 
                <td colspan="8">&nbsp; </td>
              </tr>
              <tr> 
                <td width="63" align="center" bgcolor="#FFFFFF"><strong>NO.</strong></td>
                <td width="99" align="center" bgcolor="#FFFFFF"><b>Booking Date</b><b></b></td>
                <td width="50" align="center" bgcolor="#FFFFFF"><strong>Quantum</strong></td>
                <td width="91" align="center" bgcolor="#FFFFFF"><b>User Name</b></td>
                <td width="263" align="center" bgcolor="#FFFFFF"><strong>Description</strong></td>
                <td width="59" align="center" bgcolor="#FFFFFF"><b>Operation</b></td>
              </tr>
              <%
do while not mrs.eof
%>
              <tr> 
                <td align="center" bgcolor="#f5f5f5"><%=mrs("id")%>&nbsp; </td>
                <td align="center" bgcolor="#f5f5f5"><%=mrs("ddate")%>&nbsp; </td>
                <td height="35" align="center" bgcolor="#f5f5f5"><%=mrs("quantum")%>&nbsp; </td>
                <td align="center" bgcolor="#f5f5f5"><%=mrs("username")%>&nbsp;</td>
                <td align="left" valign="top" bgcolor="#f5f5f5"> <%=mrs("description")%>&nbsp;</td>
                <td align="center" bgcolor="#f5f5f5"><a href="field_book_v.asp?option=del&id=<%=mrs("id")%>" onClick="return confirm(' Are you sure to del the tournament ?')">Del</a></td>
              </tr>
              <%
mrs.movenext
loop
mrs.close
set mrs=nothing
end if
%>
            </table>
            <!-- ========================================================================================================		  

													fixture or result ============stop

 ======================================================================================================== -->
          </td>
        </tr>
      </table></td>
    <td>&nbsp;</td>
  </tr>
</table>
