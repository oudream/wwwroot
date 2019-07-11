<SCRIPT LANGUAGE="JavaScript">
<!--

function checkform()
{
	if (zform.Played.value.length == 0) {
		alert('Please enter a valid Played');
		zform.Played.focus();
		return false;
		}
	if(! isNumber(zform.Played.value)) {
		alert('Use integers only for "Played" field');
		zform.Played.focus();
		return false;
		}
	if(! isNumber(zform.Win.value)) {
		alert('Use integers only for "Win" field');
		zform.Win.focus();
		return false;
		}
	if(! isNumber(zform.Lose.value)) {
		alert('Use integers only for "Lose" field');
		zform.Lose.focus();
		return false;
		}
	if(! isNumber(zform.Draw.value)) {
		alert('Use integers only for "Draw" field');
		zform.Draw.focus();
		return false;
		}
	if(! isNumber(zform.Scored.value)) {
		alert('Use integers only for "Scored" field');
		zform.Scored.focus();
		return false;
		}
	if(! isNumber(zform.Conceded.value)) {
		alert('Use integers only for "Conceded" field');
		zform.Conceded.focus();
		return false;
		}
	if(! isNumber(zform.Gd.value)) {
		alert('Use integers only for "Gd" field');
		zform.Gd.focus();
		return false;
		}
	if(! isNumber(zform.Points.value)) {
		alert('Use integers only for "Points" field');
		zform.Points.focus();
		return false;
		}
	return true;
}

function isNumber(name)
{
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
tid=request("tid")
if tid="" then
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Please select the team you want to add standing at first ! ');</SCRIPT>")
response.End()
end if
%>
<%
if request("option")="edit" then
played=trim(request("played"))
if played<>"" and tid<>"" then
win=request("win")
if win="" then win=0
lose=request("lose")
if lose="" then lose=0
draw=request("draw")
if draw="" then draw=0
scored=request("scored")
if scored="" then scored=0
conceded=request("conceded")
if conceded="" then conceded=0
gd=request("gd")
if gd="" then gd=0
points=request("points")
if points="" then points=0
sql="update standing set played="&played&",win="&win&",lose="&lose&",draw="&draw&",scored="&scored&",conceded="&conceded&",gd="&gd&",points="&points&" where tid=" & tid
conn.Execute(sql)
response.Redirect("team_table.asp?option=edittableissucc&cid="&request("cid"))
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
          <td height="30" align="center" valign="middle"><strong>Team's Standing 
            / Table Edit</strong></td>
        </tr>
      </table></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>
<%
tsql="select * from standing where tid="&tid
Set trs=Server.CreateObject("ADODB.RecordSet") 
trs.Open tsql,conn,1,1
if not ( trs.eof or trs.bof ) then
%>
<form id="zform" action="team_table_e.asp" onSubmit="return checkform();" method="post">
	    <table width="100%" border="0" cellspacing="2" cellpadding="2">
          <tr> 
            <td colspan="3"><strong>Update standings:</strong></td>
          </tr>
          <tr> 
            <td width="12%">Played:&nbsp;</td>
            <td width="84%"><input name="Played" id="Played" type="text" size="10" maxlength="5" value=<%=trs("Played")%> ></td>
            <td width="4%">&nbsp;</td>
          </tr>
          <tr> 
            <td>Win: </td>
            <td><input name="Win" id="Win" type="text" size="10" maxlength="5" value=<%=trs("Win")%> ></td>
            <td>&nbsp;</td>
          </tr>
          <tr> 
            <td>Lose: </td>
            <td><input name="Lose" id="Lose" type="text" size="10" maxlength="5" value=<%=trs("Lose")%> ></td>
            <td>&nbsp;</td>
          </tr>
          <tr> 
            <td>Draw: </td>
            <td><input name="Draw" id="Draw" type="text" size="10" maxlength="5" value=<%=trs("Draw")%> ></td>
            <td>&nbsp;</td>
          </tr>
          <tr> 
            <td>Scored: </td>
            <td><input name="Scored" id="Scored" type="text" size="10" maxlength="5" value=<%=trs("Scored")%> ></td>
            <td>&nbsp;</td>
          </tr>
          <tr> 
            <td>Conceded:</td>
            <td><input name="Conceded" id="Conceded" type="text" size="10" maxlength="5" value=<%=trs("Conceded")%> ></td>
            <td>&nbsp;</td>
          </tr>
          <tr> 
            <td>Gd: </td>
            <td><input name="Gd" id="Gd" type="text" size="10" maxlength="5" value=<%=trs("Gd")%> ></td>
            <td>&nbsp;</td>
          </tr>
          <tr> 
            <td>Points:</td>
            <td><input name="Points" id="Points" type="text" size="10" maxlength="5" value=<%=trs("Points")%> ></td>
            <td>&nbsp;</td>
          </tr>
          <tr> 
            <td colspan="3">
        <input type="hidden" id="tid" name="tid" value=<%=trs("tid")%>> 
        <input type="hidden" id="tname" name="tname" value=<%=trs("tname")%>> 
        <input type="submit" id="option" name="option" value="edit"> 
        <input type="button" name="goback" value="go back" onClick="history.go( -1 );return true;"> 
			&nbsp;</td>
          </tr>
        </table>
</form>
<%
end if
trs.close
set trs=nothing
%>
</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
</table>
