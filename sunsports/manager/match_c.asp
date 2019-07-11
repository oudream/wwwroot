<SCRIPT LANGUAGE="JavaScript">
<!--

function gettarg(targ,selObj,restore){ //v3.0
  eval(targ+".location='match_c.asp?cid="+selObj.options[selObj.selectedIndex].value+"'");
  if (restore) selObj.selectedIndex=0;
}

function checkform()
{
	if(! isNumber(zform.mweek.value)) {
		alert('Use integers only for "Match Week" field');
		zform.mweek.focus();
		return false;
		}
	if(! isNumber(zform.mmonth.value)) {
		alert('Use integers only for "Match Month" field');
		zform.mmonth.focus();
		return false;
		}
	if (zform.mmonth.value > 12) {
		alert('Please enter a valid "Match Month"');
		zform.mmonth.focus();
		return false;
		}
	if(! isNumber(zform.mday.value)) {
		alert('Use integers only for "Match Day" field');
		zform.mday.focus();
		return false;
		}
	if (zform.mday.value > 31) {
		alert('Please enter a valid "Match Day"');
		zform.mday.focus();
		return false;
		}
	if(! isNumber(zform.myear.value)) {
		alert('Use integers only for "Match Year" field');
		zform.myear.focus();
		return false;
		}
	if (zform.myear.value > 2099) {
		alert('Please enter a valid "Match Year"');
		zform.myear.focus();
		return false;
		}
	if(! isNumber(zform.mhour.value)) {
		alert('Use integers only for "Match Hour" field');
		zform.mhour.focus();
		return false;
		}
	if (zform.mhour.value > 24) {
		alert('Please enter a valid "Match Hour"');
		zform.mhour.focus();
		return false;
		}
	if(! isNumber(zform.mminute.value)) {
		alert('Use integers only for "Match Minute" field');
		zform.mminute.focus();
		return false;
		}
	if (zform.mminute.value > 60) {
		alert('Please enter a valid "Match Minute"');
		zform.mminute.focus();
		return false;
		}
	if (zform.fid.value.length == 0) {
		alert('Please enter a valid "Match Field"');
		zform.fid.focus();
		return false;
		}
	if (zform.htid.value.length == 0) {
		alert('Please enter a valid "Match Home Team"');
		zform.htid.focus();
		return false;
		}
	if (zform.atid.value.length == 0) {
		alert('Please enter a valid "Match Away Team"');
		zform.atid.focus();
		return false;
		}
	if (zform.htid.value==zform.atid.value) {
		alert('"Match Home Team" Can not equal "Match Away Team" !');
		zform.htid.focus();
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
cid=request("cid")
cname=request("cname")
if cid="" then cid=210
csql="select * from tournament where cid="&cid
Set crs=Server.CreateObject("ADODB.RecordSet")
crs.Open csql,conn,1,1
if crs.eof or crs.bof then
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' No league you want in the databse ! ');</SCRIPT>")
response.End()
crs.close
set crs=nothing
end if
cname=crs("name")
crs.close
set crs=nothing
%>
<%
dim ateam()
tsql="select tid from ct where cid="&cid
Set trs=Server.CreateObject("ADODB.RecordSet") 
trs.Open tsql,conn,1,1 
if not ( trs.eof or trs.bof ) then
redim ateam(trs.recordcount-1)
for i=0 to ubound(ateam)
ateam(i)=trs("tid")
trs.movenext
next
trs.close
set trs=nothing
else
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' There are not team in The League(cup) ! ');</SCRIPT>")
%>
            
            <table width="96%" border="0" align="center" cellpadding="2" cellspacing="2">
              <tr>
                <td align="center"><strong><a href="tournament_a.asp?cid=<%=cid%>">ADD 
                  TEAM TO THE TOURNAMENT</a></strong></td>
              </tr>
            </table> 
<%
response.End()
end if
%>
<%
if request("option")="add" then

mweek=trim(request("mweek"))
mmonth=trim(request("mmonth"))
mday=trim(request("mday"))
myear=trim(request("myear"))
mhour=trim(request("mhour"))
mminute=trim(request("mminute"))
if mweek="" or mmonth="" or mday="" or myear="" or mhour="" or mminute="" then 
response.write("<SCRIPT LANGUAGE=JavaScript> alert(' Please Enter week and date and time againt. ');</SCRIPT>")
response.End()
end if

fstr=split(trim(request("fid")),"zooz",-1,1)
if ubound(fstr)<>1 then
response.write("<SCRIPT LANGUAGE=JavaScript> alert(' The frield is error ! ');</SCRIPT>")
response.End()
end if
fid=clng(fstr(0))
fname=fstr(1)

htstr=split(trim(request("htid")),"zooz",-1,1)
if ubound(htstr)<>2 then
response.write("<SCRIPT LANGUAGE=JavaScript> alert(' The home team is error ! ');</SCRIPT>")
response.End()
end if
htid=clng(htstr(0))
htname=htstr(1)
hhcolor=htstr(2)

atstr=split(trim(request("atid")),"zooz",-1,1)
if ubound(htstr)<>2 then
response.write("<SCRIPT LANGUAGE=JavaScript> alert(' The away team is error ! ');</SCRIPT>")
response.End()
end if
atid=clng(atstr(0))
atname=atstr(1)
acolor=atstr(2)

hhconfirm=trim(request("hconfirm"))
if hhconfirm="" then hhconfirm="no"
aconfirm=trim(request("aconfirm"))
if aconfirm="" then aconfirm="no"

if fid="" or fname="" or htid="" or htname="" or hhcolor="" or atid="" or atname="" or acolor="" then
response.write("<SCRIPT LANGUAGE=JavaScript> alert(' Your database is error ! ');</SCRIPT>")
response.End()
end if

sql="insert into match (cid,cname,mweek,mmonth,mday,myear,mhour,mminute,fid,fname,htid,htname,hcolor,atid,atname,acolor,hconfirm,aconfirm,inserttime) values ("&cid&",'"&cname&"',"&mweek&","&mmonth&","&mday&","&myear&","&mhour&","&mminute&","&fid&",'"&fname&"',"&htid&",'"&htname&"','"&hhcolor&"',"&atid&",'"&atname&"','"&acolor&"','"&hhconfirm&"','"&aconfirm&"','"&now()&"')"
conn.Execute(sql)
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Create standing is complete. ');</SCRIPT>")
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
          <td height="50" align="center" valign="middle"><strong><font color="#FF0000"><%=cname%></font>Match Create</strong></td>
        </tr>
      </table></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td align="center">
<table width="100%" cellpadding=0 cellspacing=0>
                <tbody>
                  <tr> 
                    <td width="50%" align="right"><b><font color=#000066 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        size=2>Select the league the match in:</font></b></td>
                    <td width=50% align="left"> 
					<select name="cid" class="display_dropdown" id="cid" onchange="gettarg('self',this,0)">
                        <option selected value="">SELECT ... </option>
<%
sql="select * from tournament order by cid desc" 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.Open sql,conn,1,1 
if rs.bof or rs.eof then response.Redirect("error.asp?err=1")
do while not rs.eof
%>
                        <option value=<%=rs("cid")%>><%=rs("name")%> </option>
                        <%
rs.movenext
loop
rs.close
set rs=nothing
%>
                      </select></td>
                  </tr>
                </tbody>
              </table>
<form id="zform" action="match_c.asp" onSubmit="return checkform();" method="post">
	    <table width="100%" border="0" cellspacing="2" cellpadding="2">
          <tr align="center" valign="middle"> 
            <td height="10" colspan="3"></td>
          </tr>
          <tr> 
            <td width="21%">Match Week :&nbsp;</td>
            <td width="75%"><input name="mweek" id="mweek" type="text" size="5" maxlength="2" ></td>
            <td width="4%">&nbsp;</td>
          </tr>
          <tr> 
            <td width="21%">Match Date :&nbsp;</td>
            <td width="75%"><input name="mmonth" id="mmonth" type="text" size="5" maxlength="2" >
              -Month- 
              <input name="mday" id="mday" type="text" size="5" maxlength="2" >
              -Day- 
              <input name="myear" id="myear" type="text" size="5" maxlength="4" >
              -Year</td>
            <td width="4%">&nbsp;</td>
          </tr>
          <tr> 
            <td width="21%">Match Time :&nbsp;</td>
            <td width="75%"><input name="mhour" id="mhour" type="text" size="5" maxlength="2" >
              -Hour- 
              <input name="mminute" id="mminute" type="text" size="5" maxlength="2" >
              -Minute</td>
            <td width="4%">&nbsp;</td>
          </tr>
          <tr> 
            <td width="21%">Match Field :&nbsp;</td>
            <td width="75%">
			<select class="display_dropdown" id="fid" name="fid">
                <option selected value="">SELECT ... </option>
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
            <td width="21%">Home team :&nbsp;</td>
            <td width="75%"><select name="htid" class="display_dropdown" id="htid">
                <option selected value="">SELECT ... </option>
<%
tsql="select * from team where tid=0 "
for i=0 to ubound(ateam)
tsql=tsql&" or tid="&ateam(i)
next
Set trs=Server.CreateObject("ADODB.RecordSet") 
trs.Open tsql,conn,1,1 
if not ( trs.bof or trs.eof ) then
do while not trs.eof
%>
                <option value="<%=cstr(trs("tid")) & "zooz" & trs("name") & "zooz" & trs("hcolor")%>"><%=trs("tid")%>.<%=trs("name")%>.<font color="<%=trs("hcolor")%>">Home-Color(<%=trs("hcolorname")%>)</font> </option>
                <option value="<%=cstr(trs("tid")) & "zooz" & trs("name") & "zooz" & trs("acolor")%>"><%=trs("tid")%>.<%=trs("name")%>.<font color="<%=trs("acolor")%>">Home-Color(<%=trs("acolorname")%>)</font> </option>
<%
trs.movenext
loop
end if
trs.close
set trs=nothing
%>
              </select></td>
            <td width="4%">&nbsp;</td>
          </tr>
          <tr> 
            <td width="21%">Home Confirm :&nbsp;</td>
            <td width="75%"><input name="hconfirm" type="checkbox" id="hconfirm" value="yes"></td>
            <td width="4%">&nbsp;</td>
          </tr>
          <tr> 
            <td>Away team : </td>
            <td><select name="atid" class="display_dropdown" id="atid">
                <option selected value="">SELECT ... </option>
<%
tsql="select * from team where tid=0 "
for i=0 to ubound(ateam)
tsql=tsql&" or tid="&ateam(i)
next
Set trs=Server.CreateObject("ADODB.RecordSet") 
trs.Open tsql,conn,1,1 
if not ( trs.bof or trs.eof ) then
do while not trs.eof
%>
                <option value="<%=cstr(trs("tid")) & "zooz" & trs("name") & "zooz" & trs("hcolor")%>"><%=trs("tid")%>.<%=trs("name")%>.<font color="<%=trs("hcolor")%>">Home-Color(<%=trs("hcolorname")%>)</font> </option>
                <option value="<%=cstr(trs("tid")) & "zooz" & trs("name") & "zooz" & trs("acolor")%>"><%=trs("tid")%>.<%=trs("name")%>.<font color="<%=trs("acolor")%>">Home-Color(<%=trs("acolorname")%>)</font> </option>
<%
trs.movenext
loop
end if
trs.close
set trs=nothing
%>
              </select></td>
            <td>&nbsp;</td>
          </tr>
          <tr> 
            <td width="21%">Away Confirm :&nbsp;</td>
            <td width="75%"><input name="aconfirm" type="checkbox" id="aconfirm" value="yes"></td>
            <td width="4%">&nbsp;</td>
          </tr>
          <tr> 
            <td>&nbsp;</td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
          </tr>
          <tr> 
            <td colspan="3">
<%
csql="select * from tournament where cid="&cid
Set crs=Server.CreateObject("ADODB.RecordSet") 
crs.Open csql,conn,1,1
if not ( crs.eof or crs.bof ) then
%>
        <input type="hidden" id="cid" name="cid" value=<%=crs("cid")%>> 
        <input type="hidden" id="cname" name="cname" value=<%=crs("name")%>> 
<%
end if
crs.close
set crs=nothing
%>
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

<br>
<br>
<%=cid%>