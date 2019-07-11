<SCRIPT LANGUAGE="JavaScript">
<!--
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
	if (zform.myear.value > 2099 || zform.myear.value < 2000) {
		alert('"Match Year" must be 2000-2099 .');
		zform.myear.focus();
		return false;
		}
	if(! isNumber(zform.mhour.value)) {
		alert('Use integers only for "Match Hour" field');
		zform.mhour.focus();
		return false;
		}
	if (zform.mhour.value > 24 && zform.mhour.value < 1) {
		alert('Please enter a valid "Match Hour"');
		zform.mhour.focus();
		return false;
		}
	if(! isNumber(zform.mminute.value)) {
		alert('Use integers only for "Match Minute" field');
		zform.mminute.focus();
		return false;
		}
	if (zform.mminute.value > 60 && zform.mminute.value < 1) {
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
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="css.css" type=text/css rel=stylesheet>
</head>
<body text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!--#include file="adminconn.asp" -->
<%
cid=request("cid")
cname=request("cname")
if cid="" then cid=230
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
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' There are no team in The League(cup/friendly) Match! ');</SCRIPT>")
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

fstr=request("fid")
fstr=replace(fstr,"'","''")
fstr=split(trim(fstr),"zooz",-1,1)
if ubound(fstr)<>1 then
response.write("<SCRIPT LANGUAGE=JavaScript> alert(' The frield is error ! ');</SCRIPT>")
response.End()
end if
fid=clng(fstr(0))
fname=fstr(1)

htstr=request("htid")
htstr=replace(htstr,"'","''")
htstr=split(trim(htstr),"zooz",-1,1)
if ubound(htstr)<>2 then
response.write("<SCRIPT LANGUAGE=JavaScript> alert(' The home team is error ! ');</SCRIPT>")
response.End()
end if
htid=clng(htstr(0))
htname=htstr(1)
hhcolor=htstr(2)

atstr=request("atid")
atstr=replace(atstr,"'","''")
atstr=split(trim(atstr),"zooz",-1,1)
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


msql="select * from match order by mid desc"
Set mrs=Server.CreateObject("ADODB.RecordSet")
mrs.Open msql,conn,1,1
mmid=mrs("mid")
htid=mrs("htid")
atid=mrs("atid")

tsql="select * from rresult where mid="&mmid&" and tid="&htid
Set trs=Server.CreateObject("ADODB.RecordSet") 
trs.Open tsql,conn,1,1 
if trs.eof or trs.bof then
conn.Execute("insert into rresult (mid,tid,hscore,ascore,refer) values ("&mmid&","&htid&","&request("hhacountscore")&","&request("aacountscore")&",10)")
else
conn.Execute( "update rresult set hscore="&request("hhacountscore")&", ascore="&request("aacountscore")&" where mid="&mmid&" and tid="&htid )
end if
trs.close
tsql="select * from rresult where mid="&mmid&" and tid="&atid
trs.Open tsql,conn,1,1 
if trs.eof or trs.bof then
conn.Execute("insert into rresult (mid,tid,hscore,ascore,refer) values ("&mmid&","&atid&","&request("hhacountscore")&","&request("aacountscore")&",10)")
else
conn.Execute( "update rresult set hscore="&request("hhacountscore")&", ascore="&request("aacountscore")&" where mid="&mmid&" and tid="&atid )
end if
trs.close
set trs=nothing


mrs.close
set mrs=nothing

response.write("<SCRIPT LANGUAGE=JavaScript> alert(' Your 1111 is error ! ');</SCRIPT>")
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
          <td height="50" align="center" valign="middle"><H2><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000066>Fixture Creation</FONT></H2></td>
        </tr>
      </table></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td> 
			<input type="hidden" name="cid" id="cid" value=230>
<br>
<br>
      <form id="zform" action="data_match_c1.asp" onSubmit="return checkform();" method="post">
        <table width="100%" border="1" cellspacing="0" cellpadding="0">
          <tr bgcolor="#999999"> 
            <td height="10" colspan="7"><strong>Match Week :&nbsp; 
              <input name="mweek" id="mweek" type="text" size="5" maxlength="2" class="form">
              </strong></td>
          </tr>
          <tr> 
            <td align="center"><strong>Date</strong></td>
            <td align="center"><strong>Time</strong></td>
            <td align="center"><strong>HOME</strong></td>
            <td colspan="2" align="center"><strong>vs</strong></td>
            <td align="center"><strong>AWAY</strong></td>
            <td align="center"><strong>Location</strong></td>
          </tr>
          <tr> 
            <td><input name="mday" id="mday2" type="text" size="3" maxlength="2" class="form">
              / 
              <input name="mmonth" id="mmonth2" type="text" size="3" maxlength="2" class="form">
              <input name="myear" id="myear2" type="text" size="5" maxlength="4" value="2003" class="form">
            </td>
            <td><input name="mhour" id="mhour2" type="text" size="3" maxlength="2"  class="form">
              : 
              <input name="mminute" id="mminute2" type="text" size="3" maxlength="2" class="form"></td>
            <td>
<input name="hconfirm" type="hidden" id="hconfirm2" value="yes"> 
              <select name="htid" class="display_dropdown" id="select2">
                <option selected value="">SELECT ... </option>
                <option>---------------------</option>
                <%
tsql="select * from team where tid=0 "
for i=0 to ubound(ateam)
tsql=tsql&" or tid="&ateam(i)
next
tsql=tsql & " order by name "
Set trs=Server.CreateObject("ADODB.RecordSet") 
trs.Open tsql,conn,1,1 
if not ( trs.bof or trs.eof ) then
do while not trs.eof
%>
                <option value="<%=cstr(trs("tid")) & "zooz" & trs("name") & "zooz" & trs("hcolor")%>"><%=trs("name")%><font color="<%=trs("hcolor")%>"> 
                (HomeColor: <%=trs("hcolorname")%>)</font> </option>
                <option value="<%=cstr(trs("tid")) & "zooz" & trs("name") & "zooz" & trs("acolor")%>"><%=trs("name")%><font color="<%=trs("acolor")%>"> 
                (AwayColor: <%=trs("acolorname")%>)</font> </option>
                <%
trs.movenext
loop
end if
trs.close
set trs=nothing
%>
              </select></td>
            <td><input name="hhacountscore" type="text" id="hhacountscore2" size="3" maxlength="2" class="form"></td>
            <td><input name="aacountscore" id="aacountscore2" type="text" size="3" maxlength="2" class="form"></td>
            <td><input name="aconfirm" type="hidden" id="aconfirm2" value="yes"> 
              <select name="atid" class="display_dropdown" id="select3">
                <option selected value="">SELECT ... </option>
                <option>---------------------</option>
                <%
tsql="select * from team where tid=0 "
for i=0 to ubound(ateam)
tsql=tsql&" or tid="&ateam(i)
next
tsql=tsql & " order by name "
Set trs=Server.CreateObject("ADODB.RecordSet") 
trs.Open tsql,conn,1,1 
if not ( trs.bof or trs.eof ) then
do while not trs.eof
%>
                <option value="<%=cstr(trs("tid")) & "zooz" & trs("name") & "zooz" & trs("hcolor")%>"><%=trs("name")%><font color="<%=trs("hcolor")%>"> 
                (HomeColor: <%=trs("hcolorname")%>)</font> </option>
                <option value="<%=cstr(trs("tid")) & "zooz" & trs("name") & "zooz" & trs("acolor")%>"><%=trs("name")%><font color="<%=trs("acolor")%>"> 
                (AwayColor: <%=trs("acolorname")%>)</font> </option>
                <%
trs.movenext
loop
end if
trs.close
set trs=nothing
%>
              </select></td>
            <td><select class="display_dropdown" id="fid" name="fid">
                <option selected value="">SELECT ... </option>
                <option>---------------------</option>
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
          </tr>
          <td colspan="7"> <%
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
%> <input type="submit" id="option" name="option" value="add"> 
            &nbsp;</td>
          </tr>
        </table>
</form>
</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td height="100">&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
</table>

<br>
<br>
</body>
</html>