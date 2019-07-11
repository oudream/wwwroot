<SCRIPT LANGUAGE="JavaScript">
<!--
function checkhowform()
{
	if(! isNumber(howform.snum.value)) {
		alert('Please enter a valid "Player NO. of scorer"');
		howform.snum.focus();
		return false;
		}
	if(! isNumber(howform.dnum.value)) {
		alert('Please enter a valid "Player NO. of discipline"');
		howform.dnum.focus();
		return false;
		}
}


function checkzform(zformsnumvalue,zformdnumvalue)
{
if(zformsnumvalue>0 || zformdnumvalue>0){
	if(! isNumber(zform.hscore.value)) {
		alert('Please enter a valid "Home of score"');
		zform.hscore.focus();
		return false;
		}
	if(! isNumber(zform.ascore.value)) {
		alert('Please enter a valid "Away of score"');
		zform.ascore.focus();
		return false;
		}
	if(zform.refer.value.length == 0) {
		alert('Please select a "refer"');
		zform.refer.focus();
		return false;
		}
}
if(zformsnumvalue>0 && zformsnumvalue!=2){
j=0;
	for(i = 0; i < zformsnumvalue; i++) 
	{
	if(! isNumber(eval("zform.s" + j + ".value"))) {
		alert('Please enter a valid "Player NO. of scorer"  '+j);
		eval("zform.s" + j).focus();
		return false;
		}
	if( j == zformsnumvalue-1 )
		break;
	j=j+1;
	}
}
if(zformsnumvalue==2){
	if(! isNumber(zform.s0.value)) {
		alert('Please enter a valid "Player NO. of scorer"');
		zform.s0.focus();
		return false;
		}
	if(! isNumber(zform.s1.value)) {
		alert('Please enter a valid "Player NO. of scorer"');
		zform.s1.focus();
		return false;
		}
}
if(zformdnumvalue>0 && zformdnumvalue!=2){
j=0;
	for(i = 0; i < zformdnumvalue; i++) 
	{
	if(! isNumber(eval("zform.d" + j + ".value"))) {
		alert('Please enter a valid "Player NO. of discipline"  '+j);
		eval("zform.d" + j).focus();
		return false;
		}
	if( j == zformdnumvalue-1 )
		break;
	j=j+1;
	}
}
if(zformdnumvalue==2){
	if(! isNumber(zform.d0.value)) {
		alert('Please enter a valid "Player NO. of discipline"');
		zform.d0.focus();
		return false;
		}
	if(! isNumber(zform.d1.value)) {
		alert('Please enter a valid "Player NO. of discipline"');
		zform.d1.focus();
		return false;
		}
}
return true;
}

function isNumber(name)
{
		if(name.length == 0)
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
if request("option")="createresultissucc" then response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Create match is complete. ');</SCRIPT>")
if request("option")="editresultissucc" then response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Update match is complete. ');</SCRIPT>")
if request("option")="deledit" then response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' The matchs result is deleted , now you can add a now one ! ');</SCRIPT>")
tid=session("manager_tid")
mmid=request("mid")
if tid="" and mmid="" then
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Please select the team you want to add scorer or discipline at first ! ');</SCRIPT>")
response.End()
end if
%>
<%
snum=request("snum")
dnum=request("dnum")
hscore=request("hscore")
ascore=request("ascore")
refer=request("refer")
%>
<%
msql="select * from match where mid="&mmid
Set mrs=Server.CreateObject("ADODB.RecordSet")
mrs.Open msql,conn,1,1
if mrs.eof then
mrs.close
set mrs=nothing
response.End()
end if
%>
<%
dim asi()
dim aspid()
dim adi()
dim adpid()
dim aci()

redim asi(-1)
redim aspid(-1)
redim adi(-1)
redim adpid(-1)
redim aci(-1)
hdhsnum=0
hdhcnum=0

i=0
sql="select * from scorer where mid=" & mmid &" and tid=" & tid
Set rs=Server.CreateObject("ADODB.RecordSet")
rs.Open sql,conn,1,1
if not rs.eof then
hdhsnum=rs.recordcount
redim asi(hdhsnum-1)
redim aspid(hdhsnum-1)
do while not rs.eof
asi(i)=rs("score")
aspid(i)=rs("pid")
i=i+1
rs.movenext
loop
end if
rs.close

i=0
sql="select * from card where mid=" & mmid &" and tid=" & tid
rs.Open sql,conn,1,1
if not rs.eof then
hdhcnum=rs.recordcount
redim adi(hdhcnum-1)
redim adpid(hdhcnum-1)
redim aci(hdhcnum-1)
do while not rs.eof
adpid(i)=rs("pid")
if rs("yellow")<>0 then
adi(i)=rs("yellow")
aci(i)="yellow"
end if
if rs("red")<>0 then
adi(i)=rs("red")
aci(i)="red"
end if
i=i+1
rs.movenext
loop
end if
rs.close

%>
<%
if snum="" and dnum="" then
snum=hdhsnum
dnum=hdhcnum
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
          <td height="30" align="center" valign="middle"><strong>Team <font color="#FF0000"><%=session("manager_tname")%></font> Match Result 
            Create</strong></td>
        </tr>
      </table></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><table width="100%" border="0" cellspacing="2" cellpadding="2">
        <tr> 
          <td height="30" align="center" valign="middle">
<form id="howform" action="match_r_e.asp" onSubmit="return checkhowform();" method="post">
              <input type="hidden" id="mid" name="mid" value=<%=mmid%>> 
              <table width="100%" border="0" cellspacing="2" cellpadding="2">
                <tr align="center" valign="middle"> 
                  <td height="30" colspan="3"><table width="100%" border="0" cellspacing="2" cellpadding="2">
                      <tr>
                        <td width="33%"><font color="#FF0000">STEP 1</font></td>
                        <td width="50%" align="center">&nbsp;</td>
                        <td width="33%">&nbsp;</td>
                      </tr>
                    </table>
                  </td>
                </tr>

                <tr> 
                  <td width="51%">&nbsp;</td>
                  <td width="46%">&nbsp;</td>
                  <td width="3%">&nbsp;</td>
                </tr>
                <tr> 
                  <td width="51%">How many player scorer :</td>
                  <td width="46%"><input name="snum" type="text" id="snum" size="5" maxlength="2" value=<%=snum%>></td>
                  <td width="3%">&nbsp;</td>
                </tr>
                <tr> 
                  <td>How many player discipline :</td>
                  <td><input name="dnum" id="dnum" type="text" size="5" maxlength="2" value=<%=dnum%>></td>
                  <td>&nbsp;</td>
                </tr>
                <tr> 
                  <td height="50" colspan="3"><input type="submit" name="option" id="option" value=" STEP 2 -> "></td>
                </tr>
              </table>
</form>
			</td>
        </tr>
      </table></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>



























<form id="zform" action="match_r_save.asp" onSubmit="return checkzform(<%=snum%>,<%=dnum%>);" method="post">
              <input type="hidden" id="mid" name="mid" value=<%=mmid%>> 
			<input name="dnum" id="dnum" type="hidden" size="5" maxlength="2" value=<%=dnum%>>
			<input name="snum" type="hidden" id="snum" size="5" maxlength="2" value=<%=snum%>>
	    <table width="100%" border="1" cellspacing="2" cellpadding="2">
          <tr align="center" valign="middle"> 
            <td height="30" colspan="4"><table width="100%" border="0" cellspacing="2" cellpadding="2">
                <tr> 
                  <td width="28%"><font color="#FF0000">STEP 2</font></td>
                  <td width="43%" align="center">&nbsp;</td>
                  <td width="29%">&nbsp;</td>
                </tr>
              </table></td>
          </tr>
          <tr> 
            <td width="17%">&nbsp;</td>
            <td width="26%">&nbsp;</td>
            <td width="32%">&nbsp;</td>
            <td width="25%">&nbsp;</td>
          </tr>
          <tr> 
            <td width="17%"><b>Week</b></td>
            <td width="26%"><%=mrs("mweek")%></td>
            <td width="32%">&nbsp;</td>
            <td width="25%">&nbsp;</td>
          </tr>
          <tr> 
            <td width="17%"><b>Date</b></td>
            <td width="26%"><%=mrs("mmonth")&"/"&mrs("mday")&"/"&mrs("myear")%></td>
            <td width="32%">&nbsp;</td>
            <td width="25%">&nbsp;</td>
          </tr>
          <tr> 
            <td width="17%"><b>Time</b></td>
            <td width="26%"><%=mrs("mhour")&":"&mrs("mminute")%></td>
            <td width="32%">&nbsp;</td>
            <td width="25%">&nbsp;</td>
          </tr>
          <%
sql="select * from rresult where mid=" & mmid &" and tid=" & tid
rs.Open sql,conn,1,1
if not rs.eof then
hscore=rs("hscore")
ascore=rs("ascore")
refer=rs("refer")
end if
rs.close
%>
          <tr> 
            <td width="17%"><b>Home Team</b></td>
            <td width="26%"><%=mrs("htname")%></td>
            <td width="32%"><input name="hscore" type="text" id="hscore" size="5" maxlength="2" value=<%=hscore%>> 
              <b>Score</b></td>
            <td width="25%">&nbsp;</td>
          </tr>
          <tr> 
            <td width="17%"><b>Away Team </b></td>
            <td width="26%"><%=mrs("atname")%></td>
            <td width="32%"><input name="ascore" type="text" id="ascore2" size="5" maxlength="2" value=<%=ascore%>> 
              <b>Score</b></td>
            <td width="25%">&nbsp;</td>
          </tr>
          <tr> 
            <td width="17%"><b>Location</b></td>
            <td width="26%"><%=mrs("fname")%></td>
            <td width="32%">&nbsp;</td>
            <td width="25%">&nbsp;</td>
          </tr>
          <tr> 
            <td width="17%">&nbsp;</td>
            <td width="26%">&nbsp;</td>
            <td width="32%">&nbsp;</td>
            <td width="25%">&nbsp;</td>
          </tr>
          <tr> 
            <td width="17%">&nbsp;</td>
            <td width="26%">&nbsp;</td>
            <td width="32%">&nbsp;</td>
            <td width="25%">&nbsp;</td>
          </tr>
          <tr> 
            <td width="17%">&nbsp;</td>
            <td width="26%">&nbsp;</td>
            <td width="32%">&nbsp;</td>
            <td width="25%">&nbsp;</td>
          </tr>
          <tr> 
            <td colspan="4"><font color="#FF0000">For <%=session("manager_tname")%></font></td>
          </tr>
          <tr> 
            <td width="17%">&nbsp;</td>
            <td width="26%">&nbsp;</td>
            <td width="32%">&nbsp;</td>
            <td width="25%">&nbsp;</td>
          </tr>
          <%
if snum>0 then
%>
          <tr> 
            <td height="30">Score : </td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
          </tr>
          <%
end if
for i=0 to snum-1
if i<= ubound(aspid) then
%>
          <tr> 
            <td>Player : </td>
            <td> <select name="<%="spid"&cstr(i)%>" id="<%="spid"&cstr(i)%>">
                <%
psql="select * from player where pid=" & aspid(i)
Set prs=Server.CreateObject("ADODB.RecordSet") 
prs.Open psql,conn,1,1 
if not prs.eof then
%>
                <option value=<%=prs("pid")%> selected><%=prs("name")%></option>
                <%
end if
prs.close
%>
                <%
psql="select * from player where pid<>" & aspid(i) &" and tid=" & tid
prs.Open psql,conn,1,1 
if prs.eof or prs.bof then
response.Write("<tr><td rowspan=5 align=center bgcolor=#FFFFFF>NO DATA</td></tr></table>")
response.End()
prs.close
set prs=nothing
end if
do while not prs.eof
%>
                <option value=<%=prs("pid")%>><%=prs("name")%></option>
                <%
prs.movenext
loop
prs.close
set prs=nothing 
%>
              </select> </td>
            <td> <input name="<%="s"&cstr(i)%>" id="<%="s"&cstr(i)%>" type="text" size="5" maxlength="2" value="<%=asi(i)%>">
              -score</td>
            <td>&nbsp;</td>
          </tr>
          <%
else
%>
          <tr> 
            <td>Player : </td>
            <td> <select name="<%="spid"&cstr(i)%>" id="<%="spid"&cstr(i)%>">
                <%
psql="select * from player where tid=" & tid
Set prs=Server.CreateObject("ADODB.RecordSet") 
prs.Open psql,conn,1,1 
if prs.eof or prs.bof then
response.Write("<tr><td rowspan=5 align=center bgcolor=#FFFFFF>NO DATA</td></tr></table>")
response.End()
prs.close
set prs=nothing
end if
do while not prs.eof
%>
                <option value=<%=prs("pid")%>><%=prs("name")%></option>
                <%
prs.movenext
loop
prs.close
set prs=nothing 
%>
              </select> </td>
            <td> <input name="<%="s"&cstr(i)%>" id="<%="s"&cstr(i)%>" type="text" size="5" maxlength="2">
              -score</td>
            <td>&nbsp;</td>
          </tr>
          <%
end if
next
%>
          <tr> 
            <td colspan="4">&nbsp;</td>
          </tr>
          <tr> 
            <td width="17%">&nbsp;</td>
            <td width="26%">&nbsp;</td>
            <td width="32%">&nbsp;</td>
            <td width="25%">&nbsp;</td>
          </tr>
          <%
if dnum>0 then
%>
          <tr> 
            <td height="30">discipline: </td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
          </tr>
          <%
end if
for i=0 to dnum-1
if i<=ubound(adpid) then
%>
          <tr> 
            <td>Player : </td>
            <td> <select name="<%="dpid"&cstr(i)%>" id="<%="dpid"&cstr(i)%>">
                <%
psql="select * from player where pid=" & adpid(i)
Set prs=Server.CreateObject("ADODB.RecordSet") 
prs.Open psql,conn,1,1 
if not prs.eof then
%>
                <option value=<%=prs("pid")%> selected><%=prs("name")%></option>
                <%
end if
prs.close
%>
                <%
psql="select * from player where pid<>" & adpid(i) &" and tid=" & tid
prs.Open psql,conn,1,1 
if prs.eof or prs.bof then
response.Write("<tr><td rowspan=5 align=center bgcolor=#FFFFFF>NO DATA</td></tr></table>")
response.End()
prs.close
set prs=nothing
end if
do while not prs.eof
%>
                <option value=<%=prs("pid")%>><%=prs("name")%></option>
                <%
prs.movenext
loop
prs.close
set prs=nothing 
%>
              </select> </td>
            <td><select name="<%="c"&cstr(i)%>" id="<%="c"&cstr(i)%>">
                <%
if aci(i)="yellow" then
%>
                <option value="<%=aci(i)%>" selected><%=aci(i)%></option>
                <option value="red">red</option>
                <%
elseif aci(i)="red" then
%>
                <option value="<%=aci(i)%>" selected><%=aci(i)%></option>
                <option value="yellow">yellow</option>
                <%
else
%>
                <option value="yellow" selected>yellow</option>
                <option value="red">red</option>
                <%
end if
%>
              </select>
              -Card </td>
            <td> <input name="<%="d"&cstr(i)%>" type="text" id="<%="d"&cstr(i)%>" size="5" maxlength="2" value=<%=adi(i)%>>
              -NUM</td>
          </tr>
          <%
else
%>
          <tr> 
            <td>Player : </td>
            <td> <select name="<%="dpid"&cstr(i)%>" id="<%="dpid"&cstr(i)%>">
                <%
psql="select * from player where tid=" & tid
Set prs=Server.CreateObject("ADODB.RecordSet") 
prs.Open psql,conn,1,1 
if prs.eof or prs.bof then
response.Write("<tr><td rowspan=5 align=center bgcolor=#FFFFFF>NO DATA</td></tr></table>")
response.End()
prs.close
set prs=nothing
end if
do while not prs.eof
%>
                <option value=<%=prs("pid")%>><%=prs("name")%></option>
                <%
prs.movenext
loop
prs.close
set prs=nothing 
%>
              </select> </td>
            <td><select name="<%="c"&cstr(i)%>" id="<%="c"&cstr(i)%>">
                <option value="yellow" selected>yellow</option>
                <option value="red">red</option>
              </select>
              -Card </td>
            <td> <input name="<%="d"&cstr(i)%>" type="text" id="<%="d"&cstr(i)%>" size="5" maxlength="2">
              -NUM</td>
          </tr>
          <%
end if
next
%>
          <tr> 
            <td width="17%">&nbsp;</td>
            <td width="26%">&nbsp;</td>
            <td width="32%">&nbsp;</td>
            <td width="25%">&nbsp;</td>
          </tr>
          <tr> 
            <td width="17%">&nbsp;</td>
            <td width="26%">&nbsp;</td>
            <td width="32%">&nbsp;</td>
            <td width="25%">&nbsp;</td>
          </tr>
<%
if session("manager_tid")=mrs("htid") then
homeaway="h"
anothertname=mrs("atname")
else
homeaway="a"
anothertname=mrs("htname")
end if
%>
          <tr> 
            <td colspan="4"><font color="#FF0000">For <%=anothertname%></font></td>
          </tr>
          <tr> 
            <td colspan="4">Yellow Card : <input name="anotheryellow" id="anotheryellow" type="text" size="5" maxlength="2">&nbsp;&nbsp;&nbsp;
              Red Card : 
              <input name="anotherred" id="anotherred" type="text" size="5" maxlength="2"></td>
          </tr>
          <tr> 
            <td colspan="4">&nbsp;</td>
          </tr>
          <tr> 
            <td width="17%">Refer :&nbsp;</td>
            <td width="26%"><select name="refer" id="select2">
                <option value="<%=refer%>" selected><%=refer%></option>
                <%
i=1
do while i<=10
if i<>refer then
%>
                <option value="<%=i%>"><%=i%></option>
                <%
end if
i=i+1
loop
%>
              </select></td>
            <td width="32%">&nbsp;</td>
            <td width="25%">&nbsp;</td>
          </tr>
          <tr> 
            <td colspan="4">&nbsp;</td>
          </tr>
          <tr> 
            <td colspan="4">
			<input type="hidden" id="homeaway" name="homeaway" value="<%=homeaway%>">
			<input type="submit" id="option" name="option" value="update"> 
            <input type="button" name="goback" value="go back" onClick="history.go( -1 );return true;"> 
              &nbsp;</td>
          </tr>
        </table>
</form>






















<%
mrs.close
set mrs=nothing
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
<br>
<br>
</body>
</html>