<SCRIPT LANGUAGE="JavaScript">
<!--

function check_howhform()
{
	if(! isNumber(howhform.oufinalh.value)) {
		alert('Please enter a valid "Home team NO. of score"');
		howhform.oufinalh.focus();
		return false;
		}
	if(! isNumber(howhform.oufinala.value)) {
		alert('Please enter a valid "Away team NO. of score"');
		howhform.oufinala.focus();
		return false;
		}
	if(! isNumber(howhform.hsnum.value)) {
		alert('Please enter a valid "Player NO. of scorer"');
		howhform.hsnum.focus();
		return false;
		}
	if(! isNumber(howhform.hcnum.value)) {
		alert('Please enter a valid "Player NO. of discipline"');
		howhform.hcnum.focus();
		return false;
		}
	if(! isNumber(howhform.asnum.value)) {
		alert('Please enter a valid "Player NO. of scorer"');
		howhform.asnum.focus();
		return false;
		}
	if(! isNumber(howhform.acnum.value)) {
		alert('Please enter a valid "Player NO. of discipline"');
		howhform.acnum.focus();
		return false;
		}
		return true;
}


function checkzhform(hsnum,hcnum)
{
if(hsnum!=2){
j=0;
	for(i = 0; i < hsnum; i++) 
	{
	if(! isNumber(eval("zhform.s" + j + ".value"))) {
		alert('Please enter a valid "Player NO. of scorer"  '+j);
		eval("zhform.s" + j).focus();
		return false;
		}
	if( j == hsnum-1 )
		break;
	j=j+1;
	}
}
if(hsnum==2){
	if(! isNumber(zhform.s0.value)) {
		alert('Please enter a valid "Player NO. of scorer"');
		zhform.s0.focus();
		return false;
		}
	if(! isNumber(zhform.s1.value)) {
		alert('Please enter a valid "Player NO. of scorer"');
		zhform.s1.focus();
		return false;
		}
}
if(hcnum!=2){
j=0;
	for(i = 0; i < hcnum; i++) 
	{
	if(! isNumber(eval("zhform.d" + j + ".value"))) {
		alert('Please enter a valid "Player NO. of discipline"  '+j);
		eval("zhform.d" + j).focus();
		return false;
		}
	if( j == hcnum-1 )
		break;
	j=j+1;
	}
}
if(hcnum==2){
	if(! isNumber(zhform.d0.value)) {
		alert('Please enter a valid "Player NO. of discipline"');
		zhform.d0.focus();
		return false;
		}
	if(! isNumber(zhform.d1.value)) {
		alert('Please enter a valid "Player NO. of discipline"');
		zhform.d1.focus();
		return false;
		}
}
	if(! isNumber(zhform.hscore.value)) {
		alert('Please enter a valid "Home of score" hsnum ='+ hsnum + 'hcnum=' + hcnum );
		zhform.hscore.focus();
		return false;
		}
	if(! isNumber(zhform.ascore.value)) {
		alert('Please enter a valid "Away of score"');
		zhform.ascore.focus();
		return false;
		}
	if(zhform.refer.value.length == 0) {
		alert('Please select a "refer"');
		zhform.refer.focus();
		return false;
		}
return true;
}

function checkzaform(asnum,acnum)
{
if(asnum!=2){
j=0;
	for(i = 0; i < asnum; i++) 
	{
	if(! isNumber(eval("zaform.s" + j + ".value"))) {
		alert('Please enter a valid "Player NO. of scorer"  '+j);
		eval("zaform.s" + j).focus();
		return false;
		}
	if( j == asnum-1 )
		break;
	j=j+1;
	}
}
if(asnum==2){
	if(! isNumber(zaform.s0.value)) {
		alert('Please enter a valid "Player NO. of scorer"');
		zaform.s0.focus();
		return false;
		}
	if(! isNumber(zaform.s1.value)) {
		alert('Please enter a valid "Player NO. of scorer"');
		zaform.s1.focus();
		return false;
		}
}
if(acnum!=2){
j=0;
	for(i = 0; i < acnum; i++) 
	{
	if(! isNumber(eval("zaform.d" + j + ".value"))) {
		alert('Please enter a valid "Player NO. of discipline"  '+j);
		eval("zaform.d" + j).focus();
		return false;
		}
	if( j == acnum-1 )
		break;
	j=j+1;
	}
}
if(acnum==2){
	if(! isNumber(zaform.d0.value)) {
		alert('Please enter a valid "Player NO. of discipline"');
		zaform.d0.focus();
		return false;
		}
	if(! isNumber(zaform.d1.value)) {
		alert('Please enter a valid "Player NO. of discipline"');
		zaform.d1.focus();
		return false;
		}
}
	if(! isNumber(zaform.hscore.value)) {
		alert('Please enter a valid "Home of score" asnum ='+ asnum + 'acnum=' + acnum );
		zaform.hscore.focus();
		return false;
		}
	if(! isNumber(zaform.ascore.value)) {
		alert('Please enter a valid "Away of score"');
		zaform.ascore.focus();
		return false;
		}
	if(zaform.refer.value.length == 0) {
		alert('Please select a "refer"');
		zaform.refer.focus();
		return false;
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
if request("option")="editresultissucc" then response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Update match result is complete ! ');</SCRIPT>")
tid=session("manager_tid")
mmid=request("mid")
if tid="" and mmid="" then
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Please select the team you want to add scorer or discipline at first ! ');</SCRIPT>")
response.End()
end if
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
oufinalh=request("oufinalh")
oufinala=request("oufinala")
asnum=request("asnum")
acnum=request("acnum")
hsnum=request("hsnum")
hcnum=request("hcnum")
%>
<table width="100%" border="0" cellspacing="2" cellpadding="2">
  <tr> 
    <td>&nbsp;</td>
    <td><table width="100%" border="0" cellspacing="2" cellpadding="2">
        <tr> 
          <td height="10">&nbsp;</td>
        </tr>
        <tr> 
          <td height="30" align="center" valign="middle"><h2><%=mrs("htname")%> <font color="#FF0000"> VS </font> <%=mrs("atname")%></h2></td>
        </tr>
      </table></td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td> <table width="100%" border="1" cellpadding="3" cellspacing="0" bordercolor="#CCCCCC">
        <tr> 
          <td width="19%"><b>Week:</b></td>
          <td width="30%"><%=mrs("mweek")%></td>
          <td width="51%">&nbsp;</td>
        </tr>
        <tr> 
          <td width="19%"><b>Date:</b></td>
          <td width="30%"><%=mrs("mmonth")&"/"&mrs("mday")&"/"&mrs("myear")%></td>
          <td width="51%">&nbsp;</td>
        </tr>
        <tr> 
          <td width="19%"><b>KO Time:</b></td>
          <td width="30%"><%=mrs("mhour")&":"&mrs("mminute")%></td>
          <td width="51%">&nbsp;</td>
        </tr>
        <tr> 
          <td width="19%"><b>Pitch:</b></td>
          <td width="30%"><%=mrs("fname")%></td>
          <td width="51%">&nbsp;</td>
        </tr>
        <%
sql="select * from rresult where mid=" & mmid &" and tid=" & mrs("htid")
Set rs=Server.CreateObject("ADODB.RecordSet")
rs.Open sql,conn,1,1
if not rs.eof then
hscore=rs("hscore")
ascore=rs("ascore")
end if
rs.close
%>
        <tr> 
          <td width="19%"><b>Home Team Reported:</b></td>
          <td width="30%"><font color="#FF0000"><%=mrs("htname")%></font></td>
          <td width="51%"><font color="#FF0000"><%=hscore&" : "&ascore%></font></td>
        </tr>
        <%
sql="select * from rresult where mid=" & mmid &" and tid=" & mrs("atid")
rs.Open sql,conn,1,1
if not rs.eof then
hscore=rs("hscore")
ascore=rs("ascore")
hrefer=rs("refer")
end if
rs.close
%>
        <tr> 
          <td width="19%"><b>Away Team Reported:</b></td>
          <td width="30%"><font color="#FF0000"><%=mrs("atname")%></font></td>
          <td width="51%"><font color="#FF0000"><%=hscore&" : "&ascore%></font></td>
        </tr>
      </table></td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td height="60" align="center"><h2><br>
        <br>
        Set Final Score</h2></td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td>
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
hdasnum=0
hdacnum=0

i=0
sql="select * from scorer where mid=" & mmid &" and tid=" & mrs("htid")
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
sql="select * from card where mid=" & mmid &" and tid=" & mrs("htid")
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
if hsnum="" and hcnum="" then
hsnum=hdhsnum
hcnum=hdhcnum
end if
%>









<%
i=0
sql="select * from scorer where mid=" & mmid &" and tid=" & mrs("atid")
Set rs=Server.CreateObject("ADODB.RecordSet")
rs.Open sql,conn,1,1
if not rs.eof then
hdasnum=rs.recordcount
end if
rs.close

i=0
sql="select * from card where mid=" & mmid &" and tid=" & mrs("atid")
rs.Open sql,conn,1,1
if not rs.eof then
hdacnum=rs.recordcount
end if
rs.close
%>
<%
if asnum="" and acnum="" then
asnum=hdasnum
acnum=hdacnum
end if
%> 
















	 <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td><form id="howhform" action="match_r_edeff.asp" onSubmit="return check_howhform();" method="post">
              <input type="hidden" id="mid" name="mid" value=<%=mmid%>>
              <table width="100%" border="0" cellspacing="2" cellpadding="2">
                <tr> 
                  <td width="33%"><font color="#FF0000"><strong>STEP 1 </strong></font></td>
                  <td width="50%" align="center">&nbsp;</td>
                  <td width="33%">&nbsp;</td>
                </tr>
              </table>
              <table width="100%" border="1" cellpadding="3" cellspacing="0" bordercolor="#CCCCCC">
                <tr align="center"> 
                  <td height="30" valign="middle"><strong>ITEMS</strong></td>
                  <td valign="middle"><strong>HOME TEAM</strong></td>
                  <td height="30" valign="middle"><strong>AWAY TEAM</strong></td>
                </tr>
                <tr> 
                  <td height="24">TEAM NAME :</td>
                  <td height="24" align="center"><font color="#FF0000"><%=mrs("htname")%></font></td>
                  <td height="24" align="center"><font color="#FF0000"><%=mrs("atname")%></font></td>
                </tr>
                <tr> 
                  <td width="26%">SCORE :</td>
                  <td align="center"><input name="oufinalh" type="text" id="oufinalh" size="5" maxlength="2" value=<%=oufinalh%>></td>
                  <td align="center"><input name="oufinala" type="text" id="oufinala" size="5" maxlength="2" value=<%=oufinala%>> 
                  </td>
                </tr>
                <tr> 
                  <td width="26%">How many player kicked in :</td>
                  <td align="center"><input name="hsnum2" type="text" id="hsnum2" size="5" maxlength="2" value=<%=hsnum%>></td>
                  <td align="center"><input name="asnum2" type="text" id="asnum2" size="5" maxlength="2" value=<%=asnum%>> 
                  </td>
                </tr>
                <tr> 
                  <td>How many player disciplined :</td>
                  <td align="center"><input name="hcnum2" id="hcnum2" type="text" size="5" maxlength="2" value=<%=hcnum%>></td>
                  <td align="center"><input name="acnum2" id="acnum2" type="text" size="5" maxlength="2" value=<%=acnum%>> 
                  </td>
                </tr>
                <tr> 
                  <td height="50" colspan="3"><input type="submit" name="option" id="option" value=" Click &amp; Set Step 2 -&gt; "></td>
                </tr>
              </table>
            </form>
            &nbsp;</td>
        </tr>
        <tr> 
          <td> </td>
        </tr>
        <tr> 
          <td><br>
            <br>
            <table width="100%" border="0" cellspacing="2" cellpadding="2">
              <tr> 
                <td width="28%"><font color="#FF0000"><strong>STEP 2</strong></font></td>
              </tr>
            </table></td>
        </tr>
      </table>
      
    </td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td> <form id="zhform" action="match_r_savedeff.asp" method="post">
              <input type="hidden" id="hmid" name="hmid" value=<%=mmid%>>
              <input type="hidden" id="htid" name="htid" value=<%=mrs("htid")%>>
              <input type="hidden" id="htname" name="htname" value=<%=mrs("htname")%>>
              <input name="hcnum" id="hcnum" type="hidden" size="5" maxlength="2" value=<%=hcnum%>>
              <input name="hsnum" type="hidden" id="hsnum" size="5" maxlength="2" value=<%=hsnum%>>
              <input type="hidden" id="mid" name="mid" value=<%=mmid%>>
              <input type="hidden" id="tid" name="tid" value=<%=mrs("atid")%>>
              <input type="hidden" id="tname" name="tname" value=<%=mrs("atname")%>>
              <input name="acnum" id="acnum" type="hidden" size="5" maxlength="2" value=<%=acnum%>>
              <input name="asnum" type="hidden" id="asnum" size="5" maxlength="2" value=<%=asnum%>>
              <table width="100%" border="1" cellpadding="3" cellspacing="0" bordercolor="#CCCCCC">
                <tr> 
                  <td colspan="4"><font color="#FF0000"><b>For <%=mrs("htname")%></b></font></td>
                </tr>
                <%
if hsnum>0 then
%>
                <tr> 
                  <td height="30" colspan="4">Score : </td>
                </tr>
                <%
end if
for i=0 to hsnum-1
if i<= ubound(aspid) then
%>
                <tr> 
                  <td width="10%">Player : </td>
                  <td width="25%"> <select name="<%="hspid"&cstr(i)%>" id="<%="hspid"&cstr(i)%>">
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
psql="select * from player where pid<>" & aspid(i) &" and tid=" & mrs("htid")
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
                  <td width="27%"> <input name="<%="hs"&cstr(i)%>" id="<%="hs"&cstr(i)%>" type="text" size="5" maxlength="2" value="<%=asi(i)%>">
                    -score</td>
                  <td width="38%">&nbsp;</td>
                </tr>
                <%
else
%>
                <tr> 
                  <td>Player : </td>
                  <td> <select name="<%="hspid"&cstr(i)%>" id="<%="hspid"&cstr(i)%>">
                      <%
psql="select * from player where tid=" & mrs("htid")
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
                  <td> <input name="<%="hs"&cstr(i)%>" id="<%="hs"&cstr(i)%>" type="text" size="5" maxlength="2">
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
                <%
if hcnum>0 then
%>
                <tr> 
                  <td height="30" colspan="4">discipline: </td>
                </tr>
                <%
end if
for i=0 to hcnum-1
if i<=ubound(adpid) then
%>
                <tr> 
                  <td>Player : </td>
                  <td> <select name="<%="hdpid"&cstr(i)%>" id="<%="hdpid"&cstr(i)%>">
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
psql="select * from player where pid<>" & adpid(i) &" and tid=" & mrs("htid")
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
                  <td><select name="<%="hc"&cstr(i)%>" id="<%="hc"&cstr(i)%>">
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
                  <td> <input name="<%="hd"&cstr(i)%>" type="text" id="<%="hd"&cstr(i)%>" size="5" maxlength="2" value=<%=adi(i)%>>
                    -NUM</td>
                </tr>
                <%
else
%>
                <tr> 
                  <td>Player : </td>
                  <td> <select name="<%="hdpid"&cstr(i)%>" id="<%="hdpid"&cstr(i)%>">
                      <%
psql="select * from player where tid=" & mrs("htid")
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
                  <td><select name="<%="hc"&cstr(i)%>" id="<%="hc"&cstr(i)%>">
                      <option value="yellow" selected>yellow</option>
                      <option value="red">red</option>
                    </select>
                    -Card </td>
                  <td> <input name="<%="hd"&cstr(i)%>" type="text" id="<%="hd"&cstr(i)%>" size="5" maxlength="2">
                    -NUM</td>
                </tr>
                <%
end if
next
%>
                <tr> 
                  <td height="50" colspan="4">&nbsp;</td>
                </tr>
                <%
i=0
sql="select * from scorer where mid=" & mmid &" and tid=" & mrs("atid")
Set rs=Server.CreateObject("ADODB.RecordSet")
rs.Open sql,conn,1,1
if not rs.eof then
hdasnum=rs.recordcount
redim asi(hdasnum-1)
redim aspid(hdasnum-1)
do while not rs.eof
asi(i)=rs("score")
aspid(i)=rs("pid")
i=i+1
rs.movenext
loop
end if
rs.close

i=0
sql="select * from card where mid=" & mmid &" and tid=" & mrs("atid")
rs.Open sql,conn,1,1
if not rs.eof then
hdacnum=rs.recordcount
redim adi(hdacnum-1)
redim adpid(hdacnum-1)
redim aci(hdacnum-1)
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
                <tr> 
                  <td colspan="4"><font color="#FF0000"><b>For <%=mrs("atname")%></b></font></td>
                </tr>
                <%
if asnum>0 then
%>
                <tr> 
                  <td height="30">Score : </td>
                  <td>&nbsp;</td>
                  <td>&nbsp;</td>
                  <td>&nbsp;</td>
                </tr>
                <%
end if
for i=0 to asnum-1
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
psql="select * from player where pid<>" & aspid(i) &" and tid=" & mrs("atid")
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
psql="select * from player where tid=" & mrs("atid")
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
                <%
if acnum>0 then
%>
                <tr> 
                  <td height="30" colspan="4">discipline: </td>
                </tr>
                <%
end if
for i=0 to acnum-1
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
psql="select * from player where pid<>" & adpid(i) &" and tid=" & mrs("atid")
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
psql="select * from player where tid=" & mrs("atid")
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
                  <td colspan="4"><br>
                    <br>
                    Referee:</td>
                </tr>
                <tr>
                  <td colspan="4">Ratings
<select name="hrefer" id="hrefer">
                      <option value="<%=hrefer%>" selected><%=hrefer%></option>
                      <%
ourefer=1
for i=1 to 10
if ourefer<>hrefer then
%>
                      <option value="<%=ourefer%>"><%=ourefer%></option>
                      <%
end if
ourefer=ourefer+1
next
%>
                    </select></td>
                </tr>
                <tr> 
                  <td colspan="4"> <input type="submit" id="option" name="option" value="update"> 
                    <input type="button" name="goback" value="go back" onClick="history.go( -1 );return true;"> 
                    &nbsp;</td>
                </tr>
                <tr> 
                  <td colspan="4">&nbsp;</td>
                </tr>
              </table>
            </form>
            &nbsp;</td>
        </tr>
      </table></td>
    <td>&nbsp;</td>
  </tr>
</table>
<%
mrs.close
set mrs=nothing
%>
<br>
<br>
</body>
</html>