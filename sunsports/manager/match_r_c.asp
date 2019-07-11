<SCRIPT LANGUAGE="JavaScript">
<!--
var z001
function checkzform1()
{
z001="no";
return true;
}
function checkzform2()
{
z001="yes";
return true;
}

function checkzform()
{
if((zform.snum.value.length>0 || zform.dnum.value.length>0) && z001=="yes"){
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
	if(! isNumber(zform.snum.value)) {
		alert('Please enter a valid "Player NO. of scorer"');
		zform.snum.focus();
		return false;
		}
	if(! isNumber(zform.dnum.value)) {
		alert('Please enter a valid "Player NO. of discipline"');
		zform.dnum.focus();
		return false;
		}
if(zform.snum.value>0 && z001=="yes" && zform.snum.value!=2){
j=0;
	for(i = 0; i < zform.snum.value; i++) 
	{
	if(! isNumber(eval("zform.s" + j + ".value"))) {
		alert('Please enter a valid "Player NO. of scorer"  '+j);
		eval("zform.s" + j).focus();
		return false;
		}
	if( j == zform.snum.value-1 )
		break;
	j=j+1;
	}
}
if(zform.snum.value==2 && z001=="yes"){
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
if(zform.dnum.value>0 && z001=="yes" && zform.dnum.value!=2){
j=0;
	for(i = 0; i < zform.dnum.value; i++) 
	{
	if(! isNumber(eval("zform.d" + j + ".value"))) {
		alert('Please enter a valid "Player NO. of discipline"  '+j);
		eval("zform.d" + j).focus();
		return false;
		}
	if( j == zform.dnum.value-1 )
		break;
	j=j+1;
	}
}
if(zform.dnum.value==2 && z001=="yes"){
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
if request("option")="add" then






if snum>0 then
dim asi()
dim aspid()
redim asi(snum-1)
redim aspid(snum-1)
for i=0 to ubound(asi)
aspid(i)=request("spid"&cstr(i))
asi(i)=request("s"&cstr(i))
ssql="select * from scorer where mid="&mmid&" and pid="&aspid(i)
Set srs=Server.CreateObject("ADODB.RecordSet") 
srs.Open ssql,conn,1,1 

if not ( srs.eof or srs.bof ) then
tfstr="yes"
else
tfstr="no"
end if

srs.close
set srs=nothing

if tfstr="yes" then 
conn.Execute("update scorer set score="&asi(i)&" where mid="&mmid&" and pid="&aspid(i))
end if

if tfstr="no" then
psql="select * from player where pid=" & aspid(i)
Set prs=Server.CreateObject("ADODB.RecordSet") 
prs.Open psql,conn,1,1 
pname=prs("name")
jersey=prs("jersey")
prs.close
set prs=nothing
conn.Execute("insert into scorer (mid,tid,tname,pid,pname,jersey,score) values ("&mmid&","&session("manager_tid")&",'"&session("manager_tname")&"',"&aspid(i)&",'"&pname&"',"&jersey&","&asi(i)&")")
end if
next
erase aspid
erase asi
end if




if dnum>0 then
dim adi()
dim adpid()
dim aci()
redim adi(dnum-1)
redim adpid(dnum-1)
redim aci(dnum-1)
for i=0 to ubound(adpid)
adpid(i)=request("dpid"&cstr(i))
aci(i)=request("c"&cstr(i))
adi(i)=request("d"&cstr(i))
ssql="select * from card where mid="&mmid&" and pid="&adpid(i)
Set srs=Server.CreateObject("ADODB.RecordSet") 
srs.Open ssql,conn,1,1 

if not ( srs.eof or srs.bof ) then
tfstr="yes"
else
tfstr="no"
end if

srs.close
set srs=nothing

if tfstr="yes" then
if aci(i)="yellow" then conn.Execute("update card set yellow="&adi(i)&" , red=0 where mid="&mmid&" and pid="&adpid(i))
if aci(i)="red" then conn.Execute("update card set red="&adi(i)&" , yellow=0 where mid="&mmid&" and pid="&adpid(i))
end if

if tfstr="no" then
psql="select * from player where pid=" & adpid(i)
Set prs=Server.CreateObject("ADODB.RecordSet") 
prs.Open psql,conn,1,1 
pname=prs("name")
jersey=prs("jersey")
prs.close
set prs=nothing
if aci(i)="yellow" then conn.Execute("insert into card (mid,tid,tname,pid,pname,jersey,yellow,red) values ("&mmid&","&session("manager_tid")&",'"&session("manager_tname")&"',"&adpid(i)&",'"&pname&"',"&jersey&","&adi(i)&",0)")
if aci(i)="red" then conn.Execute("insert into card (mid,tid,tname,pid,pname,jersey,red,yellow) values ("&mmid&","&session("manager_tid")&",'"&session("manager_tname")&"',"&adpid(i)&",'"&pname&"',"&jersey&","&adi(i)&",0)")
end if
next
erase adpid
erase aci
erase adi
end if


acountscore=0
sql="select * from scorer where mid="&mmid&" and tid="&tid
rs.Open sql,conn,1,1
do while not rs.eof 
acountscore=acountscore + rs("score")
rs.movenext
loop
rs.close
set rs=nothing

if request("homeaway")="h" then hscore=acountscore
if request("homeaway")="a" then ascore=acountscore

tsql="select * from rresult where mid="&mmid&" and tid="&tid
Set trs=Server.CreateObject("ADODB.RecordSet") 
trs.Open tsql,conn,1,1 
if trs.eof or trs.bof then
conn.Execute("insert into rresult (mid,tid,hscore,ascore,refer) values ("&mmid&","&tid&","&hscore&","&ascore&","&refer&")")
else
conn.Execute( "update rresult set hscore="&hscore&", ascore="&ascore&" where mid="&mmid&" and tid="&tid )
end if
trs.close
set trs=nothing


response.Redirect("match_r_e.asp?option=createresultissucc&mid="&mmid)
end if
%>
<form id="zform" action="match_r_c.asp" onSubmit="return checkzform();" method="post">
              <input type="hidden" id="mid" name="mid" value=<%=mmid%>> 
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
                  <td height="50" colspan="3"><input type="submit" name="option" id="option" value=" STEP 2 -> " onClick="return checkzform1();"></td>
                </tr>
              </table>
			</td>
        </tr>
      </table></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>
<%
if snum<>"" and dnum<>"" then
%>
<%
msql="select * from match where mid="&mmid
Set mrs=Server.CreateObject("ADODB.RecordSet")
mrs.Open msql,conn,1,1
%>
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
            <td width="13%">&nbsp;</td>
            <td width="17%">&nbsp;</td>
            <td width="22%">&nbsp;</td>
            <td width="48%">&nbsp;</td>
          </tr>
          <tr> 
            <td width="13%"><b>Week</b></td>
            <td width="17%"><%=mrs("mweek")%></td>
            <td width="22%">&nbsp;</td>
            <td width="48%">&nbsp;</td>
          </tr>
          <tr> 
            <td width="13%"><b>Date</b></td>
            <td width="17%"><%=mrs("mmonth")&"/"&mrs("mday")&"/"&mrs("myear")%></td>
            <td width="22%">&nbsp;</td>
            <td width="48%">&nbsp;</td>
          </tr>
          <tr> 
            <td width="13%"><b>Time</b></td>
            <td width="17%"><%=mrs("mhour")&":"&mrs("mminute")%></td>
            <td width="22%">&nbsp;</td>
            <td width="48%">&nbsp;</td>
          </tr>
          <tr> 
            <td width="13%"><b>Home Team</b></td>
            <td width="17%"><%=mrs("htname")%></td>
            <td width="22%"><input name="hscore" type="text" id="hscore" size="5" maxlength="2" value=<%=hscore%>> 
              <b>Score</b></td>
            <td width="48%">&nbsp;</td>
          </tr>
          <tr> 
            <td width="13%"><b>Away Team </b></td>
            <td width="17%"><%=mrs("atname")%></td>
            <td width="22%"><input name="ascore" type="text" id="ascore2" size="5" maxlength="2" value=<%=ascore%>> 
              <b>Score</b></td>
            <td width="48%">&nbsp;</td>
          </tr>
          <tr> 
            <td width="13%"><b>Location</b></td>
            <td width="17%"><%=mrs("fname")%></td>
            <td width="22%">&nbsp;</td>
            <td width="48%">&nbsp;</td>
          </tr>
          <tr> 
            <td width="13%">&nbsp;</td>
            <td width="17%">&nbsp;</td>
            <td width="22%">&nbsp;</td>
            <td width="48%">&nbsp;</td>
          </tr>
          <tr> 
            <td width="13%">&nbsp;</td>
            <td width="17%">&nbsp;</td>
            <td width="22%">&nbsp;</td>
            <td width="48%">&nbsp;</td>
          </tr>
          <tr> 
            <td width="13%">&nbsp;</td>
            <td width="17%">&nbsp;</td>
            <td width="22%">&nbsp;</td>
            <td width="48%">&nbsp;</td>
          </tr>
          <%
if snum>0 or dnum>0 then
%>
          <tr> 
            <td colspan="4"><font color="#FF0000">For <%=session("manager_tname")%></font></td>
          </tr>
          <tr> 
            <td width="13%">&nbsp;</td>
            <td width="17%">&nbsp;</td>
            <td width="22%">&nbsp;</td>
            <td width="48%">&nbsp;</td>
          </tr>
          <%
end if
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
                <option value=<%=prs("pid")%> selected><%=prs("name")%></option>
                <%
prs.movenext
loop
prs.close
set prs=nothing 
%>
              </select> </td>
            <td> <input name="<%="s"&cstr(i)%>" id="<%="s"&cstr(i)%>" type="text" size="5" maxlength="2" >
              -score</td>
            <td>&nbsp;</td>
          </tr>
          <%
next
%>
          <tr> 
            <td colspan="4">&nbsp;</td>
          </tr>
          <tr> 
            <td width="13%">&nbsp;</td>
            <td width="17%">&nbsp;</td>
            <td width="22%">&nbsp;</td>
            <td width="48%">&nbsp;</td>
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
                <option value=<%=prs("pid")%> selected><%=prs("name")%></option>
                <%
prs.movenext
loop
prs.close
set prs=nothing 
%>
              </select> </td>
            <td><select name="<%="c"&cstr(i)%>" id="<%="c"&cstr(i)%>">
                <option value="red">red</option>
                <option value="yellow" selected>yellow</option>
              </select>
              -Card </td>
            <td> <select name="<%="d"&cstr(i)%>" id="<%="d"&cstr(i)%>">
                <option value="1" selected>1</option>
                <option value="2">2</option>
              </select>
              -NUM</td>
          </tr>
          <%
next
%>
          <tr> 
            <td width="13%">&nbsp;</td>
            <td width="17%">&nbsp;</td>
            <td width="22%">&nbsp;</td>
            <td width="48%">&nbsp;</td>
          </tr>
          <tr> 
            <td width="13%">&nbsp;</td>
            <td width="17%">&nbsp;</td>
            <td width="22%">&nbsp;</td>
            <td width="48%">&nbsp;</td>
          </tr>
          <tr> 
            <td width="13%">&nbsp;</td>
            <td width="17%">&nbsp;</td>
            <td width="22%">&nbsp;</td>
            <td width="48%">&nbsp;</td>
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
            <td colspan="4">Yellow Card : <input name="anotheryellow" id="anotheryellow" type="text" size="5" maxlength="2">
             &nbsp;&nbsp;&nbsp; Red Card : <input name="anotherred" id="anotherred" type="text" size="5" maxlength="2"></td>
          </tr>
          <tr> 
            <td colspan="4">&nbsp;</td>
          </tr>
          <tr> 
            <td width="13%">Refer :&nbsp;</td>
            <td width="17%"><select name="refer" id="select2">
                <option value="" selected>SELECT REFER ...</option>
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
              </select></td>
            <td width="22%">&nbsp;</td>
            <td width="48%">&nbsp;</td>
          </tr>
          <tr> 
            <td colspan="4">&nbsp;</td>
          </tr>
          <tr> 
            <td colspan="4"> 
			<input type="hidden" id="homeaway" name="homeaway" value="<%=homeaway%>">
			<input type="submit" id="option" name="option" value="add" onClick="return checkzform2();"> 
              <input type="button" name="goback" value="go back" onClick="history.go( -1 );return true;"> 
              &nbsp;</td>
          </tr>
        </table>
<%
mrs.close
set mrs=nothing
%>
<%
end if
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
</form>
<br>
<br>
</body>
</html>