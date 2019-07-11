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
if(zform.dnum.value>0 && z001=="yes"){
j=0
	for(i = 0; i < zform.dnum.value; i++) {
	if(! isNumber(eval("zform.d" + j + ".value"))) {
		alert('Please enter a valid "Player NO. of discipline"');
		eval("zform.d" + j).focus();
		return false;
		}
if(j==zform.dnum.value-1)
break;
j=j+1
	}
}
if(zform.snum.value>0 && z001=="yes"){
j=0
	for(i = 0; i < zform.snum.value; i++) {
	if(! isNumber(eval("zform.s" + j + ".value"))) {
		alert('Please enter a valid "Player NO. of scorer"');
		eval("zform.s" + j).focus();
		return false;
		}
if(j==zform.snum.value-1)
break;
j=j+1
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


dim adi()
dim adpid()
dim aci()
redim adi(dnum-1)
redim adpid(dnum-1)
redim aci(dnum-1)

for i=0 to ubound(adi)
adi(i)=request("d"&cstr(i))
next

played=trim(request("played"))
tname=request("tname")
if played<>"" and tid<>"" and tname<>"" then
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
sql="insert into standing (tid,tname,played,win,lose,draw,scored,conceded,gd,points) values ("&tid&",'"&tname&"',"&played&","&win&","&lose&","&draw&","&scored&","&conceded&","&gd&","&points&")"
conn.Execute(sql)
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Create standing is complete. ');</SCRIPT>")
end if
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
          <td height="30" align="center" valign="middle"><strong>Match Result 
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
                  <td height="30" colspan="3">Player NO. of scorer and discipline</td>
                </tr>
                <tr> 
                  <td width="29%">Player NO. of scorer :</td>
                  <td width="68%"><input name="snum" type="text" id="snum" size="5" maxlength="2" value=<%=snum%>></td>
                  <td width="3%">&nbsp;</td>
                </tr>
                <tr> 
                  <td>Player NO. of discipline :</td>
                  <td><input name="dnum" id="dnum" type="text" size="5" maxlength="2" value=<%=dnum%>></td>
                  <td>&nbsp;</td>
                </tr>
                <tr> 
                  <td colspan="3"><input type="submit" name="option" id="option" value="Player NO. Submit" onClick="return checkzform1();"></td>
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
	    <table width="100%" border="0" cellspacing="2" cellpadding="2">
          <tr align="center" valign="middle"> 
            <td height="30" colspan="4">Match Result</td>
          </tr>
          <tr> 
            <td width="13%">Scorer:&nbsp;</td>
            <td width="10%">&nbsp;</td>
            <td width="26%">&nbsp;</td>
            <td width="51%">&nbsp;</td>
          </tr>
          <%
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
            <td>discipline: </td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
          </tr>
          <%
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
                <option value="red" selected>red</option>
                <option value="yellow">yellow</option>
              </select>
              -Card </td>
            <td><input name="<%="d"&cstr(i)%>" id="<%="d"&cstr(i)%>" type="text" size="5" maxlength="2" >
              -NUM</td>
          </tr>
          <%
next
%>
          <tr> 
            <td colspan="4">&nbsp;</td>
          </tr>
          <tr> 
            <td colspan="4">
			  <input type="submit" id="option" name="option" value="add" onClick="return checkzform2();"> 
              <input type="button" name="goback" value="go back" onClick="history.go( -1 );return true;"> 
              &nbsp;</td>
          </tr>
        </table>
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
