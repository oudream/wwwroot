<SCRIPT LANGUAGE="JavaScript">
<!--

function check_howhform()
{
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
asnum=request("asnum")
acnum=request("acnum")
hsnum=request("hsnum")
hcnum=request("hcnum")
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
<table width="100%" border="0" cellspacing="2" cellpadding="2">
  <tr> 
    <td>&nbsp;</td>
    <td><table width="100%" border="0" cellspacing="2" cellpadding="2">
        <tr> 
          <td height="10">&nbsp;</td>
        </tr>
        <tr> 
          <td height="30" align="center" valign="middle"><strong>Team <font color="#FF0000"><%=session("manager_tname")%></font> Match Result Create</strong></td>
        </tr>
      </table></td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td> <table width="100%">
        <tr> 
          <td width="13%">&nbsp;</td>
          <td width="10%">&nbsp;</td>
          <td width="26%">&nbsp;</td>
          <td width="51%">&nbsp;</td>
        </tr>
        <tr> 
          <td width="13%"><b>Week</b></td>
          <td width="10%"><%=mrs("mweek")%></td>
          <td width="26%">&nbsp;</td>
          <td width="51%">&nbsp;</td>
        </tr>
        <tr> 
          <td width="13%"><b>Date</b></td>
          <td width="10%"><%=mrs("mmonth")&"/"&mrs("mday")&"/"&mrs("myear")%></td>
          <td width="26%">&nbsp;</td>
          <td width="51%">&nbsp;</td>
        </tr>
        <tr> 
          <td width="13%"><b>Time</b></td>
          <td width="10%"><%=mrs("mhour")&":"&mrs("mminute")%></td>
          <td width="26%">&nbsp;</td>
          <td width="51%">&nbsp;</td>
        </tr>
        <tr> 
          <td width="13%"><b>Home Team</b></td>
          <td width="10%"><%=mrs("htname")%></td>
          <td width="26%">&nbsp;</td>
          <td width="51%">&nbsp;</td>
        </tr>
        <tr> 
          <td width="13%"><b>Away Team </b></td>
          <td width="10%"><%=mrs("atname")%></td>
          <td width="26%">&nbsp;</td>
          <td width="51%">&nbsp;</td>
        </tr>
        <tr> 
          <td width="13%"><b>Location</b></td>
          <td width="10%"><%=mrs("fname")%></td>
          <td width="26%">&nbsp;</td>
          <td width="51%">&nbsp;</td>
        </tr>
        <tr> 
          <td width="13%">&nbsp;</td>
          <td width="10%">&nbsp;</td>
          <td width="26%">&nbsp;</td>
          <td width="51%">&nbsp;</td>
        </tr>
      </table></td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td align="center">&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td>
<table width="100%" border="0" cellspacing="2" cellpadding="2">
        <tr>
          <td align="center">For home</td>
        </tr>
        <tr>
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

i=0
sql="select * from scorer where mid=" & mmid &" and tid=" & mrs("htid")
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


sql="select * from rresult where mid=" & mmid &" and tid=" & mrs("htid")
rs.Open sql,conn,1,1
if not rs.eof then
hscore=rs("hscore")
ascore=rs("ascore")
end if
rs.close
set rs=nothing
%>
<%
if hsnum="" and hcnum="" then
hsnum=hdhsnum
hcnum=hdhcnum
end if
%>


	  <form id="howhform" action="match_r_e.asp" onSubmit="return check_howhform();" method="post">
              <input type="hidden" id="mid" name="mid" value=<%=mmid%>>
              <table width="100%" border="0" cellspacing="2" cellpadding="2">
                <tr align="center" valign="middle"> 
                  <td height="30" colspan="3"><table width="100%" border="0" cellspacing="2" cellpadding="2">
                      <tr> 
                        <td width="33%"><font color="#FF0000">STEP 1</font></td>
                        <td width="50%" align="center">&nbsp;</td>
                        <td width="33%">&nbsp;</td>
                      </tr>
                    </table></td>
                </tr>
                <tr> 
                  <td height="24" colspan="3">&nbsp; </td>
                </tr>
                <tr> 
                  <td width="51%">How many player scorer :</td>
                  <td width="46%"><input name="hsnum" type="text" id="hsnum" size="5" maxlength="2" value=<%=hsnum%>></td>
                  <td width="3%">&nbsp;</td>
                </tr>
                <tr> 
                  <td>How many player discipline :</td>
                  <td><input name="hcnum" id="hcnum" type="text" size="5" maxlength="2" value=<%=hcnum%>></td>
                  <td>&nbsp;</td>
                </tr>
                <tr> 
                  <td height="50" colspan="3"><input type="submit" name="option" id="option" value=" STEP 2 -> "></td>
                </tr>
              </table>
            </form></td>
        </tr>
        <tr>
          <td>&nbsp;</td>
        </tr>
        <tr>
          <td>
<form id="zhform" action="match_r_save.asp" onSubmit="return checkzhform(<%=hsnum%>,<%=hcnum%>);" method="post">
<input type="hidden" id="mid" name="mid" value=<%=mmid%>>
<input type="hidden" id="tid" name="tid" value=<%=mrs("htid")%>>
<input type="hidden" id="tname" name="tname" value=<%=mrs("htname")%>>
<input type="hidden" id="tstr" name="tstr" value="h">
<input name="acnum" id="acnum" type="hidden" size="5" maxlength="2" value=<%=hcnum%>>
<input name="hsnum" type="hidden" id="hsnum" size="5" maxlength="2" value=<%=hsnum%>>
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
                  <td width="13%">&nbsp; </td>
                  <td width="10%">&nbsp;</td>
                  <td width="26%">&nbsp;</td>
                  <td width="51%">&nbsp;</td>
                </tr>
                <tr> 
                  <td width="13%">&nbsp;</td>
                  <td width="10%">&nbsp;</td>
                  <td width="26%">&nbsp;</td>
                  <td width="51%">&nbsp;</td>
                </tr>
                <tr> 
                  <td width="13%">&nbsp;</td>
                  <td width="10%">&nbsp;</td>
                  <td width="26%"><input name="hscore" type="text" id="hscore" size="5" maxlength="2" value=<%=hscore%>> 
                    <b>Score</b></td>
                  <td width="51%"><input name="ascore" type="text" id="ascore" size="5" maxlength="2" value=<%=ascore%>> 
                    <b>Score</b></td>
                </tr>
                <tr> 
                  <td width="13%">&nbsp;</td>
                  <td width="10%">&nbsp;</td>
                  <td width="26%">&nbsp;</td>
                  <td width="51%">&nbsp;</td>
                </tr>
                <tr> 
                  <td colspan="4"><font color="#FF0000">For <%=mrs("htname")%></font></td>
                </tr>
                <tr> 
                  <td width="13%">&nbsp;</td>
                  <td width="10%">&nbsp;</td>
                  <td width="26%">&nbsp;</td>
                  <td width="51%">&nbsp;</td>
                </tr>
<%
if hsnum>0 then
%>
                <tr> 
                  <td height="30">Score : </td>
                  <td>&nbsp;</td>
                  <td>&nbsp;</td>
                  <td>&nbsp;</td>
                </tr>
                <%
end if
for i=0 to hsnum-1
if i<= ubound(aspid) then
%>
                <tr> 
                  <td>Player : </td>
                  <td> 
				  
				  <select name="<%="spid"&cstr(i)%>" id="<%="spid"&cstr(i)%>">
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
                  <td> <input name="<%="s"&cstr(i)%>" id="<%="s"&cstr(i)%>" type="text" size="5" maxlength="2" value="<%=asi(i)%>">
                    -score</td>
                  <td>&nbsp;</td>
                </tr>
<%
else
%>
                <tr> 
                  <td>Player : </td>
                  <td> 
				  
				  <select name="<%="spid"&cstr(i)%>" id="<%="spid"&cstr(i)%>">
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
                  <td width="13%">&nbsp;</td>
                  <td width="10%">&nbsp;</td>
                  <td width="26%">&nbsp;</td>
                  <td width="51%">&nbsp;</td>
                </tr>
<%
if hcnum>0 then
%>
                <tr> 
                  <td height="30">discipline: </td>
                  <td>&nbsp;</td>
                  <td>&nbsp;</td>
                  <td>&nbsp;</td>
                </tr>
                <%
end if
for i=0 to hcnum-1
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
                  <td>
                    <input name="<%="d"&cstr(i)%>" type="text" id="<%="d"&cstr(i)%>" size="5" maxlength="2" value=<%=adi(i)%>>
                    -NUM</td>
                </tr>
<%
else
%>
                <tr> 
                  <td>Player : </td>
                  <td> <select name="<%="dpid"&cstr(i)%>" id="<%="dpid"&cstr(i)%>">
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
                  <td><select name="<%="c"&cstr(i)%>" id="<%="c"&cstr(i)%>">
                      <option value="yellow" selected>yellow</option>
                      <option value="red">red</option>
                    </select>
                    -Card </td>
                  <td>
                    <input name="<%="d"&cstr(i)%>" type="text" id="<%="d"&cstr(i)%>" size="5" maxlength="2">
                    -NUM</td>
                </tr>
<%
end if
next
%>
                <tr> 
                  <td width="13%">&nbsp;</td>
                  <td width="10%">&nbsp;</td>
                  <td width="26%">&nbsp;</td>
                  <td width="51%">&nbsp;</td>
                </tr>
                <tr> 
                  <td colspan="4"> <input type="submit" id="option" name="option" value="add"> 
                    <input type="button" name="goback" value="go back" onClick="history.go( -1 );return true;"> 
                    &nbsp;</td>
                </tr>
              </table>
      </form>
</td>
        </tr>
        <tr>
          <td>&nbsp;</td>
        </tr>
      </table>	 </td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td><table width="100%" border="0" cellspacing="2" cellpadding="2">
        <tr> 
          <td height="30" align="center" valign="middle"> f</td>
        </tr>
      </table></td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td><table width="100%" border="0" cellspacing="2" cellpadding="2">
        <tr> 
          <td align="center">For away</td>
        </tr>
        <tr> 
          <td>
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


sql="select * from rresult where mid=" & mmid &" and tid=" & mrs("atid")
rs.Open sql,conn,1,1
if not rs.eof then
hscore=rs("hscore")
ascore=rs("ascore")
end if
rs.close
set rs=nothing
%> <%
if asnum="" and acnum="" then
asnum=hdasnum
acnum=hdacnum
end if
%> <form id="howhform" action="match_r_e.asp" onSubmit="return check_howhform();" method="post">
              <input type="hidden" id="mid" name="mid" value=<%=mmid%>>
              <table width="100%" border="0" cellspacing="2" cellpadding="2">
                <tr align="center" valign="middle"> 
                  <td height="30" colspan="3"><table width="100%" border="0" cellspacing="2" cellpadding="2">
                      <tr> 
                        <td width="33%"><font color="#FF0000">STEP 1</font></td>
                        <td width="50%" align="center">&nbsp;</td>
                        <td width="33%">&nbsp;</td>
                      </tr>
                    </table></td>
                </tr>
                <tr> 
                  <td height="24" colspan="3">&nbsp; </td>
                </tr>
                <tr> 
                  <td width="51%">How many player scorer :</td>
                  <td width="46%"><input name="asnum" type="text" id="asnum" size="5" maxlength="2" value=<%=asnum%>></td>
                  <td width="3%">&nbsp;</td>
                </tr>
                <tr> 
                  <td>How many player discipline :</td>
                  <td><input name="acnum" id="acnum" type="text" size="5" maxlength="2" value=<%=acnum%>></td>
                  <td>&nbsp;</td>
                </tr>
                <tr> 
                  <td height="50" colspan="3"><input type="submit" name="option" id="option" value=" STEP 2 -> "></td>
                </tr>
              </table>
            </form></td>
        </tr>
        <tr> 
          <td>&nbsp;</td>
        </tr>
        <tr> 
          <td> <form id="zaform" action="match_r_save.asp" onSubmit="return checkzaform(<%=asnum%>,<%=acnum%>);" method="post">
              <input type="hidden" id="mid" name="mid" value=<%=mmid%>>
              <input type="hidden" id="tid" name="tid" value=<%=mrs("atid")%>>
              <input type="hidden" id="tname" name="tname" value=<%=mrs("atname")%>>
              <input type="hidden" id="tstr" name="tstr" value="a">
              <input name="acnum" id="acnum" type="hidden" size="5" maxlength="2" value=<%=acnum%>>
              <input name="hsnum" type="hidden" id="hsnum" size="5" maxlength="2" value=<%=asnum%>>
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
                  <td width="13%">&nbsp; </td>
                  <td width="10%">&nbsp;</td>
                  <td width="26%">&nbsp;</td>
                  <td width="51%">&nbsp;</td>
                </tr>
                <tr> 
                  <td width="13%">&nbsp;</td>
                  <td width="10%">&nbsp;</td>
                  <td width="26%">&nbsp;</td>
                  <td width="51%">&nbsp;</td>
                </tr>
                <tr> 
                  <td width="13%">&nbsp;</td>
                  <td width="10%">&nbsp;</td>
                  <td width="26%"><input name="hscore" type="text" id="hscore" size="5" maxlength="2" value=<%=hscore%>> 
                    <b>Score</b></td>
                  <td width="51%"><input name="ascore" type="text" id="ascore" size="5" maxlength="2" value=<%=ascore%>> 
                    <b>Score</b></td>
                </tr>
                <tr> 
                  <td width="13%">&nbsp;</td>
                  <td width="10%">&nbsp;</td>
                  <td width="26%">&nbsp;</td>
                  <td width="51%">&nbsp;</td>
                </tr>
                <tr> 
                  <td colspan="4"><font color="#FF0000">For <%=mrs("atname")%></font></td>
                </tr>
                <tr> 
                  <td width="13%">&nbsp;</td>
                  <td width="10%">&nbsp;</td>
                  <td width="26%">&nbsp;</td>
                  <td width="51%">&nbsp;</td>
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
                <tr> 
                  <td width="13%">&nbsp;</td>
                  <td width="10%">&nbsp;</td>
                  <td width="26%">&nbsp;</td>
                  <td width="51%">&nbsp;</td>
                </tr>
                <%
if acnum>0 then
%>
                <tr> 
                  <td height="30">discipline: </td>
                  <td>&nbsp;</td>
                  <td>&nbsp;</td>
                  <td>&nbsp;</td>
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
                  <td width="13%">&nbsp;</td>
                  <td width="10%">&nbsp;</td>
                  <td width="26%">&nbsp;</td>
                  <td width="51%">&nbsp;</td>
                </tr>
                <tr> 
                  <td colspan="4"> <input type="submit" id="option" name="option" value="add"> 
                    <input type="button" name="goback" value="go back" onClick="history.go( -1 );return true;"> 
                    &nbsp;</td>
                </tr>
              </table>
            </form></td>
        </tr>
        <tr> 
          <td>&nbsp;</td>
        </tr>
      </table></td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
</table>
      <%
mrs.close
set mrs=nothing
%>
