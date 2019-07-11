
<!--#include file="adminconn.asp" -->
<%
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Create standing is complete. ');</SCRIPT>")
if request("option")="deledit" then response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' The matchs result is deleted , now you can add a now one ! ');</SCRIPT>")

msql="select * from match order by mid desc"
Set mrs=Server.CreateObject("ADODB.RecordSet")
mrs.Open msql,conn,1,1
mmid=mrs("mid")
%>
<%
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
    <td> <table width="100%" border="0" cellspacing="2" cellpadding="2">
        <tr> 
          <td><form id="howhform" action="match_r_e.asp" onSubmit="return check_howhform();" method="post">
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
                  <td width="46%"><input name="hsnum" type="text" id="hsnum" size="5" maxlength="2" value=<%=hsnum%>>
                    --home</td>
                  <td width="3%">&nbsp;</td>
                </tr>
                <tr> 
                  <td>How many player discipline :</td>
                  <td><input name="hcnum" id="hcnum" type="text" size="5" maxlength="2" value=<%=hcnum%>>
                    --home</td>
                  <td>&nbsp;</td>
                </tr>
                <tr> 
                  <td width="51%">How many player scorer :</td>
                  <td width="46%"><input name="asnum" type="text" id="asnum" size="5" maxlength="2" value=<%=asnum%>>
                    --away</td>
                  <td width="3%">&nbsp;</td>
                </tr>
                <tr> 
                  <td>How many player discipline :</td>
                  <td><input name="acnum" id="acnum" type="text" size="5" maxlength="2" value=<%=acnum%>>
                    --away</td>
                  <td>&nbsp;</td>
                </tr>
                <tr> 
                  <td height="50" colspan="3"><input type="submit" name="option" id="option" value=" STEP 2 -> "></td>
                </tr>
              </table>
            </form>
            &nbsp;</td>
        </tr>
        <tr> 
          <td> </td>
        </tr>
        <tr> 
          <td>&nbsp;</td>
        </tr>
      </table></td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td><table width="100%" border="0" cellspacing="2" cellpadding="2">
        <tr> 
          <td> <form id="zhform" action="match_r_save.asp" method="post">
              <input type="hidden" id="hmid" name="hmid" value=<%=mmid%>>
              <input type="hidden" id="htid" name="htid" value=<%=mrs("htid")%>>
              <input type="hidden" id="htname" name="htname" value=<%=mrs("htname")%>>
              <input name="hcnum" id="hcnum" type="hidden" size="5" maxlength="2" value=<%=hcnum%>>
              <input name="hsnum" type="hidden" id="hsnum" size="5" maxlength="2" value=<%=hsnum%>>
              <input type="hidden" id="mid" name="mid" value=<%=mmid%>>
              <input type="hidden" id="tid" name="tid" value=<%=mrs("atid")%>>
              <input type="hidden" id="tname" name="tname" value=<%=mrs("atname")%>>
              <input type="hidden" id="tstr" name="tstr" value="a">
              <input name="acnum" id="acnum" type="hidden" size="5" maxlength="2" value=<%=acnum%>>
              <input name="asnum" type="hidden" id="asnum" size="5" maxlength="2" value=<%=asnum%>>
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
                  <td colspan="4"><font color="#FF0000">For <%=mrs("htname")%></font></td>
                </tr>
                <tr> 
                  <td width="13%">&nbsp;</td>
                  <td width="10%">&nbsp;</td>
                  <td width="26%">&nbsp;</td>
                  <td width="51%">&nbsp;</td>
                </tr>
                <tr> 
                  <td height="30">Score : </td>
                  <td>&nbsp;</td>
                  <td>&nbsp;</td>
                  <td>&nbsp;</td>
                </tr>
                <%
for i=0 to hsnum-1
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
next
%>
                <tr> 
                  <td colspan="4">&nbsp;</td>
                </tr>
                <tr> 
                  <td height="30">discipline: </td>
                  <td>&nbsp;</td>
                  <td>&nbsp;</td>
                  <td>&nbsp;</td>
                </tr>
                <%
for i=0 to hcnum-1
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
next
%>
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
                  <td height="30">Score : </td>
                  <td>&nbsp;</td>
                  <td>&nbsp;</td>
                  <td>&nbsp;</td>
                </tr>
                <%
for i=0 to asnum-1
%>
                <tr> 
                  <td height="28">Player : </td>
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
next
%>
                <tr> 
                  <td colspan="4">&nbsp;</td>
                </tr>
                <tr> 
                  <td height="30">discipline: </td>
                  <td>&nbsp;</td>
                  <td>&nbsp;</td>
                  <td>&nbsp;</td>
                </tr>
                <%
for i=0 to acnum-1
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
