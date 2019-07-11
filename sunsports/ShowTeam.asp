<!--#include file="Top_sun.asp"-->
<table width="776" height="100%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="181" valign="top"> <br>
      <table width="100%" border="0" cellpadding="3" cellspacing="0">
        <tr> 
          <td>&nbsp;</td>
        </tr>
        <tr> 
          <td> 
            <!-- ========================================================================================================		  

													cup table ============star

 ======================================================================================================== -->
            <table width="170" border="0" align="center" cellpadding="0" cellspacing="3">
              <tr> 
                <td><a href="League.asp"><b><font color="#1F4A71">LEAGUE MATCHES</font></b></a><font color="#1F4A71">+ 
                  +</font></td>
              </tr>
            </table>
            <table width="170" border="0" align="center" cellpadding="0" cellspacing="3">
              <tr> 
                <td><a href="Cup.asp"><b><font color="#1F4A71">CUP MATCHES </font></b></a><font color="#1F4A71">&nbsp;+ 
                  +</font></td>
              </tr>
            </table>
            <table width="170" border="0" align="center" cellpadding="0" cellspacing="3">
              <tr> 
                <td><a href="Friendlies.asp"><b><font color="#1F4A71">FRIENDLY 
                  MATCHES </font></b></a></td>
              </tr>
            </table>
            <table width="170" border="0" align="center" cellpadding="0" cellspacing="3">
              <tr> 
                <td><a href="TeamProfile.asp"><b><font color="#1F4A71">TEAMPORFILE</font></b></a> 
                  <font color="#1F4A71">&nbsp;+ +</font> </td>
              </tr>
            </table>
            <table width="170" border="0" align="center" cellpadding="0" cellspacing="3">
              <tr> 
                <td><a href="TeamProfile.asp"><b><font color="#1F4A71">CHOOSE 
                  A TEAM</font></b></a></td>
              </tr>
            </table>
            <table width="170" border="0" align="center" cellpadding="0" cellspacing="3">
              <tr> 
                <td>
                  <!-- ========================================================================================================		  

													Tournament table ============star

 ======================================================================================================== -->
                  <table width="98%" border="0" cellpadding="0" cellspacing="3">
<%
cid=request("cid")
if cid="" then cid=109
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
response.End()
end if
%>
<%
sql="select * from team where tid=0 "
for i=0 to ubound(ateam)
sql=sql&" or tid="&ateam(i)
next
sql=sql&" order by name"
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.Open sql,conn,1,1 
if rs.bof or rs.eof then response.Redirect("error.asp?err=1")
do while not rs.eof
%>
                    <tr> 
					  <td width="5" valign="top"><img src="images/sun_l_sp.gif" width="5" height="10" border="0"></td>
                      <td><a href="ShowTeam.asp?cid=<%=cid%>&tid=<%=rs("tid")%>" class="b"> 
                        <%=rs("name")%></a></td>
                    </tr>
                    <%
rs.movenext
loop
rs.close
set rs=nothing
%>
                  </table>
                  <!-- ========================================================================================================		  

													Tournament table ============star

 ======================================================================================================== -->
                </td>
              </tr>
            </table>
            <!-- ========================================================================================================		  

													cup table ============stop

 ======================================================================================================== -->
          </td>
        </tr>
        <tr> 
          <td>&nbsp;</td>
        </tr>
      </table> </td>
    <td valign="top"><br>
      <br>
<%
tid=request("tid")
if tid="" then
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Please select a team ! ');</SCRIPT>")
response.End()
end if
%>
<%
tsql="select * from team where tid="&tid&" order by tid desc" 
Set trs=Server.CreateObject("ADODB.RecordSet") 
trs.Open tsql,conn,1,1 
%>
<%
if trs.eof or trs.bof then
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Team is not exist ! ');</SCRIPT>")
response.End()
trs.close
set trs=nothing
end if
%>
      <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td width="18">&nbsp;</td>
          <td width="278" valign="top"><TABLE width=100% border=0 align="center" cellPadding=0 cellSpacing=0>
              <TBODY>
                <TR> 
                  <TD align="right"> <TABLE bgColor=#BECFDF border=0 cellPadding=7 cellSpacing=0 
            width=275>
                      <TBODY>
                        <TR> 
                          <TD width="260"> <TABLE bgColor=#ffffff border=0 cellPadding=0 cellSpacing=0 
                  width=100%>
                              <TBODY>
                                <TR> 
                                  <TD width="242"><a href="<%=trs("photo")%>" target="_blank"><img src="<%=trs("photo")%>" alt="Click Here To Zoom In <%=trs("name")%>'s Photo" width="260" height="153" border="0"></a></TD>
                                </TR>
                              </TBODY>
                            </TABLE></TD>
                        </TR>
                      </TBODY>
                    </TABLE></TD>
                </TR>
              </TBODY>
            </TABLE>
            <table width="100%" border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td height="3"><img src="1.gif" width="1" height="1"></td>
              </tr>
            </table>
            <table width="275" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#eae7e7">
              <tr> 
                <td width="30" height="18" align="center">&nbsp;</td>
                <td width="80" align="center">Home color:</td>
                <td width="20"><table width="10" border="0" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td height="10" bgcolor="<%=trs("hcolor")%>"><img src="1.gif" width="1" height="1"></td>
                    </tr>
                  </table></td>
                <td>&nbsp;</td>
                <td width="80" align="center">Away color:</td>
                <td width="20"><table width="10" border="0" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td height="10" bgcolor="<%=trs("acolor")%>"><img src="1.gif" width="1" height="1"></td>
                    </tr>
                  </table></td>
                <td width="30">&nbsp;</td>
              </tr>
            </table></td>
          <td width="3">&nbsp;</td>
          <td width="278" valign="top"><TABLE border=0 cellPadding=0 cellSpacing=0 width=100%>
              <TBODY>
                <TR> 
                  <TD> <TABLE bgColor=#BECFDF border=0 cellPadding=7 cellSpacing=0 
            width=100%>
                      <TBODY>
                        <TR> 
                          <TD> <TABLE bgColor=#ffffff border=0 cellPadding=0 cellSpacing=0 
                  width=100%>
                              <TBODY>
                                <TR> 
                                  <TD><table width="97%" border="0" align="center" cellpadding="2" cellspacing="1">
                                      <tr bgcolor="#FFFFFF" onmouseout="this.style.backgroundColor='#FFFFFF';" 
                          onmouseover="this.style.backgroundColor='#BECFDF';"> 
                                        <td width="75">Team name:</td>
                                        <td width="170"><%=trs("name")%></td>
                                      </tr>
                                      <tr bgcolor="#FFFFFF" onmouseout="this.style.backgroundColor='#FFFFFF';" 
                          onmouseover="this.style.backgroundColor='#BECFDF';"> 
                                        <td>Found at:</td>
                                        <td><%=trs("createdate")%></td>
                                      </tr>
                                      <tr bgcolor="#FFFFFF" onmouseout="this.style.backgroundColor='#FFFFFF';" 
                          onmouseover="this.style.backgroundColor='#BECFDF';"> 
                                        <td>City:</td>
                                        <td><%=trs("city")%></td>
                                      </tr>
                                      <tr bgcolor="#FFFFFF" onmouseout="this.style.backgroundColor='#FFFFFF';" 
                          onmouseover="this.style.backgroundColor='#BECFDF';"> 
                                        <td>TEL/FAX:</td>
                                        <td><%=trs("telfax")%></td>
                                      </tr>
                                      <tr bgcolor="#FFFFFF" onmouseout="this.style.backgroundColor='#FFFFFF';" 
                          onmouseover="this.style.backgroundColor='#BECFDF';"> 
                                        <td>Email:</td>
                                        <td><%=trs("email")%></td>
                                      </tr>
                                      <tr bgcolor="#FFFFFF" onmouseout="this.style.backgroundColor='#FFFFFF';" 
                          onmouseover="this.style.backgroundColor='#BECFDF';"> 
                                        <td>Manager:</td>
                                        <td> <%
psql="select * from player where pid="&trs("mpid")&" order by tid desc" 
Set prs=Server.CreateObject("ADODB.RecordSet") 
prs.Open psql,conn,1,1
if not ( prs.eof or prs.bof ) then
response.Write(prs("name"))
end if
prs.close
set prs=nothing
%> </td>
                                      </tr>
                                      <tr bgcolor="#FFFFFF" onmouseout="this.style.backgroundColor='#FFFFFF';" 
                          onmouseover="this.style.backgroundColor='#BECFDF';"> 
                                        <td>Ast.Manager:</td>
                                        <td> <%
psql="select * from player where pid="&trs("apid")&" order by tid desc" 
Set prs=Server.CreateObject("ADODB.RecordSet") 
prs.Open psql,conn,1,1
if not ( prs.eof or prs.bof ) then
response.Write(prs("name"))
end if
prs.close
set prs=nothing
%> </td>
                                      </tr>
                                      <tr bgcolor="#FFFFFF" onmouseout="this.style.backgroundColor='#FFFFFF';" 
                          onmouseover="this.style.backgroundColor='#BECFDF';"> 
                                        <td>Home page:</td>
                                        <td> <%
homepage=trim(trs("homepage"))
if isnull(homepage) then homepage="TeamHomePageDefault.asp"
%> <a href="<%=homepage%>" target="_blank" class="a_team">Team's 
                                          HomePage</a> </td>
                                      </tr>
                                    </table></TD>
                                </TR>
                              </TBODY>
                            </TABLE></TD>
                        </TR>
                      </TBODY>
                    </TABLE>
                    <table width="100%" border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td height="3"><img src="1.gif" width="1" height="1"></td>
                      </tr>
                    </table>
                    <table width="100%" border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td bgcolor="#eae7e7"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                            <tr> 
                              <td width="13" height="18">&nbsp;</td>
                              <td width="88">The sponsor: </td>
                              <td width="2">&nbsp;</td>
                              <td width="186"><a href="<%=trs("sponsorhomepage")%>" target="_blank" class="a_team"><%=trs("sponsor")%></a></td>
                            </tr>
                          </table></td>
                      </tr>
                    </table></TD>
                </TR>
              </TBODY>
            </TABLE></td>
          <td width="18">&nbsp;</td>
        </tr>
      </table> 
      <br>
      <br>
      <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="18">&nbsp;</td>
          <td><table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#CCCCCC">
              <tr> 
                <td height="30" align="center" bgcolor="#FFFFFF"><b><%=trs("name")%></b></td>
              </tr>
              <tr> 
                <td bgcolor="#FFFFFF"><b>Manager:</b></td>
              </tr>
              <tr> 
                <td bgcolor="#FFFFFF">
                  <!-- ========================================================================================================		  

													Player table ============star

 ======================================================================================================== -->
                  <table width="96%" border="0" align="center" cellpadding="0" cellspacing="1">
                    <tr> 
                      <td width="93" align="center" bgcolor="#FFFFFF">&nbsp;</td>
                      <td width="111" bgcolor="#FFFFFF"><b>Name</b></td>
                      <td width="149" bgcolor="#FFFFFF"><b>Email</b></td>
                      <td width="160" bgcolor="#FFFFFF"><strong>Tel</strong></td>
                      <td width="160" bgcolor="#FFFFFF"><strong>Contact</strong></td>
                    </tr>
                    <%
sql="select * from player where ( ptype='m' or ptype='a' ) and tid=" & tid & " order by ptype "
Set srs=Server.CreateObject("ADODB.RecordSet") 
srs.Open sql,conn,1,1
if not ( srs.eof or srs.bof ) then
%>
                    <tr> 
                      <td bgcolor="#FFFFFF"><strong>Manager</strong></td>
                      <td bgcolor="#FFFFFF"><%=srs("name")%></td>
                      <td bgcolor="#FFFFFF"><%=srs("email")%></td>
                      <td bgcolor="#FFFFFF"><%=srs("tel")%></td>
                      <td bgcolor="#FFFFFF"><%=srs("contact")%></td>
                    </tr>
<%
srs.movenext
if not ( srs.eof or srs.bof ) then
%>
                    <tr> 
                      <td bgcolor="#FFFFFF"><strong>Asst.Manager</strong></td>
                      <td bgcolor="#FFFFFF"><%=srs("name")%></td>
                      <td bgcolor="#FFFFFF"><%=srs("email")%></td>
                      <td bgcolor="#FFFFFF"><%=srs("tel")%></td>
                      <td bgcolor="#FFFFFF"><%=srs("contact")%></td>
                    </tr>
                    <%
end if
end if
srs.close
set srs=nothing
%>
                  </table>
                  <!-- ========================================================================================================		  

													Player table ============star

 ======================================================================================================== -->
                </td>
              </tr>
              <tr> 
                <td bgcolor="#FFFFFF"><b>player:</b></td>
              </tr>
              <tr> 
                <td bgcolor="#FFFFFF"> 
                  <!-- ========================================================================================================		  

													Player table ============star

 ======================================================================================================== -->
                  <table width="96%" border="0" align="center" cellpadding="0" cellspacing="1">
                    <tr> 
                      <td width="170" bgcolor="#FFFFFF"><b>Jersey</b></td>
                      <td width="345" bgcolor="#FFFFFF"><b>Name</b></td>
                    </tr>
<%
sql="select * from player where ptype<>'m' and ptype<>'a' and tid="&tid
Set srs=Server.CreateObject("ADODB.RecordSet") 
srs.Open sql,conn,1,1
if not ( srs.eof or srs.bof ) then
do
%>
                    <tr> 
                      <td bgcolor="#FFFFFF"><%=srs("jersey")%></td>
                      <td bgcolor="#FFFFFF"><%=srs("name")%></td>
                    </tr>
<%
srs.movenext
loop while not srs.eof
end if
srs.close
set srs=nothing
%>
                  </table>
                  <!-- ========================================================================================================		  

													Player table ============star

 ======================================================================================================== -->
                </td>
              </tr>
              <tr> 
                <td height="15" bgcolor="#FFFFFF"><strong>description:</strong></td>
              </tr>
              <tr> 
                <td height="200" align="center" valign="top" bgcolor="#FFFFFF"> <table width="94%" border="0" cellspacing="2" cellpadding="2">
                    <tr> 
                      <td>
					  <%
					  zdescription=trs("description")
					  zdescription=replace(zdescription,chr(13),"<br>")
					  response.Write(zdescription)
					  %>
					  </td>
                    </tr>
                  </table></td>
              </tr>
            </table></td>
          <td width="18">&nbsp;</td>
        </tr>
      </table>
      
    </td>
  </tr>
</table>
<%
trs.close
set trs=nothing
%>
<!--#include file="Copyright.asp"-->
</body>
</html>
