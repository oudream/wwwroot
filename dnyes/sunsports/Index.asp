<!--#include file="Top_sun.asp"-->
<TABLE align=center bgColor=#ffffff border=0 borderColor=#ffffff cellPadding=0 
cellSpacing=0 width=776>
  <TBODY>
    <TR> 
      <TD height=44> <TABLE border=0 cellPadding=0 cellSpacing=0 height=44 width="100%">
          <TBODY>
            <TR> 
              <TD width="100%" align=left><img src="images/7sun.gif" width="776" height="45"></TD>
            </TR>
          </TBODY>
        </TABLE></TD>
    </TR>
    <TR> 
      <TD bgColor=#efefef height=18 width=778> <TABLE border=0 cellPadding=0 cellSpacing=0 height=18 width="100%">
          <TBODY>
            <TR> 
              <TD width=13>&nbsp;</TD>
              <TD class=locator width=763>Announcements</TD>
            </TR>
          </TBODY>
        </TABLE></TD>
    </TR>
  </TBODY>
</TABLE>
<table width="776" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="181" rowspan="4" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td valign="top"><!--#include file="announcements.asp"--></td>
        </tr>
        <tr> 
          <td>&nbsp;</td>
        </tr>
        <tr> 
          <td valign="top"> 
            <table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
<%
dim rrhscore(1)
dim rrascore(1)
sql="select * from match order by mid desc" 
Set rs=Server.CreateObject("ADODB.RecordSet")
rs.Open sql,conn,1,1
i=0                                                                    
%>
              <tr bgcolor="#FFFFFF"> 
                <td colspan="3" align="center"><strong><font color="#1F4A71">FIXTURE</font></strong></td>
              </tr>
              <tr bgcolor="#FFFFFF"> 
                <td width="42%" align="center"><strong>HOME TEAM</strong></td>
                <td width="14%" align="center"><strong>VS</strong></td>
                <td width="44%" align="center"><strong>AWAY TEAM</strong></td>
              </tr>
<%
do while not rs.eof
rrhscore(0)=100
rrascore(0)=100
rrhscore(1)=101
rrascore(1)=101
rsql="select * from rresult where mid="&rs("mid")
Set rrs=Server.CreateObject("ADODB.RecordSet")
rrs.Open rsql,conn,1,1
zi=0
do while not rrs.eof
rrhscore(zi)=rrs("hscore")
rrascore(zi)=rrs("ascore")
zi=zi+1
rrs.movenext
loop
rrs.close
set rrs=nothing
if isempty(rrhscore(0)) or isempty(rrhscore(1)) or isempty(rrascore(0)) or isempty(rrascore(1)) then
rrstr="no"
else
if rrhscore(0)=rrhscore(1) and rrascore(0)=rrascore(1) then rrstr="yes"
end if



if rrstr="yes" then

i=i+1

matchdata=cstr(rs("mmonth"))&"/"&cstr(rs("mday"))&"/"&cstr(rs("myear"))
matchtime=cstr(rs("mhour"))&":"&cstr(rs("mminute"))
%>
              <tr bgcolor="#FFFFFF"> 
                <td align="center"><%=rs("htname")%></td>
                <td align="center"><%=rrhscore(0)&":"&rrascore(0)%></td>
                <td align="center"><%=rs("atname")%></td>
              </tr>
              <tr align="center" bgcolor="#FFFFFF"> 
                <td colspan="3"><strong>TIME: <%=matchdata%> &nbsp; <%=FormatDateTime(matchtime,4)%></strong></td>
              </tr>
              <tr bgcolor="#FFFFFF"> 
                <td colspan="3" align="center">&nbsp;</td>
              </tr>
<%
end if
rs.movenext
if i>=5 then exit do                                                           
loop
rs.close
set rs=nothing
%>
            </table>
          </td>
        </tr>
        <tr> 
          <td>&nbsp;</td>
        </tr>
        <tr> 
          <td>
		  
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1">
              <tr bgcolor="#FFFFFF"> 
                <td height="25"><font color="#1F4A71"><b>TOP 10 OF ALL TEAMS</b></font></td>
              </tr>
              <tr bgcolor="#FFFFFF"> 
                <td><b>Team Name (PT)</b></td>
              </tr>
              <%
sssql="select * from standing order by points desc"
Set ssrs=Server.CreateObject("ADODB.RecordSet") 
ssrs.Open sssql,conn,1,1
if not ( ssrs.eof or ssrs.bof ) then
i=0
do while not ssrs.eof
%>
              <tr bgcolor="#FFFFFF"> 
                <td height="18"><font color="#666666"><%=ssrs("tname")%> (<%=ssrs("points")%>)</font>
                  </td>
              </tr>
              <%
ssrs.movenext
i=i+1
if i>=10 then exit do                                                           
loop
end if
ssrs.close
set ssrs=nothing
%>
            </table>		  
		  </td>
        </tr>
        <tr> 
          <td>&nbsp;</td>
        </tr>
        <tr> 
          <td>
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1">
              <tr bgcolor="#FFFFFF"> 
                <td height="25"><font color="#1F4A71"><b>TOP 10 PLAYERS HOT SHOTS</b></font></td>
              </tr>
              <tr bgcolor="#FFFFFF"> 
                <td><b>Player Name (Goals)</b></td>
              </tr>
              <%
ssssql="SELECT default_score ,name FROM player order by default_score desc"
Set sssrs=Server.CreateObject("ADODB.RecordSet") 
sssrs.Open ssssql,conn,1,1
if not ( sssrs.eof or sssrs.bof ) then
i=0
do while not sssrs.eof
%>
              <tr bgcolor="#FFFFFF"> 
                <td><%=sssrs("name")%> ( <%=sssrs("default_score")%> )</td>
              </tr>
              <%
sssrs.movenext
i=i+1
if i>=10 then exit do                                                           
loop
end if
sssrs.close
set sssrs=nothing
%>
            </table>		  
		  
		  </td>
        </tr>
        <tr> 
          <td>&nbsp;</td>
        </tr>
      </table></td>
    <td width="20" rowspan="4">&nbsp;</td>
    <td rowspan="4" valign="top">
<!-- =====================================================================================================
             DESIGN BY WWW.DNYES.COM
			 support@dnyes.com
			 2004-3-10
====================================================================================================== -->	
<%
sql="select * from events order by eid desc" 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1 
i=0
%> 
<table border="0" cellspacing="0" cellpadding="0" width="100%">
        <tr> 
          <td height="29" colspan="6" valign="bottom" bgcolor="#FFFFFF"><font color="#1F4A71"><strong><br>
            Sunsports New Events</strong></font></td>
        </tr>
<%
if not ( rs.eof or rs.bof ) then
%>
        <tr> 
          <td colspan="3">
		  <table width="100%" border="0" cellspacing="2" cellpadding="8">
              <tr> 
                <td height="22" valign="top">
<TABLE cellSpacing=0 cellPadding=8 align=left border=0>
                      <TBODY>
                        <TR> 
                          <TD>
<%
if trim(rs("photo"))<>"" then
%>
<a href="events_show.asp?eid=<%=rs("eid")%>"><img src=<%=rs("photo")%> border="0" width="175" height="120"></a> 
<%
end if
%>
                                  </TD>
                        </TR>
                      </TBODY>
                    </TABLE>
					
					
          
		  <font color="#666666">
		  "
		  <%
		  if len(rs("memo"))>300 then
		  response.Write(left(rs("memo"),298))
		  else
		  response.Write(rs("memo"))
		  end if
		  %>
		  " ... 
		  </font>
		  <a href="events_show.asp?eid=<%=rs("eid")%>" class="b">more info</a>
                </td>
              </tr>
            </table></td>
        </tr>
<%
do while not rs.eof 
%>
        <tr> 
          <td height="18" colspan="4">
		  <a href="events_show.asp?eid=<%=rs("eid")%>" class="b"><%=rs("name")%></a>
		  </td>
        </tr>
        <%
rs.movenext
i=i+1                                                                     
if i>=10 then exit do                                                           
loop
%>
        <tr> 
          <td colspan="6">&nbsp;</td>
        </tr>
<%
end if
%>
      </table>
      <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td><hr></td>
        </tr>
      </table> <br>
      <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td><font color="#1F4A71"><strong>Sunsports New Teams</strong></font></td>
        </tr>
        <tr> 
          <td><font color="#666666">Latest teams joined in our team profile here. 
            If your team want to join us, call us please. Let's fun in our soccer 
            !! </font></td>
        </tr>
        <%
sql="select top 8 * from team order by tid desc" 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.Open sql,conn,1,1 
do while not rs.eof
%>
        <tr> 
          <td height="19">
		   <li><a href="ShowTeam.asp?tid=<%=rs("tid")%>" target="_blank" class="b"><%=rs("name")%></a><FONT color="#666666">
<%
psql="select name from player where pid="&rs("mpid")
Set prs=Server.CreateObject("ADODB.RecordSet") 
prs.Open psql,conn,1,1 
response.Write(" , manager: "&prs("name"))
prs.close
%>
<%
psql="select name from player where pid="&rs("apid")
prs.Open psql,conn,1,1 
response.Write(" , asst.manager: "&prs("name"))
prs.close
set prs=nothing
%>
		   </FONT>
		   </td>
        </tr>
        <%
rs.movenext
loop
rs.close
set rs=nothing
%>
        <tr> 
          <td>&nbsp;</td>
        </tr>
      </table>
      <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td><hr></td>
        </tr>
      </table>
      <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td><strong><font color="#1F4A71">Sunsports &amp; Recreation</font></strong><br>
            <font color="#666666">Sunsports is a private events organiser for 
            sports related activities that caters to individuals, companies, MNC's, 
            schools, institutions and government agencies. We actively encourage 
            participation in sporting activities for the joint benefit of promoting 
            a healthy lifestyle as well as team co-operation through fun sports. 
            Organised sporting events offered include: </font><br>
			<FONT color="#666666">
            <li>Soccer <br>
            <li>Ball Games <br>
            <li>X country <br>
            <li>Island Cycling <br>
            <li>Water Sports <br>
			</FONT>
			</td>
        </tr>
      </table> </td>
    <td width="20" rowspan="4" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td height="43">&nbsp;</td>
        </tr>
      </table>
      <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td>&nbsp;</td>
          <td height="930" background="images/10x1_blue.gif">&nbsp;</td>
        </tr>
      </table></td>
    <td width="181" valign="top">
	                <!--#include file="User_l.asp" -->
	            <br><!--#include file="Shopping_Search.asp" -->
	            <br>
<table width="170" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="3">
              <tr>
                
          <td><a href="Shopping_sort_view.asp"><b><font color="#1F4A71">SUNSPORTS PRODUCTS 
            </font></b></a></td>
              </tr>
            </table>
<%
sql="select top 3 * from sort" 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1 
do while not rs.eof
%>
	<table width="100%" border="0" cellspacing="3" cellpadding="0">
        <tr> 
          <td width="5" valign="top"><img src="images/sun_l_sp.gif" width="5" height="10" border="0"></td>
          <td><a href="Shopping_sort_view.asp?sort_id=<%=rs("sort_id")%>" class="b"><%=rs("name")%></a></td>
        </tr>
      </table>
<%
rs.movenext
loop
rs.close
set rs=nothing
%>
    </td>
  </tr>
</table>				<br><!--#include file="Shopping_New_in.asp" -->


           </td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
   
<!--#include file="Copyright.asp"-->