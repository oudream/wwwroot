
<!--#include file="adminconn.asp" -->
<%
atid=request("atid")
htid=request("htid")
tid=atid
mmid=request("mid")
if tid="" and mmid="" then
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Please select the team you want to add scorer or discipline at first ! ');</SCRIPT>")
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
            
          <td height="30" align="center" valign="middle"><strong> AWAY Match Result 
            VIEW</strong></td>
          </tr>
        </table></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td><table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#6184A3">
          <tr bgcolor="#FFFFFF"> 
            <td height="25" colspan="4"><b> Hot Shots</b></td>
          </tr>
          <tr bgcolor="#FFFFFF" onmouseout="this.style.backgroundColor='#FFFFFF';" 
                          onmouseover="this.style.backgroundColor='#BECFDF';"> 
            <td width="190" height="20" align="center"><b><b>Team Name</b></b></td>
            <td width="45" align="center"><b>Jersey</b></td>
            <td align="center"><b>Name</b></td>
            <td width="91" align="center"><b>Goals</b></td>
          </tr>
          <%
ssql="select * from scorer where mid="&mmid&" and tid="&tid
Set srs=Server.CreateObject("ADODB.RecordSet") 
srs.Open ssql,conn,1,1
if not ( srs.eof or srs.bof ) then
do
%>
          <tr bgcolor="#FFFFFF" onmouseout="this.style.backgroundColor='#FFFFFF';" 
                                 onmouseover="this.style.backgroundColor='#BECFDF';"> 
            <td height="18" align="center"><%=session("manager_tname")%></td>
            <td align="center"><%=srs("jersey")%></td>
            <td align="center"><%=srs("pname")%></td>
            <td align="center"><%=srs("score")%></td>
          </tr>
          <%
srs.movenext
loop while not srs.eof
else
response.Write("<tr bgcolor=#FFFFFF><td height=18 colspan=4 align=center><strong>NONE DISCIPLINE</strong></td></tr>")
end if
srs.close
set srs=nothing
%>
        </table></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td><table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#6184A3">
          <tr bgcolor="#FFFFFF"> 
            <td height="25" colspan="5"><b> Worst Offenders</b></td>
          </tr>
          <tr bgcolor="#FFFFFF" onmouseout="this.style.backgroundColor='#FFFFFF';" 
                          onmouseover="this.style.backgroundColor='#BECFDF';"> 
            <td width="190" height="20" align="center"><b>Team Name</b></td>
            <td width="45" align="center"><b>Jersey</b></td>
            <td align="center"><b>Name</b></td>
            <td width="45" align="center"><b>Yellow</b></td>
            <td width="45" align="center"><b>Red</b></td>
          </tr>
          <%
csql="select * from card where mid="&mmid&" and tid="&tid
Set crs=Server.CreateObject("ADODB.RecordSet") 
crs.Open csql,conn,1,1
if not ( crs.eof or crs.bof ) then
do
%>
          <tr bgcolor="#FFFFFF" onmouseout="this.style.backgroundColor='#FFFFFF';" 
                                 onmouseover="this.style.backgroundColor='#BECFDF';"> 
            <td height="18" align="center"><%=session("manager_tname")%></td>
            <td align="center"><%=crs("jersey")%></td>
            <td align="center"><%=crs("pname")%></td>
            <td align="center"><%=crs("yellow")%></td>
            <td align="center"><%=crs("red")%></td>
          </tr>
          <%
crs.movenext
loop while not crs.eof
else
response.Write("<tr bgcolor=#FFFFFF><td height=18 colspan=5 align=center><strong>NONE SCORE</strong></td></tr>")
end if
crs.close
set crs=nothing
%>
        </table></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td> 
        <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#6184A3">
        <tr bgcolor="#FFFFFF"> 
          <td width="305" height="25"><b><strong><a href="match_r_c2.asp?option=addold&mid=<%=mmid%>&htid=<%=htid%>&atid=<%=atid%>">Add 
            Result as the sample</a></strong></b></td>
          <td width="36">&nbsp;</td>
          <td width="245"><b><strong><a href="match_r_c2.asp?option=deledit&mid=<%=mmid%>&htid=<%=htid%>&atid=<%=atid%>">Edit 
            it and Add</a></strong></b></td>
          <td width="186">&nbsp;</td>
          <td width="189">&nbsp;</td>
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
