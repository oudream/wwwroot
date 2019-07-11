<!--#include file="Top_sun.asp"-->
<table width="776" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="181" valign="top"><table width="100%" border="0" cellpadding="3" cellspacing="0">
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
            <!-- ========================================================================================================		  

													cup table ============stop

 ======================================================================================================== -->
          </td>
        </tr>
        <tr> 
          <td>&nbsp;</td>
        </tr>
      </table></td>
    <td width="20">&nbsp;</td>
    <td> <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td height="27">&nbsp;</td>
        </tr>
      </table>
      <%
eid=request("eid")
if eid="" then
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Please select the Events you want to View at first ! ');</SCRIPT>")
response.End()
end if
%>
<%
sql="select * from events where eid="&eid&" order by eid desc" 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1 
if rs.eof or rs.bof then response.Redirect("error.asp?err=003")
%>
<table border="0" cellspacing="0" cellpadding="0" width="360" align="center">
  <tr> 
          <td height="50" colspan="6" valign="top" bgcolor="#FFFFFF" class="header_b">RECENT NEWS 
           </td>
  </tr>
  <tr> 
    <td colspan="4"> <table width="360" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td height="22" valign="middle"> <%
if trim(rs("photo"))<>"" then
%> <a href="events_show.asp?eid=<%=rs("eid")%>"><img src=<%=rs("photo")%> border="0" width="360" height="200"></a> 
                  <%
end if
%> &nbsp;</td>
              </tr>
              <tr> 
                <td height="25" align="center"><b><a href="events_show.asp?eid=<%=rs("eid")%>" class="b"><%=rs("name")%></a></b>
                  </td>
              </tr>
              <tr align="left" valign="top">
                <td height="50" align="left" valign="top">&nbsp;&nbsp;&nbsp;&nbsp; 
                  <%
		  response.Write(rs("memo"))
		  %>
                </td>
              </tr>
            </table></td>
  </tr>
<%
rs.close
set rs=nothing
%>



</table>

	</td>
    <td width="20">&nbsp;</td>
    <td width="181" valign="top">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td height="18">&nbsp;</td>
        </tr>
      </table>
<table width="100%" border="0" cellpadding="3" cellspacing="0">
        <tr> 
          <td><table width="170" border="0" align="center" cellpadding="0" cellspacing="3">
              <tr> 
                <td><a href="Events_all.asp"><b><font color="#1F4A71">OTHER NEW EVENTS</font></b></a> 
                  <font color="#1F4A71">&nbsp;+ +</font> </td>
              </tr>
            </table>
            <table width="170" border="0" align="center" cellpadding="0" cellspacing="3">
              <%
sql="select * from events order by eid desc" 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1 

if not ( rs.eof or rs.bof ) then 
do
%>
              <tr> 
                <td width="5" valign="top"><img src="images/sun_l_sp.gif" width="5" height="10" border="0"></td>
                <td><a href="events_show.asp?eid=<%=rs("eid")%>" class="b"><%=rs("name")%></a> 
                </td>
              </tr>
              <%
rs.movenext
loop while not rs.eof 
end if
rs.close
set rs=nothing
%>
            </table>
            <br>
            <br>
            <br>
            <br>
          </td>
        </tr>
        <tr> 
          <td>&nbsp;</td>
        </tr>
      </table>	  
	  	</td>
  </tr>
</table>
<!--#include file="Copyright.asp"-->