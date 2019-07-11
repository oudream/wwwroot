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
            <table width="169" border="0" align="center" cellpadding="0" cellspacing="3">
              <tr> 
                <td width="163"><a href="TeamProfile.asp"><b><font color="#1F4A71">TEAMPORFILE</font></b></a> 
                  <font color="#1F4A71">&nbsp;+ +</font> </td>
              </tr>
            </table>
            <table width="170" border="0" align="center" cellpadding="0" cellspacing="3">
              <tr> 
                <td><a href="TeamProfile.asp"><b><font color="#1F4A71">FIELDS 
                  BOOKING</font></b></a></td>
              </tr>
            </table>
            <table width="170" border="0" align="center" cellpadding="0" cellspacing="3">
              <%
sql="select * from field" 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1
do while not rs.eof
%>
              <tr> 
                <td width="5" valign="top"> <img src="images/sun_l_sp.gif" width="5" height="10" border="0"> 
                </td>
                <td><a href="Calendar.asp?field=<%=rs("fid")%>" title="<%=rs("name")%>" class="b"><%=rs("name")%></a></td>
              </tr>
              <%
rs.movenext
loop
rs.close
set rs=nothing
%>
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
    <td width="20">&nbsp; </td>
    <td valign="top"> <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td height="27">&nbsp;</td>
        </tr>
      </table>
<%
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Edit is complete. ');</SCRIPT>")
sql="select * from player where pid="&session("user_pid")
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1
%>
	<table width="90%" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#6184A3">
        <tr align="center" bgcolor="#FFFFFF"> 
          <td height="30" colspan="2" ><b>SIGN UP COMPLETE</b></td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="28%" height="30">User id:</td>
          <td width="72%"><%=rs("uid")%>&nbsp;</td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td height="30">Full Name:</td>
          <td><%=rs("name")%>&nbsp;</td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td height="30">Telphone:</td>
          <td><%=rs("tel")%>&nbsp;</td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td height="30">Email:</td>
          <td><%=rs("email")%>&nbsp;</td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td height="30">ZIP:</td>
          <td><%=rs("zip")%>&nbsp;</td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td height="30">Address:</td>
          <td><%=rs("contact")%>&nbsp;</td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td height="30" colspan="2">&nbsp;</td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td height="30" colspan="2"><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr bgcolor="#FFFFFF"> 
                <td width="50%" align="center"><a href="index.asp" class="b">Home</a></td>
                <td align="center"><a href="Shopping_Reg_Edit.asp">Manage Myself</a>&nbsp;</td>
              </tr>
            </table></td>
        </tr>
      </table>
<%
rs.close
set rs=nothing
%>
	  </td>
    <td width="20">&nbsp;</td>
    <td width="181" valign="top">&nbsp;</td>
  </tr>
</table>
<!--#include file="Copyright.asp"-->