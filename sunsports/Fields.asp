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
	<td width="20">&nbsp;</td>
    <td> <%
sql="select * from field" 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1 

if rs.eof or rs.bof then response.Redirect("error.asp?err=003")

if not isempty(request("page")) then   
pagecount=cint(request("page"))   
else   
pagecount=1
end if

rs.pagesize=5
rs.AbsolutePage=pagecount
%>


 <table border="0" cellspacing="0" cellpadding="0" width="100%">
        <tr> 
          <td height="22" colspan="6" bgcolor="#FFFFFF"> 
            <table width="100%" border="0" cellpadding="3" cellspacing="0">
              <tr> 
                <td height="25">&nbsp;</td>
              </tr>
              <tr>
                <td class="header_b">Sunsports List of Fields</td>
              </tr>
            </table>
            <br>
          </td>
        </tr>
        <%
do while not rs.eof 
%>
        <tr> 
          <td colspan="4">
		  
		  
		  
		  
		  <table width="100%" border="0" align="center" cellpadding="3" cellspacing="0">
              <tr> 
                <td rowspan="3" bgcolor="#FFFFFF"><img src=<%=rs("photo")%> border="0" width="200" height="150"></td>
                <td width="50%" height="22" bgcolor="#FFFFFF" class="header"><%=rs("name")%></td>
				
				
				
                <td width="50%" bgcolor="#FFFFFF">Night -- <%=rs("tfnight")%></td>
              </tr>
              <tr bgcolor="#FFFFFF"> 
                <td height="22" colspan="2"><%=rs("location")%></td>
              </tr>
              <tr align="left" valign="top" bgcolor="#FFFFFF"> 
                <td height="80" colspan="2">ABOUT : <%=rs("memo")%></td>
              </tr>
            </table>
			
			
			
			
			</td>
        </tr>
        <tr> 
          <td colspan="3"> <br>
            <hr width="540" color="#0099FF" noshade size=2> <br> </td>
        <tr> 
          <%
rs.movenext
i=i+1                                                                     
if i>=rs.pagesize then exit do                                                           
loop
rs.movefirst
%>
        <tr> 
          <td height="35" colspan="6"> <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td width="45%"> <table border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr> 
                      <%
zj=1
for zi=1 to rs.pagecount
%>
                      <td width="50"><a href="Fields.asp?page=<%=zj%>" class="b">|&nbsp;<%=zj%>&nbsp;|</a></td>
                      <%
zj=zj+1
next
%>
                    </tr>
                  </table></td>
                <td width="55%" align="left"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                      <td align="center"> [ FOELDS:&nbsp;&nbsp;<font color=red><b><%=rs.recordcount%></b></font>&nbsp;&nbsp;PAGE:&nbsp;&nbsp;<font color=red><b><%=pagecount%></b></font>&nbsp;&nbsp;/&nbsp;&nbsp;<font color=red><b><%=rs.pagecount%></b></font> ]</td>
                    </tr>
                  </table></td>
              </tr>
            </table></td>
        </tr>
      </table>
      <%
rs.close
set rs=nothing
%> </td>
    <td width="20">&nbsp;</td>
  </tr>
</table>
<!--#include file="Copyright.asp"-->