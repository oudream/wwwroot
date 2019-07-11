<!--#include file="Top_sun.asp"-->
<table width="776" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="181" valign="top">
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
    <td> <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td height="27">&nbsp;</td>
        </tr>
      </table>
      <%
sql="select * from events order by eid desc" 
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
<table border="0" cellspacing="0" cellpadding="0" width="370">
        <tr> 
          <td width="370" height="40" colspan="6" valign="top" bgcolor="#FFFFFF" class="header_b">
            SUNSPORTS NEW EVENTS</td>
        </tr>
        <tr> 
          <td colspan="4"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td> 
<%
if trim(rs("photo"))<>"" then
%> 
<a href="events_show.asp?eid=<%=rs("eid")%>"><img src=<%=rs("photo")%> border="0" width="360" height="200"></a> 
<%
end if
%></td>
              </tr>
            </table></td>
        </tr>
<%
do while not rs.eof 
%>
        <tr> 
          <td colspan="4"><br>
<a href="events_show.asp?eid=<%=rs("eid")%>" class="b"><b><%=rs("name")%></b></a></td>
        </tr>
        <tr> 
          <td colspan="4">            <%
		  if len(rs("memo"))>300 then
		  response.Write(left(rs("memo"),298)&"...")
		  else
		  response.Write(rs("memo"))
		  end if
		  %>
            <br><br>
          </td>
        </tr>
        <tr> 
          <td colspan="4"><table width="100%" border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td height="5" bgcolor="#f5f5f5"><img src="1.gif" width="1" height="1"></td>
              </tr>
            </table>
            
          </td>
        </tr>
        <%
rs.movenext
i=i+1                                                                     
if i>=rs.pagesize then exit do                                                           
loop
rs.movefirst
%>
        <tr> 
          <td height="35" colspan="6">&nbsp;</td>
        </tr>
      </table>
      <%
rs.close
set rs=nothing
%> </td>
    <td width="20">&nbsp;</td>
    <td width="181" valign="top"> <table width="100%" border="0" cellpadding="0" cellspacing="0">
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