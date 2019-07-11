<script language=JavaScript>                                              
<!--//       
       
var version = "other"       
browserName = navigator.appName;       
browserVer = parseInt(navigator.appVersion);       
       
if (browserName == "Netscape" && browserVer >= 3) version = "n3";       
else if (browserName == "Netscape" && browserVer < 3) version = "n2";       
else if (browserName == "Microsoft Internet Explorer" && browserVer >= 4) version = "e4";       
else if (browserName == "Microsoft Internet Explorer" && browserVer < 4) version = "e3";       
       
function marquee1()       
{       
	if (version == "e4")       
	{       
		document.write("<marquee behavior=scroll direction=up width=180 height=200 scrollamount=1 scrolldelay=50 onmouseover='this.stop()' onmouseout='this.start()'>")       
	}       
}       
      
function marquee2()       
{       
	if (version == "e4")       
	{       
		document.write("</marquee>")       
	}       
}       
      
//-->       
</script>
                            <script language=JavaScript>marquee1();</script>
<%
sql="select top 10 * from announcements order by id desc" 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.Open sql,conn,1,1 
do while not rs.eof
%>
      
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1">
  <tr bgcolor="#FFFFFF" onmouseout="this.style.backgroundColor='#FFFFFF';" 
                          onmouseover="this.style.backgroundColor='#f5f5f5';"> 
    <td width="792">
      <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td width="5" valign="top" class="date"><img src="images/sun_l_an.gif" width="5" height="10"></td>
          <td valign="bottom" class="date"><%=rs("inserttime")%></td>
        </tr>
        <tr> 
          <td colspan="2"> <%
if len(rs("content"))>1000 then
response.Write(left(rs("content"),998)&"..")
else
response.Write(rs("content"))
end if
%> </td>
        </tr>
        <tr> 
          <td colspan="2" align="right">__<%=rs("name")%></td>
        </tr>
        <tr> 
          <td colspan="2"><table width="100%" border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td height="3" bgcolor="#FFFFFF"><img src="1.gif" width="1" height="1"></td>
              </tr>
            </table></td>
        </tr>
      </table> </td>
  </tr>
</table>
        <%
rs.movenext
loop
rs.close
set rs=nothing
%>
      <!-- ========================================================================================================		  

													cup table ============stop

 ======================================================================================================== -->
