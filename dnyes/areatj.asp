<%
sql="select id,area,introt,photo from area where tuijian=true order by id" 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1
n=0
'预留分页显示功能
if not isempty(request("page")) then   
pagecount=cint(request("page"))   
else   
pagecount=1
end if

rs.pagesize=2
rs.AbsolutePage=pagecount

do while not rs.eof
%>
            
      <table width="570" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr><td colspan="3" height="10"></td></tr>
        <tr> 
          <td width="281" valign="top"> 
            <table width="270" border="0" cellspacing="0" cellpadding="0" align="center">
              <tr> 
                <td rowspan="3" valign="top" height="118" width="68">
<a href="showarea.asp?area=<%=rs("area")%>" title="<%=rs("area")%>" target="_blank"> 
              <img src="photo/<%=rs("photo")%>" border="0"></a></td>
                <td height="20" width="202">&nbsp;<a href="showarea.asp?area=<%=rs("area")%>" title="<%=rs("area")%>" target="_blank"><b><%=rs("area")%></b></a></td>
              </tr>
              <tr> 
                <td height="80" valign="top">
                  <table width="192" border="0" cellspacing="0" cellpadding="0" align="center">
                    <tr>
                      <td>
<%  
introt=replace(rs("introt"),chr(13),"<br>")
introt=replace(introt,chr(32),"&nbsp;")
response.write introt
%></td>
                    </tr>
                  </table>
                </td>
              </tr>
              <tr> 
                <td height="18">
                  <div align="right"><a href="showarea.asp?area=<%=rs("area")%>" title="<%=rs("area")%>" target="_blank"><img src="image/index_r17_c16.gif" width="44" height="9" border="0"></a>&nbsp;&nbsp;</div>
                </td>
              </tr>
            </table>
          </td>
          <td width="8" background="image/index_r13_c18.gif">&nbsp;</td>
<%
rs.movenext
if rs.eof or rs.bof then 
response.write "<td width='281'></td></tr><tr><tr><td colspan=3 height=10></td></tr><td colspan=3 height=1 bgcolor=#ADC3E7></td></tr></table>"
exit do
else
n=n+1
end if
%>
          <td width="281" valign="top"> 
            <table width="270" border="0" cellspacing="0" cellpadding="0" align="center">
              <tr> 
                <td height="20" width="202">&nbsp;<a href="showarea.asp?area=<%=rs("area")%>" title="<%=rs("area")%>" target="_blank"><b><%=rs("area")%></b></a></td>
                <td rowspan="3" valign="top" height="118" width="68">
<a href="showarea.asp?area=<%=rs("area")%>" title="<%=rs("area")%>" target="_blank"> 
                  <img src="photo/<%=rs("photo")%>" border="0"></a></td>
              </tr>
              <tr> 
                <td height="80" valign="top">
                  <table width="192" border="0" cellspacing="0" cellpadding="0" align="center">
                    <tr> 
                      <td> <%  
introt=replace(rs("introt"),chr(13),"<br>")
introt=replace(introt,chr(32),"&nbsp;")
response.write introt
%></td>
                    </tr>
                  </table>
                </td>
              </tr>
              <tr> 
                <td height="18">
                  <div align="right"><a href="showarea.asp?area=<%=rs("area")%>" title="<%=rs("area")%>" target="_blank"><img src="image/index_r17_c16.gif" width="44" height="9" border="0"></a>&nbsp;&nbsp;</div>
                </td>
              </tr>
            </table>
          </td>
        </tr>
        <tr><td colspan="3" height="10"></td></tr>
        <tr><td colspan="3" height=1 bgcolor="#ADC3E7"></td></tr>

      </table>
<%
rs.movenext
n=n+1                                                                     
if n>=rs.pagesize then exit do                                                           
loop

rs.close
set rs=nothing
%> 