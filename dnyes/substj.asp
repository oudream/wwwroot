<%
sql="select * from subs where iftj=true order by id" 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1
n=0
'预留分页显示功能
if not isempty(request("page")) then   
pagecount=cint(request("page"))   
else   
pagecount=1
end if

rs.pagesize=4
rs.AbsolutePage=pagecount

do while not rs.eof
%>
        <table width="563" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td colspan="2" height="1" bgcolor="#3D5E7D"></td>
          </tr>
          <tr> 
            <td width="1" bgcolor="#3D5E7D"></td>
            <td> 
              <table width="562" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td rowspan="2" valign="top" width="277"> 
                    <table width="277" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td rowspan="4" width="120"> 
                          <div align="center"><img src=showimage.asp?id=<%=rs("id")%>&tu=stu border="0" width="68" height="98"></div>
                        </td>
                        <td width="157" height="10"></td>
                      </tr>
                      <tr> 
                        <td height="25">品名：<%=rs("subsname")%></td>
                      </tr>
                      <tr> 
                        <td height="25">价格：
<%if rs("ifdisc")=true then%>
<%response.write FormatNumber(csng(rs("price"))*session("y_it_userdiscount"),2)%>
<%else%>
<%response.write FormatNumber(csng(rs("price")),2)%>
<%end if%>
 RMB /<%=rs("subsunit")%></td>
                      </tr>
                      <tr> 
                        <td height="66" valign="top">简介：
<%if len(rs("other"))>25 then%> <%=left(rs("other"),22)%>...<%else%> 
<%=rs("other")%> <%end if%>
                        </td>
                      </tr>
                      <tr> 
                        <td colspan="2" height="24"> 
                          <table width="277" border="0" cellspacing="0" cellpadding="0">
                            <tr> 
                              <td width="170"><a href="#" onClick="window.open('add.asp?add=<%=rs("bookbm")%>','add','scrollbars=yes,resizable=no,width=650,height=450 top=10,left=10')"><img src="image/in1a.gif" width="170" height="24" border="0"></a></td>
                              <td width="107"><a href="showsubs.asp?id=<%=rs("id")%>"><img src="image/in1b.gif" width="107" height="24" border="0"></a></td>
                            </tr>
                          </table>
                        </td>
                      </tr>
                    </table>
                  </td>
                  <td width="8" height="126" background="image/index_r13_c18.gif">&nbsp;</td>
<%
rs.movenext
if rs.eof or rs.bof then 
response.write "<td rowspan=2 valign=top width=277></td></tr><tr><td height=24 background=image/inbgnew.gif>&nbsp;</td></tr></table></td></tr></table>"
exit do
else
n=n+1
end if
%>
                  <td rowspan="2" valign="top" width="277"> 
                    <table width="277" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td rowspan="4" width="120"> 
                          <div align="center"><img src=showimage.asp?id=<%=rs("id")%>&tu=stu border="0" width="68" height="98"></div>
                        </td>
                        <td width="157" height="10"></td>
                      </tr>
                      <tr> 
                        <td height="25">品名：<%=rs("subsname")%></td>
                      </tr>
                      <tr> 
                        <td height="25">价格：
<%if rs("ifdisc")=true then%>
<%response.write FormatNumber(csng(rs("price"))*session("y_it_userdiscount"),2)%>
<%else%>
<%response.write FormatNumber(csng(rs("price")),2)%>
<%end if%>
                          RMB /<%=rs("subsunit")%></td>
                      </tr>
                      <tr> 
                        <td height="66" valign="top">简介：<%if len(rs("other"))>25 then%> <%=left(rs("other"),22)%>...<%else%> 
<%=rs("other")%> <%end if%></td>
                      </tr>
                      <tr> 
                        <td colspan="2" height="24"> 
                          <table width="277" border="0" cellspacing="0" cellpadding="0">
                            <tr> 
                              <td width="170"><a href="#" onClick="window.open('add.asp?add=<%=rs("bookbm")%>','add','scrollbars=yes,resizable=no,width=650,height=450 top=10,left=10')"><img src="image/in2a.gif" width="170" height="24" border="0"></a></td>
                              <td width="107"><a href="showsubs.asp?id=<%=rs("id")%>"><img src="image/in2b.gif" width="107" height="24" border="0"></a></td>
                            </tr>
                          </table>
                        </td>
                      </tr>
                    </table>
                  </td>
                </tr>
                <tr> 
                  <td height="24" background="image/inbgnew.gif">&nbsp;</td>
                </tr>
              </table>
            </td>
          </tr>
        </table>
<%
rs.movenext
n=n+1                                                                     
if n>=rs.pagesize then exit do                                                           
loop
%> 
        