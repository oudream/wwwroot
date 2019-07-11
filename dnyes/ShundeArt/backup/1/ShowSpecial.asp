              <%if showspecial=1 then%>
              <tr> 
                <td width="100%" height="24" valign="middle" > 
                  <p align="center"><FONT color=#ffffff>专题文章</FONT></td>
              </tr>
              <tr> 
                <td width=100% height=18> <%set rs2=server.CreateObject("ADODB.RecordSet")  '专题
rs2.Source="select Top " & top_sp & " * from "& db_Special_Table &"  order by SpecialID DESC "
rs2.Open rs2.Source,conn,1,1
if not rs2.EOF then
while not rs2.EOF

TrString="&nbsp;&nbsp;&nbsp;&nbsp;-&nbsp;<a class=class href='Special_News.asp?SpecialID=" & rs2("SpecialID") &"'>" & trim(rs2("SpecialName")) & "</a><br>"
Response.Write TrString

rs2.MoveNext
wend
%> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a class=class href=special.asp>更多专题...</a></td>
              </tr>
              <%
else
Response.Write "<tr><td width=100% align=center height=18 >暂无专题</td></tr>"
end if
rs2.Close
set rs2=nothing

  
%>
              <%end if%>