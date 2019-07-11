<%set rs=server.CreateObject("ADODB.RecordSet") 
rs.Source="select top 1 * from " & db_User_Table & " order by " & db_User_ID & " desc"
rs.Open rs.Source,ConnUser,1,1
if not rs.EOF then
%>
	&nbsp;&nbsp;○- 最近会员：<a href=user.asp?user=<%=left(rs(db_User_Name),8)%>  target="_blank" class=class><font color=red><%=rs(db_User_Name)%></font></a>
<%end if
rs.Close
set rs=nothing
%>
