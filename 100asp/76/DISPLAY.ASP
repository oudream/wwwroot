<%
dim rs
dim sqlan
dim sela
dim selb
dim selc
dim total
conn = "DBQ=" + server.mappath("search.mdb") + ";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
set rs=server.createobject("adodb.recordset")
sqlan="select * from search where id=1"
rs.open sqlan,conn,1,1
if not rs.eof then
total=rs("sela")+rs("selb")+rs("selc")
sela=(rs("sela")/total)*100
selb=(rs("selb")/total)*100
selc=(rs("selc")/total)*100
%>
<div align="center">
<center>
<P>您认为明年的CBA哪一支队伍能够夺得总冠军？<p>
<table width="60%"><tr></td>
<p>八一队：<img src=bar.gif width=<%=int(sela*5)%> height=10> <%=rs("sela")%>人 占：<%=round(sela,2)%>%</p>
<p>上海队：<img src=bar.gif width=<%=int(sela*5)%> height=10><%=rs("selb")%>人 占：<%=round(selb,2)%>%</p>
<p>说不清：<img src=bar.gif width=<%=int(selc*5)%> height=10><%=rs("selc")%>人 占：<%=round(selc,2)%>%</p>
</td></tr></table></center></div>
<p align=center>共有：<%=total%>人参加调查
<%
else
response.write "<script language=JavaScript>" & chr(13) & "alert('还没有人参与调查!');" & "history.back()" & "</script>" 
end if
rs.close
set rs=nothing
%>
