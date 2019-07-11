<table border=0 cellpadding=0 cellspacing=3 width="88" align="center">
 <tr>
<%
set rs3=server.CreateObject("ADODB.RecordSet")
if uselevel=1 then
if session("key")="super" or session("key")="typemaster" or session("key")="bigmaster" or session("key")="smallmaster" or session("key")="check" then
rs3.Source ="select top " & top_img & " * from "& db_News_Table &" where picnews=1 and checkked=1 order by NewsID DESC"
end if
if session("key")="" then
rs3.Source ="select top " & top_img & " * from "& db_News_Table &" where picnews=1 and checkked=1 order by NewsID DESC"
end if
if session("key")="selfreg" then
	if session("reglevel")=3 then
rs3.Source ="select top " & top_img & " * from "& db_News_Table &" where picnews=1 and checkked=1 order by NewsID DESC"
	end if
	if session("reglevel")=2 then
rs3.Source ="select top " & top_img & " * from "& db_News_Table &" where picnews=1 and checkked=1 order by NewsID DESC"
	end if
	if session("reglevel")=1 then
rs3.Source ="select top " & top_img & " * from "& db_News_Table &" where picnews=1 and checkked=1 order by NewsID DESC"
	end if
end if
else
rs3.Source ="select top " & top_img & " * from "& db_News_Table &" where picnews=1 and checkked=1 order by NewsID DESC"
end if
rs3.Open rs3.Source,conn,1,1
if not rs3.EOF then
while not rs3.EOF
%>
 <%if rs3("picname")<>"" then%>
    <td height=31 align="center"><a href="readnews.asp?newsid=<%=rs3("newsid")%>" target="_blank" title="<%=rs3("title")%>"><img src="<%=rs3("picname")%>" width="160" height="120" border="0"></a></td>
<%end if
rs3.MoveNext
wend
%>
</tr>
<%
else
Response.Write "目前还没有"
end if
rs3.Close
set rs3=nothing
%>

</table>