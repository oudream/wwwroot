<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="function.asp"-->
<%
dim shu,news,bigclassid,smallclassid,typeid,specialid,order,time,path,click
bigclassid=checkstr(request.querystring("bigclassid"))
typeid=checkstr(request.querystring("typeid"))
smallclassid=checkstr(request.querystring("smallclassid"))
specialid=checkstr(request.querystring("specialid"))
shu=checkstr(request.querystring("shu"))
order=checkstr(request("order"))
time=checkstr(request.querystring("time"))
title=checkstr(request.querystring("title"))
click=checkstr(request.querystring("click"))
Path =  "./"
if shu<>"" then   '显示数目，不加为10
ss=shu
else
ss=10
end if
if order="click" then  '排序方式，click为按点击率
oo="click"
else
oo="newsid"
end if
if time="0" then  '显示时间，0为不显示
tt=0
else
tt=1
end if
if click="0" then  '显示点击数，0为不显示
cc=0
else
cc=1
end if
if title<>"" then   '显示数目，不加为10
nn=title
else
nn=10
end if
if typeid<>"" then
	set rs=conn.execute("SELECT top "&ss&" * FROM "& db_News_Table &"  where typeid=cint("&typeid&") and checkked=1  order by "&oo&" desc")
else
if bigclassid<>"" then
	set rs=conn.execute("SELECT top "&ss&" * FROM "& db_News_Table &"  where bigclassid=cint("&bigclassid&") and checkked=1  order by "&oo&" desc")
else
if smallclassid<>"" then
	set rs=conn.execute("SELECT top "&ss&" * FROM "& db_News_Table &"  where smallclassid=cint("&smallclassid&") and checkked=1  order by "&oo&" desc")
else
if specialid<>"" then
	set rs=conn.execute("SELECT top "&ss&" * FROM "& db_News_Table &"  where specialid=cint("&specialid&") and checkked=1  order by "&oo&" desc")
else
	set rs=conn.execute("SELECT top "&ss&" * FROM "& db_News_Table &" where checkked=1  order by "&oo&" desc")
end if
end if
end if
end if
if rs.eof and rs.bof then %>
     document.write('<div align="center">没有</div>');
  <% else %>
      document.write('<table border=0 cellpadding=0 cellspacing=3 width=88 height=120 ><tr>');
<%do while not rs.eof
if (rs("picnews"))=1 then
fileExt=lcase(getFileExtName(rs("picname")))
title=htmlencode4(trim(rs("title")))
%>
document.write('<td align=center><A class=class href="<%=path%>ReadNews.asp?NewsID=<%=rs("NewsID")%>" title="<%=title%>" target="_blank"><img width="160" height="120" src="<%=FileUploadPath%><%=rs("picname")%>" border=0><br><%=CutStr(title,nn)%></A></td>');			
<%else
fileExt=lcase(getFileExtName(rs("picname")))
title=trim(rs("title"))
title=replace(title,"<br>","")%>

document.write('<td align=center><A class=class href="<%=path%>ReadNews.asp?NewsID=<%=rs("NewsID")%>" title="<%=title%>" target="_blank"><img  width="160" height="120" src="IMAGES/flashorno.gif"  border=0><br><%=CutStr(title,nn)%></A></td>');


<%end if
rs.movenext 
loop 
%>
document.write('</tr>');
<%end if
Rs.Close
set Rs=nothing
%>

<%
function getFileExtName(fileName)
dim pos
pos=instrrev(filename,".")
if pos>0 then
getFileExtName=mid(fileName,pos+1)
else
getFileExtName=""
end if
end function
%>
document.write('</table>');