<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<!--#include file="function.asp"-->
<!--#include file="char.inc"-->
<%
dim shu,news,bigclassid,smallclassid,specialid,order,time,path,click,title
bigclassid=checkstr(request.querystring("bigclassid"))
typeid=checkstr(request.querystring("typeid"))
smallclassid=checkstr(request.querystring("smallclassid"))
specialid=checkstr(request.querystring("specialid"))
shu=checkstr(request.querystring("shu"))
order=checkstr(request("order"))
time=checkstr(request.querystring("time"))
title=checkstr(request.querystring("title"))
click=checkstr(request.querystring("click"))
Path = "./"
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
nn=28
end if
if typeid<>"" then
	set rs=conn.execute("SELECT top "&ss&" * FROM "& db_News_Table &"  where typeid=cint("&typeid&") and checkked=1 order by "&oo&" desc")
else
if bigclassid<>"" then
	set rs=conn.execute("SELECT top "&ss&" * FROM "& db_News_Table &"  where bigclassid=cint("&bigclassid&") and checkked=1 order by "&oo&" desc")
else
if smallclassid<>"" then
	set rs=conn.execute("SELECT top "&ss&" * FROM "& db_News_Table &"  where smallclassid=cint("&smallclassid&") and checkked=1 order by "&oo&" desc")
else
if specialid<>"" then
	set rs=conn.execute("SELECT top "&ss&" * FROM "& db_News_Table &"  where specialid=cint("&specialid&") and checkked=1 order by "&oo&" desc")
else
	set rs=conn.execute("SELECT top "&ss&" * FROM "& db_News_Table &" where checkked=1 order by "&oo&" desc")
end if
end if
end if
end if
if rs.eof and rs.bof then %>
     document.write('没有');
  <% else  
do while not rs.eof

title=htmlencode4(trim(rs("title")))
if strcomp(CutStr(title,nn),title)<0 then
%>
document.write('<LINK href=news.css rel=stylesheet><li><font class=class><A class=class href="<%=path%>ReadNews.asp?NewsID=<%=rs("NewsID")%>" title="<%=title%>" target="_blank"><%=CutStr(title,nn)%></A><%if tt=1 then%><FONT color="#666666">(<%=month(rs("updatetime"))%>-<%=day(rs("updatetime"))%>)</FONT><%end if%><%if date()-rs("updatetime")<2 then%><FONT color="#ff0000">新！</FONT><%end if%><%if cc=1 then%><FONT color="#666666">(<%=rs("click")%>)</FONT><%end if%></font>');
<%else%>
document.write('<LINK href=news.css rel=stylesheet><li><font class=class><A class=class href="<%=path%>ReadNews.asp?NewsID=<%=rs("NewsID")%>" title="<%=title%>" target="_blank"><%=CutStr(title,nn)%></A><%if tt=1 then%><FONT color="#666666">(<%=month(rs("updatetime"))%>-<%=day(rs("updatetime"))%>)</FONT><%end if%><%if date()-rs("updatetime")<2 then%><FONT color="#ff0000">新！</FONT><%end if%><%if cc=1 then%><FONT color="#666666">(<%=rs("click")%>)</FONT><%end if%></font>');
<%
end if
rs.movenext 
loop 
end if
Rs.Close
set Rs=nothing
%>
