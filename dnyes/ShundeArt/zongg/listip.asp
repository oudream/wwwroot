<%@ LANGUAGE="VBSCRIPT" %>
<!--#include file=config.asp -->
<html><head>
<meta http-equiv=Content-Language content=zh-cn>
<meta http-equiv=Content-Type content=text/html; charset=gb2312>
<meta http-equiv=Pragma content=no-cache>
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link href=style.css rel=STYLESHEET type=text/css>

<title>忠网.广告管理系统</title>
<base target=_top>
</head><body marginwidth=0 marginheight=0>


<BR><BR>
<%call  Iflogin()%>
 
 <div align=center><center>
<table border=0 width=100% cellspacing=0 cellpadding=0>
<tr>
<td bgcolor=#000000>
<div align="center">
  <center>
<table border=0 width=100% cellspacing=0 cellpadding=3 style="border-collapse: collapse" bordercolor="#111111">
<tr><td align=center bgcolor=#ffffff>
<%
if request.querystring("ip")="sip" then
%>
<font color=red><b>ID为 <%=request.querystring("id")%> 的广告条显示记录</b></font>
<%
elseif request.querystring("ip")="cip" then
%>
<font color=red><b>ID为 <%=request.querystring("id")%> 的广告条点击记录</b></font>
<%
end if
%>
<hr color="#808080" size="1">
</td></tr></table>


<table border=1 width=100% cellspacing="0" cellpadding=0 style="border-collapse: collapse" bordercolor="#808080">
<tr bgcolor=#ffffff><td align="center" bgcolor="#F5F5F5" height="20">
记录ID</td><td align=center bgcolor="#F5F5F5" height="20">IP 地址</td>
  <td align=center bgcolor="#F5F5F5" height="20">时　间</td></tr>
<%

adsconn.Open adsdata
dim MaxPerPage,adssql,adsrs,totalPut,CurrentPage,TotalPages,i,advlistact
if not isempty(request("page")) then
currentPage=cint(request("page"))
else
currentPage=1
end if
set adsrs=server.createobject("adodb.recordset")

if request.querystring("ip")="sip" then
getadid=cint(request("id"))
adssql="select * from iplist where adid="&getadid&" and class=1 order by id desc"

elseif request.querystring("ip")="cip" then
getadid=cint(request("id"))
adssql="select * from iplist where adid="&getadid&" and class=2 order by id desc"
end if

adsrs.open adssql,adsconn,1,1
if adsrs.eof and adsrs.bof then
response.write "<tr align=center><td bgcolor=#ffffff colspan=3>没有记录</td></tr></table>"
else
adsrs.pagesize=32 '每页显示的记录数
totalPut=adsrs.recordcount '记录总数
totalPage=adsrs.pagecount
MaxPerPage=adsrs.pagesize
if currentpage<1 then
currentpage=1
end if
if currentpage>totalPage then
currentpage=totalPage
end if
if currentPage=1 then
showContent
showpages
else
if (currentPage-1)*MaxPerPage<totalPut then
adsrs.move  (currentPage-1)*MaxPerPage
dim bookmark
bookmark=adsrs.bookmark '移动到开始显示的记录位置
showContent
showpages
end if
end if
adsrs.close
set adsrs=nothing
end if
adsconn.close
set adsconn=nothing

%>


  </center>
</div>






<%
sub showContent
i=0
do while not (adsrs.eof or err)
%>
<tr bgcolor=#ffffff  align=center><td><font color=red><%=adsrs("id")%></font>　</td><td align=center><%=adsrs("ip")%>　</td><td align=center><%=adsrs("time")%>　</td></tr>
<%
i=i+1
if i>=MaxPerPage then exit do 
adsrs.movenext
loop
end sub 

sub showpages()
dim n
n=totalPage
%>
</table>
<table border=1 width=100% cellpadding=0 cellspacing=0 bordercolor="#808080" style="border-collapse: collapse">
<form method=post action=listip.asp?id=<%=trim(request("id"))%>&ip=<%=trim(request("ip"))%>><tr bgcolor=#ffffff><td align=right colspan=4>
共<font color=red><%=totalput%></font>条，每页<font color=red><%=adsrs.pagesize%></font>条，第<font color=red><%=currentPage%></font>页，共<font color=red><%=totalPage%></font>页 
<%
response.write " 转到=>><select name='page' size=1>"
for i=1 to n
response.write "<option value="& i
if currentpage=i then
response.write " selected"
end if
response.write ">"& i &"</option>"
next
response.write "</select>&nbsp;<input type='submit'  value=' go '>"
%>
</td></tr></form>
</table>
<%
end sub
%></div>

</body></html>