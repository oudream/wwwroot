<%@ LANGUAGE="VBSCRIPT" %>
<!--#include file=config.asp -->
<html><head>
<meta http-equiv=Content-Language content=zh-cn>
<meta http-equiv=Content-Type content=text/html; charset=gb2312>
<meta http-equiv=Pragma content=no-cache>
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link href=style.css rel=STYLESHEET type=text/css>

<title>����.������ϵͳ</title>
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
<font color=red><b>IDΪ <%=request.querystring("id")%> �Ĺ������ʾ��¼</b></font>
<%
elseif request.querystring("ip")="cip" then
%>
<font color=red><b>IDΪ <%=request.querystring("id")%> �Ĺ���������¼</b></font>
<%
end if
%>
<hr color="#808080" size="1">
</td></tr></table>


<table border=1 width=100% cellspacing="0" cellpadding=0 style="border-collapse: collapse" bordercolor="#808080">
<tr bgcolor=#ffffff><td align="center" bgcolor="#F5F5F5" height="20">
��¼ID</td><td align=center bgcolor="#F5F5F5" height="20">IP ��ַ</td>
  <td align=center bgcolor="#F5F5F5" height="20">ʱ����</td></tr>
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
response.write "<tr align=center><td bgcolor=#ffffff colspan=3>û�м�¼</td></tr></table>"
else
adsrs.pagesize=32 'ÿҳ��ʾ�ļ�¼��
totalPut=adsrs.recordcount '��¼����
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
bookmark=adsrs.bookmark '�ƶ�����ʼ��ʾ�ļ�¼λ��
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
<tr bgcolor=#ffffff  align=center><td><font color=red><%=adsrs("id")%></font>��</td><td align=center><%=adsrs("ip")%>��</td><td align=center><%=adsrs("time")%>��</td></tr>
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
��<font color=red><%=totalput%></font>����ÿҳ<font color=red><%=adsrs.pagesize%></font>������<font color=red><%=currentPage%></font>ҳ����<font color=red><%=totalPage%></font>ҳ 
<%
response.write " ת��=>><select name='page' size=1>"
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