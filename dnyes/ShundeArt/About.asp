<%@ Language=VBScript%>
<!--#include file=conn.asp -->
<!--#include file=config.asp -->
<!--#include file="function.asp"-->
<!--#include file="char.inc"-->
<%
set rs=server.CreateObject("ADODB.RecordSet")
rs.Source="select * from "& db_Type_Table &" order by typeorder"
rs.Open rs.Source,conn,1,1

dim ArraytypeID(10000),ArraytypeName(10000),Arraytypecontent(10000)
typeCount=rs.RecordCount
for i=1 to typeCount
ArraytypeID(i)=rs("typeID")
ArraytypeName(i)=rs("typeName")
Arraytypecontent(i)=rs("typecontent")
rs.MoveNext
next
rs.Close
PageShowSize = 10            '每页显示多少个页
MyPageSize   = 20           '每页显示多少条新闻
	
If Not IsNumeric(Request("page")) Or IsEmpty(Request("page")) Or Request("page") <=0 Then
MyPage=1
Else
MyPage=Int(Abs(Request("page")))
End if
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>顺德国际艺术交流中心_记事</title>
<LINK href=news.css rel=stylesheet>
</head>

<body topmargin="0">
<!--#include file=top.asp -->
<table width="760" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <tr valign="top"> 
    <td width="180" background="IMAGES/menu-d.gif"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td valign="top"><table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="100%" id="AutoNumber2" height="23">
              <tr> 
                <td width="100%" height="25" valign="middle" background="IMAGES/menu-d-fl.gif"> 
                  <p align="center"><font color="#000000">文章排行</font></td>
              </tr>
              <tr> 
                <td width="100%" background="IMAGES/menu-d.gif"><div align="center"> 
                    <table width="99%" border="0" cellspacing="0" cellpadding="3">
                      <tr> 
                        <td> 
                          <%
set rs=server.CreateObject("ADODB.RecordSet")
if uselevel=1 then
if Request.cookies(Forcast_SN)("key")="super" or Request.cookies(Forcast_SN)("key")="typemaster" or Request.cookies(Forcast_SN)("key")="bigmaster" or Request.cookies(Forcast_SN)("key")="smallmaster" or Request.cookies(Forcast_SN)("key")="check" then
rs.Source="select top "& top_txt &" * from "& db_News_Table &" where  checkked=1 order by click DESC,newsid desc"
end if
if Request.cookies(Forcast_SN)("key")="" then
rs.Source="select top "& top_txt &" * from "& db_News_Table &" where (checkked=1 and newslevel=0) order by click DESC,newsid desc"
end if
if Request.cookies(Forcast_SN)("key")="selfreg" then
	if Request.cookies(Forcast_SN)("reglevel")=3 then
rs.Source="select top "& top_txt &" * from "& db_News_Table &" where (checkked=1 ) order by click DESC,newsid desc"
	end if
	if Request.cookies(Forcast_SN)("reglevel")=2 then
rs.Source="select top "& top_txt &" * from "& db_News_Table &" where (checkked=1 ) order by click DESC,newsid desc"
	end if
	if Request.cookies(Forcast_SN)("reglevel")=1 then
rs.Source="select top "& top_txt &" * from "& db_News_Table &" where (checkked=1 ) order by click DESC,newsid desc"
	end if
end if
else
rs.Source="select top "& top_txt &" * from "& db_News_Table &" where  checkked=1 order by click DESC,newsid desc"
end if
rs.Open rs.Source,conn,1,1
while not rs.EOF
title=htmlencode4(trim(rs("title")))

%>
                          &nbsp;· <a class=middle href="ReadNews.asp?NewsID=<%=rs("NewsID")%>" target="_blank" title="<%=title%>"><font color="<%=rs("titlecolor")%>">                           
                          <%=CutStr(title,18)%> 
                          </font></a>&nbsp;<font color=red><%=rs("click")%></font><br> 
                          <%
rs.MoveNext
wend
rs.close
%>
                        </td>
                      </tr>
                    </table>
                  </div></td>
              </tr>
            </table></td>
        </tr>
        <tr bordercolor="#999999"> 
          <td height="25" align="center" bordercolor="#999999" background="IMAGES/menu-d-hot.gif"><FONT color=#000000>本月热门</FONT></td>
        </tr>
        <tr bordercolor="#999999"> 
          <td height="36" align="center" valign=top background="IMAGES/menu-d.gif"> 
            <br> <table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="93%" id="AutoNumber5">
              <% 
 dim ii
ii = 0
if uselevel=1 then
if Request.cookies(Forcast_SN)("key")="" then
rs.Source="select top " & top_txt & " * from "& db_News_Table &" where year(updatetime)=year(now()) and month(updatetime)=month(now()) and checkked=1 and newslevel=0 order by click DESC"   '选择本月
end if
if Request.cookies(Forcast_SN)("key")="selfreg" then
	if Request.cookies(Forcast_SN)("reglevel")="3" then
	rs.Source="select top " & top_txt & " * from "& db_News_Table &" where year(updatetime)=year(now()) and month(updatetime)=month(now()) and checkked=1 order by click DESC"   '选择本月
	end if
	if Request.cookies(Forcast_SN)("reglevel")="2" then
	rs.Source="select top " & top_txt & " * from "& db_News_Table &" where year(updatetime)=year(now()) and month(updatetime)=month(now()) and checkked=1 order by click DESC"   '选择本月
	end if
	if Request.cookies(Forcast_SN)("reglevel")="1" then
	rs.Source="select top " & top_txt & " * from "& db_News_Table &" where year(updatetime)=year(now()) and month(updatetime)=month(now()) and checkked=1 order by click DESC"   '选择本月
	end if
end if
if Request.cookies(Forcast_SN)("key")="super" or Request.cookies(Forcast_SN)("key")="typemaster" or Request.cookies(Forcast_SN)("key")="bigmaster" or Request.cookies(Forcast_SN)("key")="smallmaster" or Request.cookies(Forcast_SN)("key")="check" then
rs.Source="select top " & top_txt & " * from "& db_News_Table &" where year(updatetime)=year(now()) and month(updatetime)=month(now()) and checkked=1 order by click DESC"   '选择本月
end if
else
rs.Source="select top " & top_txt & " * from "& db_News_Table &" where year(updatetime)=year(now()) and month(updatetime)=month(now()) and checkked=1 order by click DESC"   '选择本月
end if
rs.Open rs.Source,conn,1,1
	if rs.bof and rs.eof then 
		response.write "<td align=center>无</td>" 
	else 
	do while not rs.eof 
%>
              <tr> 
                <td height=18> · 
                  <%if rs("picnews")=1 then%>
                  <img src="images/news_img.gif"> 
                  <%end if%>
                  <a class=middle href="ReadNews.asp?NewsID=<%=rs("NewsID")%>" title="<%=htmlencode4(rs("title"))%>" target="_blank" > 
                  <%=htmlencode4(rs("title"))%></a><font color="#666666">[<%=rs("click")%>]</font> 
                </td>
              </tr>
              <%  ii = ii + 1
    if ii>50 then exit do
	rs.movenext     
	loop
	end if  
	rs.close   
	set rs=nothing
%>
            </table></td>
        </tr>
        <tr> 
          
        </tr>
      </table>
    </td>
    <td width="6" bgcolor="#FFFFFF">　</td>
    <td background="IMAGES/menu-l.gif"> 
      <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td height="25" background="IMAGES/menu-l-m.gif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font class=m_tittle>&nbsp;</font>栏目导航&nbsp;<b><a class="daohang" href="./" > 
            关于我们</a></b></td>
        </tr>
        <tr> 
          <td background="IMAGES/menu-l.gif">&nbsp;</td>
        </tr>
        <tr> 
          
        </tr>
      </table>
      <TABLE height=38 cellSpacing=0 cellPadding=0 width=100% border=0>
        <TBODY>
          <TR> 
            <TD vAlign=bottom align=right width=8></TD>
            <TD height="500" align=center vAlign=middle><font size="3">这里要把公司的背景、优势、服务、。。。。都写进去。还要图文并茂！<br>
              <br>
              烦请准备一下。</font></TD>
            <TD width=10>&nbsp;</TD>
          </TR>
      </TABLE></td>
  </tr>
  <tr valign="top"> 
    <td height="20" background="IMAGES/menu-d-t.gif">&nbsp;</td>
    <td height="20" bgcolor="#FFFFFF">&nbsp;</td>
    <td height="20" background="IMAGES/menu-l-t.gif" class="daohang">&nbsp;</td>
  </tr>
</table>
<!--#include file=bottom.asp -->

</body>

</html>