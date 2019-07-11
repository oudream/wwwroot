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
            记事</a>&gt;&gt;内容</b></td>
        </tr>
        <tr> 
          <td background="IMAGES/menu-l.gif"><div align="center"> </b> </div></td>
        </tr>
        <tr> 
          <td background="IMAGES/menu-l.gif"> <table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="f5f5f5" id="AutoNumber3" style="border-collapse: collapse">
              <tr> 
                <td width="50" valign=top background="IMAGES/menu-l.gif" bgcolor="f2f2f2">作家</td>
                <td width="100" valign=top background="IMAGES/menu-l.gif" bgcolor="f2f2f2">作品名称</td>
                <td width="100" valign=top background="IMAGES/menu-l.gif" bgcolor="f2f2f2">尺寸(cm)</td>
                <td width="80" valign=top background="IMAGES/menu-l.gif" bgcolor="f2f2f2">价格(元)</td>
                <td width="50" valign=top background="IMAGES/menu-l.gif" bgcolor="f2f2f2">售出情况</td>
                <td valign=top background="IMAGES/menu-l.gif" bgcolor="f2f2f2">备注</td>
              </tr>
              <tr> 
                <td height="16" valign=top bgcolor="#FFFFFF">啊日</td>
                <td valign=top bgcolor="#FFFFFF"><a href="Show_Me.asp" target="_blank">啊日的裸体画</a></td>
                <td valign=top bgcolor="#FFFFFF">10000 x 10000</td>
                <td valign=top bgcolor="#FFFFFF">- 全幅身家</td>
                <td valign=top bgcolor="#FFFFFF">未画</td>
                <td valign=top bgcolor="#FFFFFF">留给有缘人@!@</td>
              </tr>
            </table></td>
        </tr>
        <tr> 
          
        </tr>
      </table>
    </td>
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