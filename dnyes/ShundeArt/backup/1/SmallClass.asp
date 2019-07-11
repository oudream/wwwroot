<%@ Language=VBScript%>

<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="function.asp"-->
<%
typeid=request("typeid")
request_BigClassID=Request.QueryString("BigClassID")
request_SmallClassID=Request.QueryString ("SmallClassID")
if typeid="" or request_bigclassid="" or request_smallclassid="" then
Response.Write "<script>alert('未指定参数');history.back()</script>"
response.end
else
if not IsNumeric(request_SmallClassID) or not IsNumeric(request_bigclassid) or not IsNumeric(typeid) then
 response.write "<script>alert('非法参数');history.back()</script>"
 response.end
else
set rs=server.CreateObject("ADODB.RecordSet")
rs.Source="select * from "& db_BigClass_Table &" where bigclassid=" & request_BigClassID &" order by BigClassorder"
rs.Open rs.Source,conn,1,1
bigclassname=rs("bigclassname")
rs.close
set rs=nothing
set rs=server.CreateObject("ADODB.RecordSet")
rs.Source="select * from "& db_SmallClass_Table &" where smallclassid=" & request_smallClassID &" order by smallClassorder"
rs.Open rs.Source,conn,1,1
smallclassname=rs("smallclassname")
rs.close
set rs=nothing

PageShowSize = 10            '每页显示多少个页
MyPageSize   = 20           '每页显示多少条新闻
	
If Not IsNumeric(Request("page")) Or IsEmpty(Request("page")) Or Request("page") <=0 Then
MyPage=1
Else
MyPage=Int(Abs(Request("page")))
End if
end if


set rs=server.CreateObject("ADODB.RecordSet")
rs.Source="select * from "& db_BigClass_Table &" where typeid="&typeid&" order by BigClassorder"
rs.Open rs.Source,conn,1,1

dim ArrayBigClassID(10000),ArrayBigClassName(10000)
BigClassCount=rs.RecordCount
for i=1 to BigClassCount
ArrayBigClassID(i)=rs("BigClassID")
ArrayBigClassName(i)=trim(rs("BigClassName"))
rs.MoveNext
next


set rs2=server.CreateObject("ADODB.RecordSet")  '专题
rs2.Source="select Top "& top_sp &" * from "& db_Special_Table &"  order by SpecialID DESC "
rs2.Open rs2.Source,conn,1,1


set rs4=server.CreateObject("Adodb.RecordSet")
rs4.source="select * from "& db_SmallClass_Table &" Where BigClassid=" & request_BigClassid &" order by SmallClassorder"
rs4.open rs4.source,conn,1,1

SmallClassCount=rs4.RecordCount
dim ArraySmallClassID(10000),ArraySmallClassName(10000)
for i=1 to SmallClassCount
ArraySmallClassID(i)=rs4("SmallClassID")
ArraySmallClassName(i)=trim(rs4("SmallClassName"))
rs4.MoveNext
next
rs4.Close
set rs4=nothing


%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=SmallClassName%>__<%=BigClassName%>__<%=jjgn%></title>
<LINK href=news.css rel=stylesheet>
</head>

<body topmargin="0">
<!--#include file=top.asp -->
<table width="760" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <tr valign="top"> 
    <td width="180" height="412" background="IMAGES/menu-d.gif"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td valign="top" bgcolor="efefef"></td>
        </tr>
        <tr> 
          <td height="384" valign="top" background="IMAGES/menu-d.gif"> 
            <table width="100%" border="0" cellpadding="3" cellspacing="0" bgcolor="#FFFFFF">
              <tr bordercolor="#999999"> 
                <td height="25" align="center" valign="middle" bordercolor="#999999" background="IMAGES/menu-d-fl.gif"><font color="#000000">分类栏目</font></td>
              </tr>
              <tr> 
                <td width="30%" height="23" align="center" valign="top" background="IMAGES/menu-d.gif"> 
                  <table width="92%" border="0" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td> 
                        <%
    dim mode
    typeid=checkstr(request("typeid"))
    set rs=server.CreateObject("ADODB.RecordSet")
rs.Source="select * from "& db_Type_Table &" where typeid="&typeid&" order by typeorder"
rs.Open rs.Source,conn,1,1
mode=rs("mode")
rs.close
set rs=nothing
    if mode=5 then
    set rs=conn.execute("SELECT * FROM "& db_BigClass_Table &" where typeid="&typeid&" order by bigclassorder") 
    if not rs.EOF then
	do while not rs.eof

    %>
                        <table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="100%" id="AutoNumber2">
                          <tr> 
                            <td colspan="2"> <li></li>
                              <%if not Rs.eof then%>
                              <font class=m_tittle>&nbsp;</font><a class=middle href="bigclass.asp?typeid=<%=typeid%>&bigclassid=<%=Rs("bigclassid")%>"><font color="#000000"><%=Rs("bigclassname")%></font></a>&nbsp;</td>
                          </tr>
                          <%set nrs=conn.execute("SELECT * FROM "& db_SmallClass_Table &" where bigclassid="&cstr(rs("bigclassid"))&" order by smallclassorder")%>
                          <%do while not nrs.eof%>
                          <tr> 
                            <td width=100% align=center valign="middle"> 
                              <%if not nRs.eof then%>
                              <a class=class href="SmallClass.asp?typeid=<%=typeid%>&BigClassID=<%=nrs("bigclassid")%>&smallclassid=<%=nrs("smallclassid")%>"><%=nRs("smallclassname")%></a>　</td>
                            <%nrs.movenext   
		     end if %>
                          </tr>
                          <%loop
            nRs.Close
           set nRs=nothing
            %>
                        </table>
                        <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="<%=border%>" width="105%" id="AutoNumber2" height="11">
                          <tr> 
                            <td height=1 align=center></td>
                          </tr>
                        </table>
                        <%rs.movenext   
	end if
	  loop
	rs.close
%>
                        <%else%>
                        <table cellpadding=3 cellspacing=0 width=100% border=0 style="border-collapse: collapse">
                          <tr> 
                            <td height="22">&nbsp;&nbsp;&nbsp;&nbsp;<font color="#000000">暂无大类</font>&nbsp;</td>
                          </tr>
                        </table>
                        <%end if%>
                        <%else%>
                        <% typeid=checkstr(request("typeid"))
	set rs=conn.execute("SELECT * FROM "& db_BigClass_Table &" where typeid="&typeid&" order by bigclassorder") 
	do while not rs.eof
%>
                        <table width="100%" border="0" cellpadding="3" cellspacing="0" id="AutoNumber2" style="border-collapse: collapse">
                          <tr> 
                            <td height=22 colspan="2"> <li></li>
                              <%if not Rs.eof then%>
                              <a class=middle href="bigclass.asp?typeid=<%=typeid%>&bigclassid=<%=Rs("bigclassid")%>"><font color="#000000"><%=Rs("bigclassname")%></font></a></td>
                          </tr>
                          <%set nrs=conn.execute("SELECT * FROM "& db_SmallClass_Table &" where bigclassid="&cstr(rs("bigclassid"))&" order by smallclassorder")%>
                          <%do while not nrs.eof%>
                          <tr> 
                            <td width=50% align=center valign="middle"> 
                              <%if not nRs.eof then%>
                              <a class=class href="SmallClass.asp?typeid=<%=typeid%>&BigClassID=<%=nrs("bigclassid")%>&smallclassid=<%=nrs("smallclassid")%>"><%=nRs("smallclassname")%></a>　</td>
                            <%nrs.movenext   
		     end if %>
                            <td width=50% align=center valign="middle"> 
                              <%if not nRs.eof then%>
                              <a class=class href="SmallClass.asp?typeid=<%=typeid%>&BigClassID=<%=nrs("bigclassid")%>&smallclassid=<%=nrs("smallclassid")%>"><%=nRs("smallclassname")%></a>　</td>
                            <%nRs.movenext
             end if%>
                          </tr>
                          <%loop
            nRs.Close
           set nRs=nothing
            %>
                        </table>
                        <%rs.movenext   
	end if
	loop  
	rs.close
%>
                        <%end if%>
                      </td>
                    </tr>
                  </table></td>
              </tr>              
              <tr> 
                <td height="25" valign="middle" background="IMAGES/menu-d-new.gif"> 
                  <p align="center"><font color="#000000">本月最新</font></td>
              </tr>
              <tr> 
                <td height="24" align="center" background="IMAGES/menu-d.gif"> 
                  <table width="96%" border="0" cellpadding="3" cellspacing="0">
                    <tr> 
                      <td> 
                        <%  dim ii
	ii = 0
	if uselevel=1 then
if Request.cookies(Forcast_SN)("key")="" then
rs.Source="select top " & top_txt & " * from "& db_News_Table &" where year(updatetime)=year(now()) and month(updatetime)=month(now()) and checkked=1 and typeid="&typeid&"  order by click DESC,newsid desc"   '选择本月
end if
if Request.cookies(Forcast_SN)("key")="selfreg" then
	if Request.cookies(Forcast_SN)("reglevel")=3 then
	rs.Source="select top " & top_txt & " * from "& db_News_Table &" where year(updatetime)=year(now()) and month(updatetime)=month(now()) and checkked=1 and typeid="&typeid&" order by click DESC,newsid desc"   '选择本月
	end if
	if Request.cookies(Forcast_SN)("reglevel")=2 then
	rs.Source="select top " & top_txt & " * from "& db_News_Table &" where year(updatetime)=year(now()) and month(updatetime)=month(now()) and checkked=1 and typeid="&typeid&" order by click DESC,newsid desc"   '选择本月
	end if
	if Request.cookies(Forcast_SN)("reglevel")=1 then
	rs.Source="select top " & top_txt & " * from "& db_News_Table &" where year(updatetime)=year(now()) and month(updatetime)=month(now()) and checkked=1 and typeid="&typeid&" order by click DESC,newsid desc"   '选择本月
	end if
end if
if Request.cookies(Forcast_SN)("key")="super" or Request.cookies(Forcast_SN)("key")="typemaster" or Request.cookies(Forcast_SN)("key")="bigmaster" or Request.cookies(Forcast_SN)("key")="smallmaster" or Request.cookies(Forcast_SN)("key")="check" then
rs.Source="select top " & top_txt & " * from "& db_News_Table &" where year(updatetime)=year(now()) and month(updatetime)=month(now()) and checkked=1 and typeid="&typeid&" order by click DESC,newsid desc"   '选择本月
end if
else
rs.Source="select top " & top_txt & " * from "& db_News_Table &" where year(updatetime)=year(now()) and month(updatetime)=month(now()) and checkked=1 and typeid="&typeid&" order by click DESC,newsid desc"   '选择本月
end if
rs.Open rs.Source,conn,1,1
	if rs.bof and rs.eof then 
		response.write "<br><div align=center>本月无更新<div>" 
	else 
	do while not rs.eof 
title=htmlencode4(trim(rs("title")))
%>
                        &nbsp;· 
                        <%
    typeid=checkstr(request("typeid"))
    set rs11=server.CreateObject("ADODB.RecordSet")
rs11.Source="select * from "& db_Type_Table &" where typeid="&typeid&" order by typeorder"
rs11.Open rs11.Source,conn,1,1
mode=rs11("mode")
rs11.close
set rs11=nothing%>
                        <%if mode<>4 then%>
                      
                        <%if mode="2" then%>
                        <%if rs("picname")<>"" then%>
                        <img src='images/news_img.gif' border='0'> 
                        <%end if%>
                        <%end if%>
						  <a class=middle href="ReadNews.asp?NewsID=<%=rs("NewsID")%>" title="<%=title%>" target="_blank"> 
                        <font color="<%=rs("titlecolor")%>"><%=gottopic(title,14)%></font></a><font color="#FF0000"> 
                        <%=rs("click")%></font> 
                        <%else%>
                         
                        <%if mode="2" then%>
                        <%if rs("picname")<>"" then%>
                        <img src='images/news_img.gif' border='0'> 
                        <%end if%>
                        <%end if%>
						<a class=middle href="ReadNews.asp?NewsID=<%=rs("NewsID")%>" title="<%=title%>" target="_blank">
                        <font color="<%=rs("titlecolor")%>"><%=gottopic(title,12)%></font></a><font color="#FF0000"> 
                        <%=rs("click")%></font> 
                        <%end if%>
                        <br> 
                        <%  ii = ii + 1
    if ii>50 then exit do
	rs.movenext     
	loop
	end if  
	rs.close   
	set rs=nothing
%>
                      </td>
                    </tr>
                  </table></td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
    </td>
    <td width="6" bgcolor="#FFFFFF">&nbsp;</td>
    <td background="IMAGES/menu-l.gif"> 
      <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td height="25" background="IMAGES/menu-l-m.gif"> 
            <%
  dim typename
  
set rs5=server.CreateObject("ADODB.RecordSet")
rs5.Source="select * from "& db_Type_Table &" where typeid=" & typeid &" order by typeorder"
rs5.Open rs5.Source,conn,1,1
typename=rs5("typename")
rs5.Close
set rs5=nothing

%>
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;  栏目导航 <a class="daohang" href="./" >网站首页</a>&gt;&gt;<a class="daohang" href="type.asp?typeid=<%=typeid%>"><%=typename%></a>&gt;&gt;<a class="daohang" href='BigClass.asp?typeid=<%=typeid%>&BigClassID=<%=request_BigClassID%>'><%=BigClassName%></a>&gt;&gt;<%=SmallClassName%></td>
        </tr>
        <tr> 
          <td background="IMAGES/menu-l.gif"><div align="center"> 
              <script language=javascript src=./zongg/ad.asp?i=13></script>
            </div></td>
        </tr>
        <tr> 
          <td><table border="0" cellpadding="3" cellspacing="0" width="100%" style="border-collapse: collapse">
              <%

set rs3=server.CreateObject("ADODB.RecordSet")
if uselevel=1 then
if Request.cookies(Forcast_SN)("key")="super" or Request.cookies(Forcast_SN)("key")="typemaster" or Request.cookies(Forcast_SN)("key")="bigmaster" or Request.cookies(Forcast_SN)("key")="smallmaster" or Request.cookies(Forcast_SN)("key")="check" then
rs3.Source="select * from "& db_News_Table &" where (SmallClassid=" & request_SmallClassid &" and checkked=1) order by newsid DESC"
end if
if Request.cookies(Forcast_SN)("key")="" then
rs3.Source="select * from "& db_News_Table &" where (SmallClassid=" & request_SmallClassid &" and checkked=1 ) order by newsid DESC"
end if
if Request.cookies(Forcast_SN)("key")="selfreg" then
	if Request.cookies(Forcast_SN)("reglevel")=3 then
rs3.Source="select * from "& db_News_Table &" where (SmallClassid=" & request_SmallClassid &" and checkked=1 ) order by newsid DESC"
	end if
	if Request.cookies(Forcast_SN)("reglevel")=2 then
rs3.Source="select * from "& db_News_Table &" where (SmallClassid=" & request_SmallClassid &" and checkked=1 ) order by newsid DESC"
	end if
	if Request.cookies(Forcast_SN)("reglevel")=1 then
rs3.Source="select * from "& db_News_Table &" where (SmallClassid=" & request_SmallClassid &" and checkked=1 ) order by newsid DESC"
	end if
end if
else
rs3.Source="select * from "& db_News_Table &" where (SmallClassid=" & request_SmallClassid &" and checkked=1) order by newsid DESC"
end if
rs3.Open rs3.Source,conn,3,1

If Not rs3.eof then
Rs3.PageSize     = MyPageSize
MaxPages         = Rs3.PageCount
Rs3.absolutepage = MyPage
total            = Rs3.RecordCount
%>
              <tr> 
                <td width=100% height="25" valign="middle" background="IMAGES/menu-l-zj.gif" > 
                  <p align="center">&nbsp;共 <%=total%> 条，当前第 <%=Mypage%>/<%=Maxpages%> 
                    页，每页 <%=MyPageSize%> 条</td>
              </tr>
              <tr> 
                <td width="100%" height="27" background="IMAGES/menu-l.gif"> 
                  <%
i = 0
do until rs3.Eof or i = rs3.PageSize

newsurl="ReadNews.asp?NewsID=" & rs3("NewsID")
newswwwurl=rs3("titleface")

if rs3("picname")<>"" then
img="<img src='images/news_img.gif' border='0'>"
else
img=""
end if
title=htmlencode4(trim(rs3("title")))
%>
                  <table width="98%" border="0" cellpadding="3" cellspacing="0">
                    <tr> 
                      <td width="55%"> 
                        <%if mode<>4 and  mode<>1 and  mode<>3 then%>
                        <%if mode="2" then%>&nbsp;&nbsp;<img src="IMAGES/006.gif" width="8" height="10"> <%=img%> <a class=middle href="<%if rs3("titleface")="无" then %><%=newsurl%><% else %> <%=newswwwurl%><%end if%>" target=_blank title="<%=title%>">
                        <%end if%>
                        <font color="<%=rs3("titlecolor")%>"> 
                        
                        <%=CutStr(title,40)%> 
                        
                         
                        </font></a>
						</td>
                      <td width="7%" align="right"> 
                        <%if showauthor="1" then%>
                        <%=rs3("Author")%> 
                        <%end if%>
                      </td>
                      <td width="28%" align="left"><font class=middle><%=trim(rs3("UpdateTime"))%></font></td>
                      <td width="5%" align="center"><font class=middle><font color="#66666"><%=trim(rs3("click"))%></font></font></td>
                      <td width="8%" align="left"><font class=middle>&nbsp;</font> 
                        <%if rs3("goodnews")="1" then%>
                        <img src="images/g.gif"> 
                        <%end if%>
                        <%else%>
                      </td>
                    </tr>
                    <tr> 
                      <td>&nbsp;&nbsp;<img src="IMAGES/004.gif" width="8" height="10">
	                       <%if mode="4" or  mode="1" or mode="3" then%>
                        <%=img%> 
                        <%end if%>				  
					  <a class=middle href="ReadNews.asp?NewsID=<%=rs3("NewsID")%>" target=_blank title="<%=title%>"> 
                        <font color="<%=rs3("titlecolor")%>"> 
                        
                        <%=CutStr(title,40)%> 
                        
                        
                        
                        </font></a></td>
                      <td align="right"> 
                        <%if showauthor="1" then%>
                        <%=rs3("Author")%> 
                        <%end if%>
                      </td>
                      <td align="left">&nbsp;</td>
                      <td align="center"><font class=middle><%=trim(rs3("UpdateTime"))%></font></td>
                      <td align="left"><font class=middle><font color="#666666"><%=trim(rs3("click"))%></font></font> 
                        <%if rs3("goodnews")="1" then%>
                        <img src="images/g.gif" > 
                        <%end if%>
                        <%end if%>
                      </td>
                    </tr>
                  </table>
                  <%
rs3.MoveNext
i = i + 1
loop
%>
                </td>
              </tr>
              <tr> 
                <td width="100%" height="20" align=center valign="middle" background="IMAGES/menu-l.gif">共 
                  <%=total%> 条，当前第 <%=Mypage%>/<%=Maxpages%> 页，每页 <%=MyPageSize%> 
                  条 
                  <%
url="SmallClass.asp?typeid=" &typeid & "&BigClassid=" & request_BigClassid & "&SmallClassid=" & request_SmallClassid									
PageNextSize=int((MyPage-1)/PageShowSize)+1
Pagetpage=int((total-1)/rs3.PageSize)+1

if PageNextSize >1 then
PagePrev=PageShowSize*(PageNextSize-1)
Response.write "<a class=black href='" & Url & "&page=" & PagePrev & "' title='上" & PageShowSize & "页'>上一翻页</a> "
Response.write "<a class=black href='" & Url & "&page=1' title='第1页'>页首</a> "
end if
if MyPage-1 > 0 then
Prev_Page = MyPage - 1
Response.write "<a class=black href='" & Url & "&page=" & Prev_Page & "' title='第" & Prev_Page & "页'>上一页</a> "
end if

if Maxpages>=PageNextSize*PageShowSize then
PageSizeShow = PageShowSize
Else
PageSizeShow = Maxpages-PageShowSize*(PageNextSize-1)
End if
If PageSizeShow < 1 Then PageSizeShow = 1
for PageCounterSize=1 to PageSizeShow
PageLink = (PageCounterSize+PageNextSize*PageShowSize)-PageShowSize
if PageLink <> MyPage Then
Response.write "<a class=black href='" & Url & "&page=" & PageLink & "'>[" & PageLink & "]</a> "
else
Response.Write "<B>["& PageLink &"]</B> "
end if
If PageLink = MaxPages Then Exit for
Next

if Mypage+1 <=Pagetpage  then
Next_Page = MyPage + 1
Response.write "<a class=black href='" & Url & "&page=" & Next_Page & "' title='第" & Next_Page & "页'>下一页</A>"
end if

if MaxPages > PageShowSize*PageNextSize then
PageNext = PageShowSize * PageNextSize + 1
Response.write " <A class=black href='" & Url & "&page=" & Pagetpage & "' title='第"& Pagetpage &"页'>页尾</A>"
Response.write " <a class=black href='" & Url & "&page=" & PageNext & "' title='下" & PageShowSize & "页'>下一翻页</a>"
End if
else
Response.write "<tr><td align=center background=""IMAGES/menu-l.gif"">&nbsp;本类暂无信息</td></tr>"
				
End If
Rs3.close
set rs3=nothing

%>
                </td>
              </tr>
            </table></td>
        </tr>
        <tr> 
          <td background="IMAGES/menu-l.gif">&nbsp;</td>
        </tr>
      </table>
      <p>&nbsp;</p></td>
  </tr>
  <tr valign="top">
    <td height="20" background="IMAGES/menu-d-t.gif">&nbsp;</td>
    <td height="20" bgcolor="#FFFFFF">&nbsp;</td>
    <td height="20" background="IMAGES/menu-l-t.gif" class="daohang">&nbsp;</td>
  </tr>
</table>
<!--#include file=bottom.asp --></td>
<td width="4" height="209">　</td>
    
<td width="636" height="209" valign=top>&nbsp; </td>
  </tr>
</table>
</body>
</html><%end if%>