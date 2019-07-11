<%@ Language=VBScript%>
<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<!--#include file="function.asp"-->
<%
PageShowSize = 10            '每页显示多少个页
MyPageSize   = 5           '每页显示多少条新闻
	
If Not IsNumeric(Request("page")) Or IsEmpty(Request("page")) Or Request("page") <=0 Then
MyPage=1
Else
MyPage=Int(Abs(Request("page")))
End if

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
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>全部公告__<%=jjgn%></title>
<LINK href=news.css rel=stylesheet>
</head>

<body topmargin="0">
<!--#include file="top.asp"-->
<table width="760" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <tr valign="top"> 
    <td width="180" height="234" background="IMAGES/menu-d.gif"> 
      <table width="100%" border="0" cellpadding="0" cellspacing="0" id="AutoNumber2" style="border-collapse: collapse">
        <%
set rs=server.CreateObject("ADODB.RecordSet") 
rs.Source="select * from "& db_Board_Table &" where inuse=1"
rs.Open rs.Source,conn,1,1
if not rs.EOF then
%>
        <tr> 
          <td height=3 valign="middle" bgcolor="#FFFFFF"><img src="IMAGES/kb.gif" width="9" height="3"></td>
        </tr>
        <tr> 
          <td width="175" height=25 valign="middle" background="IMAGES/menu-d-gg.gif"> 
            <p align="center">当前公告<strong><br>
              </strong></td>
        </tr>
        <tr> 
          <td width="175" height="2" align=center background="IMAGES/menu-d.gif" style="WORD-WRAP: break-word"><table width="90%" border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td><div align="center"><br>
                    <b><%=rs("title")%></b><br><br>
                    <%=rs("content")%> <br>
                    发布人：<%=rs("upload")%><br>
                    发布时间：<%=rs("dateandtime")%> </div></td>
              </tr>
            </table></td>
        </tr>
        <%else
     rs.close
      set rs=nothing
      end if%>
      </table>
    </td>
    <td width="6">&nbsp;</td>
    <td width="574" background="IMAGES/menu-l.gif">
      <table width="100%" height="234" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td height="3" bgcolor="#FFFFFF"><img src="IMAGES/kb.gif" width="9" height="3"></td>
        </tr>
        <tr> 
          <td height="25" background="IMAGES/menu-l-m.gif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font class=m_tittle>&nbsp;</font>栏目导航<b>&nbsp;&nbsp; 
            当前位置：<a class="daohang" href="./" >网站首页</a>&gt;&gt;全部公告<font color=red></font></b></td>
        </tr>
        <tr> 
          <td height="25" background="IMAGES/menu-l.gif"><div align="center"> 
              <font color=red><a class=daohang href="./" > 
              <script language=javascript src=./zongg/ad.asp?i=13></script> 
              </a> </font></div></td>
        </tr>
        <tr> 
          <td height="25" background="IMAGES/menu-l-zj.gif" class="daohang">&nbsp;</td>
        </tr>
        <tr> 
          <td height="156" background="IMAGES/menu-l.gif"> 
            <div align="center"> 
              <table width="95%" height="154" border="0" align="center" cellpadding="0" cellspacing="0" id="AutoNumber3" style="border-collapse: collapse">
                <tr> 
                  <td height="10" valign=top>&nbsp;</td>
                </tr>
                <tr> 
                  <td height="52" valign=top><table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%" id="AutoNumber3" height="45">
                      <tr> 
                        <td width="100%" height="45"> 
                          <%
set rs2=server.CreateObject("ADODB.RecordSet")
rs2.Source="select * from "& db_Board_Table &" order by ID DESC "
rs2.Open rs2.Source,conn,3,1

If Not rs2.eof then
rs2.PageSize     = MyPageSize
MaxPages         = rs2.PageCount
rs2.absolutepage = MyPage
total            = rs2.RecordCount
i = 0
do until rs2.Eof or i = rs2.PageSize
%>
                          <table border="0" cellpadding="0" cellspacing="0" width="99%" style="border-collapse: collapse">
                            <tr> 
                              <td width="100%" height="20" colspan="2" ><br><b><li><%=rs2("title")%></b>
							  <br><hr width="100%" size="1">
							  </td>
                            </tr>
							
                            <tr> 
                              <td width="100%" height="20" colspan="2"><%=CutStr(rs2("content"),300)%>…</td>
                            </tr>
                            <tr> 
                              <td width="50%" height="20" ><br>发布人：<%=rs2("upload")%></td>
                              <td width="50%" height="20" align=right>发布时间：<%=rs2("dateandtime")%><br></td>
                            </tr>
                          </table>
                          <%
rs2.MoveNext
i = i + 1
loop
%>
                        </td>
                      </tr>
                      <tr> 
                        <td width="100%" align=center colspan="2"><hr width="100%" size="1"><br>共 <%=total%> 
                          条，当前第 <%=Mypage%>/<%=Maxpages%> 页，每页 <%=MyPageSize%> 
                          条 
                          <%
url="board.asp?"				
PageNextSize=int((MyPage-1)/PageShowSize)+1
Pagetpage=int((total-1)/rs2.PageSize)+1

if PageNextSize >1 then
PagePrev=PageShowSize*(PageNextSize-1)
Response.write "<a class=black href='" & Url & "page=" & PagePrev & "' title='上" & PageShowSize & "页'>上一翻页</a> "
Response.write "<a class=black href='" & Url & "page=1' title='第1页'>页首</a> "
end if
if MyPage-1 > 0 then
Prev_Page = MyPage - 1
Response.write "<a class=black href='" & Url & "page=" & Prev_Page & "' title='第" & Prev_Page & "页'>上一页</a> "
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
Response.write "<a class=black href='" & Url & "page=" & PageLink & "'>[" & PageLink & "]</a> "
else
Response.Write "<B>["& PageLink &"]</B> "
end if
If PageLink = MaxPages Then Exit for
Next

if Mypage+1 <=Pagetpage  then
Next_Page = MyPage + 1
Response.write "<a class=black href='" & Url & "page=" & Next_Page & "' title='第" & Next_Page & "页'>下一页</A>"
end if

if MaxPages > PageShowSize*PageNextSize then
PageNext = PageShowSize * PageNextSize + 1
Response.write " <A class=black href='" & Url & "page=" & Pagetpage & "' title='第"& Pagetpage &"页'>页尾</A>"
Response.write " <a class=black href='" & Url & "page=" & PageNext & "' title='下" & PageShowSize & "页'>下一翻页</a>"
End if

else
Response.write "<tr><td valign=top height=80>&nbsp;暂无信息</td></tr>"				
End If
rs2.close

%>
                        </td>
                      </tr>
                    </table></td>
                </tr>
                <tr> 
                  <td height="16" valign=top>&nbsp;</td>
                </tr>
              </table>
            </div></td>
        </tr>
      </table>    </td>
  </tr>
  <tr valign="top"> 
    <td height="20" background="IMAGES/menu-d-t.gif">&nbsp;</td>
    <td bgcolor="#FFFFFF">&nbsp;</td>
    <td height="20" background="IMAGES/menu-l-t.gif" class="daohang">&nbsp;</td>
  </tr>
</table>
<!--#include file="bottom.asp"-->
</body>
</html>