<%@ Language=VBScript%>
<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<%
request_BoardID=Request.QueryString("ID")
if request_BoardID="" then
Response.Write "<script>alert('未指定参数');history.back()</script>"
response.end
else

if not IsNumeric(request_BoardID) then
 response.write "<script>alert('非法参数');history.back()</script>"
 response.end
else

set rs=server.CreateObject("ADODB.RecordSet")
rs.Source="select * from "& db_Board_Table &" where id="&request_BoardID&" order by ID"
rs.Open rs.Source,conn,1,1
if rs.EOF then
Response.Write "<script>alert('该公告不存在');history.back()</script>"
else
title=rs("title")
rs.close
set rs=nothing
end if

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
<title>公告详细内容_<%=title%></title>
<LINK href=news.css rel=stylesheet>
</head>

<body topmargin="0">
<!--#include file="top.asp"-->
<table width="760" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <tr valign="top"> 
    <td width="180" height="234" background="IMAGES/menu-d.gif"> 
      <table width="100%" border="0" cellpadding="0" cellspacing="0" id="AutoNumber2" style="border-collapse: collapse">
        
        <tr> 
          <td height=3 valign="middle" bgcolor="#FFFFFF"><img src="IMAGES/kb.gif" width="9" height="3"></td>
        </tr>
        <tr> 
          <td width="175" height=25 valign="middle" background="IMAGES/menu-d-gg.gif"> 
            <p align="center">全部公告<strong><br>
              </strong></td>
        </tr>
		<%
set rs=server.CreateObject("ADODB.RecordSet")
rs.Source="select * from "& db_Board_Table &" order by ID DESC "
rs.Open rs.Source,conn,1,1
If Not rs.eof then
while not rs.EOF
%>
        <tr> 
          <td width="175" height="2" align=center background="IMAGES/menu-d.gif" style="WORD-WRAP: break-word"><table width="90%" border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td><div align="left"><br><li><A class=middle HREF="Board_news.asp?ID=<%=rs("ID")%>"><%=rs("title")%></a><br></div></td>
              </tr>
            </table></td>
        </tr>
<%
rs.MoveNext
wend

else
Response.write "<tr><td valign=top height=80>&nbsp;暂无信息</td></tr>"				
End If
rs.close
set rs=nothing
%>
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
            当前位置：<a class="daohang" href="./" >网站首页</a>&gt;&gt;<a class="daohang" href="./board.asp" >全部公告</a>&gt;&gt;公告详细内容</b></td>
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
rs2.Source="select * from "& db_Board_Table &" where id="&request_BoardID&" order by ID"
rs2.Open rs2.Source,conn,1,1
content=trim(rs2("content"))
%>


                          <table border="0" cellpadding="0" cellspacing="0" width="99%" style="border-collapse: collapse">
						  <tr>
							<td width="50%" height="20" ><b>公告标题：</b><%=rs2("title")%></td>
                            </tr>
							
						  <tr>
							<td width="50%" height="20" >发布人：<%=rs2("upload")%></td>
                            <td width="50%" height="20" align=right>发布时间：<%=rs2("dateandtime")%><br></td>
                            </tr>
							<tr> 
                              <td width="100%" height="20" colspan="2" >
<hr width="100%" size="1">
<%	
DisplayContent=Content
Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;" & displaycontent
%>
							  
							  </td>
                            </tr>
							
							</table>
             
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
<% end if
end if 
%>
