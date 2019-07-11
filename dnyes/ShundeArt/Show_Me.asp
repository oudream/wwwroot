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
<title>顺德国际艺术交流中心_作品</title>
<LINK href=news.css rel=stylesheet>
</head>

<body topmargin="0">
<!--#include file=top.asp -->
<%
sql="select * from art_subs where subs_id=" & request("subs_id") 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1
%>
<table width="760" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
	<tr valign="top"> 
		<td>
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr> 
					<td valign="top">
						<table width="102%" border="0" cellspacing="0" cellpadding="0">
							<tr bgcolor="#FFFFFF"> 
								<td height="3" bgcolor="#FFFFFF"><img src="images/kb.gif" width="9" height="3"></td>
							</tr>
							<tr bgcolor="#FFFFFF"> 
								
                <td width="100%" height="25" background="images/menu-guestbook.gif" >&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font class=m_tittle>&nbsp;</font>栏目导航&nbsp;<b><a class="daohang" href="./" > 
                  名家作品</a>&gt;&gt;欣赏</b></td>
							</tr>
							<tr bgcolor="#FFFFFF"> 
	                					
                <td height="25" background="images/menu-guest-l.gif" bgcolor="#FFFFFF" align="right">&nbsp;</td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr valign="top"> 
		<td background="images/menu-guest-l.gif"> 
			
			
			
			
			
			
			
			
			
			
			
			
			
			<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
				<tr> 
          				<td>
          					<table border="0" cellpadding="0" cellspacing="0" width="98%" align="center">
              <tr> 
                <td width="100%" align=center> <br>
                  <font size="+2"><%=rs("sname")%></font>
                  <HR> </td>
              </tr>
              <tr> 
                <td width="100%"  align="center"><img src="<%=rs("simg")%>"> 
                  <br> <br> 
				  
				  
			
<H2>&nbsp;</H2>
<!----------------------------------------------------------------------------------------------------------------------
<!--                                                   ERROR MSGS
<!-------------------------------------------------------------------------------------------------------------------- -->
  
				  
				  <table width="90%" border="0" cellpadding="1" cellspacing="2" bgcolor="f5f5f5">
                    <tr> 
                      <td width="68" bgcolor="#FFFFFF">作品名称:</td>
                      <td bgcolor="#FFFFFF"><font 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 ><%=rs("sname")%></font></td>
                      <td width="50" bgcolor="#FFFFFF">尺寸:</td>
                      <td bgcolor="#FFFFFF"><font 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 ><%=rs("sproperties")%></font></td>
                    </tr>
                    <tr> 
                      <td bgcolor="#FFFFFF">作者:</td>
                      <td bgcolor="#FFFFFF"><font 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 > 
                        <%
psql="select * from art_artist where artist_id=" & rs("artist_id")
Set prs=Server.CreateObject("ADODB.RecordSet")
prs.Open psql,conn,1,1
if not ( prs.eof or prs.bof ) then
%>
                        <%=prs("allname")%> 
                        <%
end if
prs.close
%>
                        </font></td>
                      <td bgcolor="#FFFFFF">价格</td>
                      <td bgcolor="#FFFFFF"><font 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 ><%=rs("price")%></font></td>
                    </tr>
                    <tr> 
                      <td bgcolor="#FFFFFF">简介:</td>
                      <td colspan="3" bgcolor="#FFFFFF"><font face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 ><%=rs("memo")%></font></td>
                    </tr>
                    <tr align="center" bgcolor="#FFFFFF"> 
                      <td colspan="4"><A 
      href="javascript:print()">〖 打印此页 〗</a>　　<A 
      href="javascript:close()">〖 关闭窗口 〗</a></td>
                    </tr>
                  </table></td>
              </tr>
              <tr> 
                <td width="100%"  height="25"> <div align="center"> </div></td>
              </tr>
              <tr> 
                <td width=100% ><hr size=1></td>
              </tr>
              <tr> 
                <td height=8></td>
              </tr>
            </table>
					</td>
				</tr>
			</table>
		</td>
	</tr>

    
	<tr valign="top">
		<td height="25" background="images/menu-l-guest.gif"></td>
	</tr>
</table>

<%
set prs=nothing
rs.close
set rs=nothing
%>
<!--#include file=bottom.asp -->
</body>
</html>
