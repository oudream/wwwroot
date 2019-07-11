<%@ Language=VBScript%>
<!--#include file=conn.asp -->
<!--#include file=config.asp -->
<%
ReViewID=Request.QueryString("ReViewID")
newsID=Request.QueryString("newsID")

if ReViewID="" then
	Response.Write "<script>alert('未指定参数');history.back()</script>"
	response.end
else
	if not IsNumeric(ReViewID) then
		response.write "<script>alert('非法参数');history.back()</script>"
		response.end
	else
		set rs=server.CreateObject("ADODB.RecordSet")
		rs.Source="select * from review where ReViewID=" & ReViewID & " order by reviewid desc"
		rs.Open rs.Source,conn,1,3
		if rs.EOF then  NoReview=1
			rs.close
			set rs=nothing
		end if
	end if
	%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>文章评论_<%=jjgn%></title>
<LINK href=news.css rel=stylesheet>
</head>
<body topmargin="0">
<!--#include file="top.asp"-->
<form method="POST" action="AddReview.asp">
  <table width="760" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
    <tr> 
      <td height="3" ><img src="IMAGES/kb.gif" width="9" height="3"></td>
    </tr>
    <tr> 
      <td height="25" background="IMAGES/menu-guestbook.gif" bgcolor="CDCCCC" >&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;栏目导航<a class=daohang href="./" >&nbsp;</a><strong><b>&nbsp;&nbsp;当前位置：<a class="daohang" href="./" >网站首页</a>&gt;&gt;评论</b></strong></td>
    </tr>
    <tr bgcolor="CDCCCC"> 
      <td background="IMAGES/menu-guest-l.gif" bgcolor="#FFFFFF"><div align="center"><a class=daohang href="./" > 
          </a>
          <script language=javascript src=./zongg/ad.asp?i=13></script></script>
          </div></td>
    </tr>
    <tr valign="top"> 
      <td height="25" background="IMAGES/menu-l-guest.gif">&nbsp;</td>
    </tr>
    <tr valign="top"> 
      <td background="IMAGES/menu-guest-l.gif"> <table width="95%"  border="0" align="center" cellpadding="0" cellspacing="0" id="AutoNumber3" style="border-collapse: collapse">
          <tr> 
            <td valign=top>&nbsp;</td>
          </tr>
          <tr> 
            <td valign=top><table border="0" cellpadding="0" cellspacing="0" width="100%">
                <tr> 
                  <td width="100%" height="20" colspan="2"></td>
                </tr>
                <tr> 
                  <td width="100%" colspan="2"> <%if NoReview then Response.Write "该信息当前没有任何评论！"%> </td>
                </tr>
                <%
if not NoReview then
set rs=server.CreateObject("ADODB.RecordSet")
rs.Source="select * from review where ReViewID=" & ReViewID & " order by reviewid desc"
rs.Open rs.Source,conn,1,1

author=trim(rs("author"))
reviewip=rs("reviewip")
email=server.HTMLEncode(trim(rs("email")))
content=trim(rs("content"))
%>
                <tr> 
                  <td colspan="2"> <table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="efefef" style="table-layout:fixed; word-break:break-all">
                      <tr> 
                        <td width="322">发表人：<%=author%></td>
                        <td width="270"> <p align="right"> 
                            <%if Request.cookies(Forcast_SN)("key")="super" or showip="1" then%>
                            IP：<%=reviewip%> 
                            <%end if%>
                        </td>
                      </tr>
                      <tr> 
                        <td colspan="2"> <table width="100%" border="0" cellspacing="0" cellpadding="0" style="table-layout:fixed; word-break:break-all">
                            <tr> 
                              <td>发表人邮件：<a href='mailto:<%=email%>'><%=email%></a></td>
                              <td align="right">发表时间：<%=rs("updatetime")%></td>
                            </tr>
                          </table></td>
                      </tr>
                      <tr> 
                        <td colspan="2" bgcolor="#FFFFFF"><%=content%></td>
                      </tr>
                    </table></td>
                </tr>
                <tr> 
                  <td colspan="2">&nbsp; </td>
                </tr>
                <%end if
				rs.close
				set rs=nothing%>
                <tr> 
                  <td width="100%" colspan="2">&nbsp; </td>
                </tr>
              </table></td>
          </tr>
          <tr> 
            <td valign=top>&nbsp;</td>
          </tr>
        </table></td>
    </tr>
    <tr valign="top">
      <td height="19" background="IMAGES/menu-guest-t.gif">&nbsp;</td>
    </tr>
  </table>
<!--#include file="bottom.asp"-->
</form>
</body>
</html>