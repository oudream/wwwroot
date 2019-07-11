<!--#include file="Conn.ASP"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp" -->
<!--#include file="ChkURL.asp"-->
<%dim jingyong
set rs=server.createobject("adodb.recordset")
sql="select * from "& db_User_Table &" where  "& db_User_Name &"='"&Request.cookies(Forcast_SN)("username")&"'"
rs.open sql,ConnUser,1,3
if rs.bof or rs.eof then
	response.redirect "login.asp"
	response.end
end if
jingyong=rs("jingyong")
rs.close
set rs=nothing

if Request.cookies(Forcast_SN)("key")="super" or Request.cookies(Forcast_SN)("key")="check" or Request.cookies(Forcast_SN)("KEY")="bigmaster" or Request.cookies(Forcast_SN)("KEY")="smallmaster" or Request.cookies(Forcast_SN)("KEY")="typemaster" or (Request.cookies(Forcast_SN)("key")="selfreg" and jingyong=0) then
	if request.QueryString("action")="del" then
		conn.execute "delete from news where newsid in ("&request.form("selectdel")&")"
		response.Redirect  "mynews.asp"
	end if 
	PageShowSize = 10            '每页显示多少个页
	MyPageSize   = 16           '每页显示多少条文章
		
	If Not IsNumeric(Request("page")) Or IsEmpty(Request("page")) Or Request("page") <=0 Then
		MyPage=1
	Else
		MyPage=Int(Abs(Request("page")))
	End if
	%>
<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="site.css" rel=stylesheet>
<title><%=copyright%><%=version%><%=ver%></title>
<script language="JavaScript">
<!--
function CheckAll(form)  {
  for (var i=0;i<form.elements.length;i++)    {
    var e = form.elements[i];
    if (e.name != 'chkall')       e.checked = form.chkall.checked; 
   }
  }
//-->
</script>
</head>
<body topmargin="0">
<!--#include file="Admin_Top.asp"-->
<div align="center">
<table border="1" cellpadding="3" cellspacing="0" style="border-collapse: collapse" bordercolor="#C0C0C0" width="100%" id="AutoNumber1">
<tr> 
<td colspan=8 class="TDtop" height=25> 
	<div align="center" >┊ <B>我的文章</B> ┊</div>
</td>
</tr>
<tr>
<td colspan=8 align=center height=25><a href="newsaddd1.asp">添加文章</a></td>
</tr>
<%set rs=server.CreateObject("ADODB.RecordSet")
rs.Source="select * from "& db_News_Table &" where editor='" & CheckStr(request.cookies(Forcast_SN)("UserName")) & "' order by NewsID desc"
rs.open rs.Source,conn,1,3
if not rs.eof then
	rs.PageSize     = MyPageSize
	MaxPages         = rs.PageCount
	rs.absolutepage = MyPage
	total            = rs.RecordCount
	i = 1
	%>
	<tr align="center" class="TDtop1">
	<td width="7%">ID号</td>
	<td width="39%">标 &nbsp;&nbsp;&nbsp;题</td>
	<td width="16%">日&nbsp;&nbsp;&nbsp;期</td>
	<td width="10%">状态</td>
	<td width="8%">编辑</td>
	<td width="4%">选择</td>
	</tr>
	<form action="MyNews.asp?action=del" method=post name=check>
	<%do while not rs.eof and (i <= MyPageSize)%>
	<tr bgcolor=ffffff height="20">
	<td width=7% align=center ><%=rs("NewsID")%></td>
	<td width=39% ><a href='ReadNews.asp?NewsID=<%=rs("NewsID")%>' target=_blank title="查看文章“<%=rs("Title")%>”"><font color="<%=rs("titlecolor")%>"><%=left(rs("Title"),40)%></font></a><%if rs("picname")<>"" then%><font color=red>[图]<%end if%></td>
	<td width=16% align=center ><%=rs("UpdateTime")%>　</td>
	<td width=10% align=center ><%if rs("checkked")<>1 then%>未审<%else%>已审<%end if%>；<%if rs("istop")<>1 then%>未固<%else%>已固<%end if%></td>
	<td width=8% align=center ><a href="NewsEdit.asp?NewsID=<%=rs("NewsID")%>">编辑</a></td>
	<td width="4%" align=center><input name="selectdel" type="checkbox" id="selectdel" value=<%=rs("NewsID")%>>	</td>
	</tr>
	<%rs.MoveNext
	i = i + 1
	loop%>
	<tr>
	<td colspan="8" align=right height="25" width="100%">
	<input type=checkbox name=chkall value=on onclick="CheckAll(this.form)">选中所有显示文章&nbsp;
	<input type=submit name=action onclick="{if(confirm('删除选定的文章吗？')){this.document.check.submit();return true;}return false;}" value="删除文章" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
	</td>
	</tr>
	</form>
<tr>
    <td colspan=8  align=center height=25>
共 <%=total%> 条，当前第 <%=Mypage%>/<%=Maxpages%> 页，每页 <%=MyPageSize%> 条 
<%
url="mynews.asp?"
							
PageNextSize=int((MyPage-1)/PageShowSize)+1
Pagetpage=int((total-1)/rs.PageSize)+1

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
%>
</td>

<%else
	Response.write "<td width='100%' align=center><B>你还没有发表过文章!请点击添加文章。</B></td>"
end if
rs.close
set rs=nothing%>
</tr>
</table>
</div>
</body>
</html>
<!--#include file=Admin_Bottom.asp-->
<%else%>
	<script language=javascript>
		history.back()
		alert("对不起，管理员尚未开通您的帐号，请稍候！")
	</script>
<%end if%>