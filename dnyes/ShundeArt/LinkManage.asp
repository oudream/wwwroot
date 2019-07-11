<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="ChkUser.asp" -->
<!--#include file="ChkManage.asp" -->
<%
if not(request.cookies(Forcast_SN)("ManageKEY")="super" or request.cookies(Forcast_SN)("ManageKEY")="bigmaster" or request.cookies(Forcast_SN)("ManageKEY")="check" or request.cookies(Forcast_SN)("ManageKEY")="typemaster") then
	Show_Err("对不起，您的管理权限不够！")
	response.end
else
	if linkmana="1" or (request.cookies(Forcast_SN)("ManagePurview")="99999" and request.cookies(Forcast_SN)("purview")="99999") then
		PageShowSize = 10            '每页显示多少个页
		MyPageSize   = 10           '每页显示多少条文章
			
		If Not IsNumeric(Request("page")) Or IsEmpty(Request("page")) Or Request("page") <=0 Then
			MyPage=1
		Else
			MyPage=Int(Abs(Request("page")))
		End if
		
		set rs=server.CreateObject("ADODB.RecordSet")
		if linkpass="1" then
			rs.Source="select * from "& db_link_Table &" order by ID desc"
		else
			rs.Source="select * from "& db_link_Table &" where pass=0 order by ID desc"
		end if
		rs.Open rs.Source,conn,1,1
		
		If Not rs.eof then
			rs.PageSize     = MyPageSize
			MaxPages         = rs.PageCount
			rs.absolutepage = MyPage
			total            = rs.RecordCount
			%>
<html><head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=site.css rel=stylesheet>
<title><%=copyright%><%=version%>&nbsp;<%=ver%> - 友情链接管理</title>
</head>
<body topmargin="0">
<!--#include file=Admin_Top.asp-->
<table border="1" cellpadding="3" cellspacing="0" style="border-collapse: collapse" bordercolor="#C0C0C0" width="100%" id="AutoNumber1">
<tr class="TDtop">
	<td colspan=8 width="100%" height="25" align="center">┊ <strong>友情链接</strong> ┊</td>
</tr>
<tr>
	<td colspan=8 align=center height=25>
		共 <%=total%> 条，当前第 <%=Mypage%>/<%=Maxpages%> 页，每页 <%=MyPageSize%> 条 
		<a href="linkreg.asp">添加链接</a>
	</td>
</tr>
<tr class="TDtop1" align="center" >
	<td width="6%">ID号</td>
	<td width="6%">类型</td>
	<td width="20%">网站名称</td>
	<td width="20%">网站LOGO</td>
	<td width="14%">站长</td>
	<td width="12%">执行</td>
	<td width="6%">编辑</td>
	<td width="6%">删除</td>
</tr>
			<%
			for i=1 to rs.PageSize
				if not rs.EOF then
					%>
<form method="POST" action="<%if rs("pass")=0 then%>LinkPass.asp<%else%>LinkUnPass.asp<%end if%>?ID=<%=rs("ID")%>">
<tr bgcolor=ffffff>
	<td align=center><%=rs("ID")%></td>
	<td align=center  > 
		<%if rs("linktype")=2 then%>首页LOGO链接<%else%><%if rs("linktype")=1 then%>其他LOGO链接<%else%>文字<%end if%><%end if%>
	</td>
	<td align=center >
		<a href="<%=rs("weburl")%>" title="<%=rs("content")%>" target="_blank"><%=rs("webname")%></a>
	</td>
	<td height="35" align=center valign="middle" > 
		<%if rs("linktype")<>0 then%><a target="_blank" title="点击查看原图" href="<%=rs("logo")%>"><img src="<%=rs("logo")%>" border=0 width=88 height=31></a><%else%>文字链接<%end if%>
	</td>
	<td align=center  >
		<a href="mailto:<%=rs("email")%>"><%=rs("webmaster")%></a>
	</td>
	<td align=center  >
		<input type=hidden name=ID value="<%=rs("ID")%>">
		<input type=submit value="<%if rs("pass")="0" then%>通过<%else%>取消<%end if%>审核" class=button onMouseOver="window.status='按这个按钮<%if rs("pass")="0" then%>通过<%else%>取消<%end if%>审核网站';return true;" onMouseOut="window.status='';return true;"target=_blank title="按这个按钮<%if rs("pass")="0" then%>通过<%else%>取消<%end if%>审核网站" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
	</td>
	<td align="center"  >
		<a href='LinkEdit.asp?id=<%=rs("ID")%>'>编辑</a>
	</td>
	<td align=center  >
		<a href='LinkDel.asp?id=<%=rs("ID")%>'>删除</a>
	</td>
</tr>
</form>
					<%
					rs.MoveNext
				end if
			next
			%>
<tr>
	<td colspan=8  align=center height=25>
		共 <%=total%> 条，当前第 <%=Mypage%>/<%=Maxpages%> 页，每页 <%=MyPageSize%> 条 
			<%
			if request.cookies(Forcast_SN)("KEY")="super" then			
				url="linkmanage.asp?" 
			else
				url="linkmanage.asp?"
			end if									
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
</tr>
</table>
</body>
</html>
<!--#include file=Admin_Bottom.asp-->
			<%
		else
			Show_Err("没有要申请链接的网站<br><br><a href='linkreg.asp'>添加链接</a>")
			response.end
		End If
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
	else
		Show_Err("对不起，该功能已经被超级系统管理员关闭，您没有权限操作！")
		response.end
	end if
end if
%>