<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp" -->
<!--#include file="ChkManage.asp" -->
<!--#include file="ChkURL.asp"-->
<%
if not(request.cookies(Forcast_SN)("ManageKEY")="super" or request.cookies(Forcast_SN)("ManageKEY")="typemaster" or request.cookies(Forcast_SN)("ManageKEY")="bigmaster") then
	Show_Err("对不起，您的后台管理权限不够操作！")
	response.end
else
	bigclassid=request("bigclassid")
	if request.cookies(Forcast_SN)("ManageKEY")="bigmaster" then
		if bigclassid="" then
			Show_Err("请选择文章大类<br><br>－－－<a href='javascript:history.back()'>回去重来</a>－－－")
			Response.End
		end if
	end if

	PageShowSize = 10            '每页显示多少个页
	MyPageSize   = 16           '每页显示多少条文章
		
	If Not IsNumeric(Request("page")) Or IsEmpty(Request("page")) Or Request("page") <=0 Then
		MyPage=1
	Else
		MyPage=Int(Abs(Request("page")))
	End if
	
	set rs4=server.createobject("adodb.recordset")
	sql4="select * from "& db_Type_Table &" where typemaster='"&session("username")&"'"
	rs4.open sql4,conn,3,3
	
	set rs=server.CreateObject("ADODB.RecordSet")
	if request.cookies(Forcast_SN)("ManageKEY")="bigmaster" then
		if checkshow=1 then
			rs.Source="select * from "& db_News_Table &" where bigclassid="&bigclassid&" order by NewsID desc"
		else
			rs.Source="select * from "& db_News_Table &" where checkked=0 and bigclassid="&bigclassid&" order by NewsID desc"
		end if
	end if
	
	if request.cookies(Forcast_SN)("ManageKEY")="super" or request.cookies(Forcast_SN)("ManageKEY")="check" then
		if checkshow=1 then
			rs.Source="select * from "& db_News_Table &" order by NewsID desc"
		else
			rs.Source="select * from "& db_News_Table &" where checkked=0 order by NewsID desc"
		end if
	end if
	
	if request.cookies(Forcast_SN)("ManageKEY")="typemaster" then
		if checkshow=1 then
			rs.Source="select * from "& db_News_Table &" where typeid="&rs4("typeid")&" order by NewsID desc"
		else
			rs.Source="select * from "& db_News_Table &" where checkked=0 and typeid="&rs4("typeid")&" order by NewsID desc"
		end if
	end if
	
	rs.Open rs.Source,conn,3,1
	If Not rs.eof then
		rs.PageSize	= MyPageSize
		MaxPages	= rs.PageCount
		rs.absolutepage	= MyPage
		total		= rs.RecordCount
	%>

<html><head>
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=site.css rel=stylesheet>
</head>
<title><%=copyright%><%=version%>&nbsp;<%=ver%> - 审核文章</title>
<body topmargin="0">
<!--#include file=Admin_Top.asp-->
<div align="center">
<table border="1" cellpadding="3" cellspacing="0" style="border-collapse: collapse" bordercolor="#C0C0C0" width="100%" id="AutoNumber1">
<tr  class="TDtop">
	<td colspan=8 width="100%" height="25" align="center">┊ <strong>审核文章</strong> ┊</td>
</tr>
<tr>
	<td colspan=8  align=center height=25>共 <%=total%> 条，当前第 <%=Mypage%>/<%=Maxpages%> 页，每页 <%=MyPageSize%> 条</td>
</tr>
<tr align="center" class="TDtop1">
	<td width="7%">ID号</td>
	<td width="51%">标 &nbsp;&nbsp;&nbsp;题</td>
	<td width="16%">日&nbsp;&nbsp;&nbsp;期</td>
	<td width="3%">图</td>
	<td width="3%">核</td>		
	<td width="6%">执行</td>
	<td width="6%">编辑</td>
	<td width="6%">删除</td>		
</tr>
		<%
		for i=1 to rs.PageSize
			if not rs.EOF then
			%>
<form method="POST" target=_self  action="<%if rs("checkked")="0" then%>checknews4.asp<%else%>checknews3.asp<%end if%>?newsid=<%=rs("NewsID")%>">
<tr bgcolor=ffffff>
	<td width=7% align=center ><%=rs("NewsID")%>　</td>
	<td>&nbsp;<a  href='ReadNews.asp?NewsID=<%=rs("NewsID")%>' onMouseOver="window.status='查看文章“<%=trim(rs("Title"))%>”';return true;" onMouseOut="window.status='';return true;"target=_blank title="查看文章“<%=trim(rs("Title"))%>”"><%if len(rs("title"))>24 then%><%=left(rs("title"),24)%>...<%else%><%=trim(rs("title"))%><%end if%></a><%if request.cookies(Forcast_SN)("ManageKEY")="super" and session("purview")="99999" then%><font color=red>(<%=rs("editor")%>)</font><%end if%></td>
	<td align=center ><%=rs("UpdateTime")%>　</td>
	<td align=center ><%=rs("picnews")%>　</td>
	<td align=center ><%=rs("checkked")%>　</td>
	<td align=center ><input type=hidden name=NewsID value="<%=rs("NewsID")%>">
		<input type=submit value="<%if rs("checkked")="0" then%>通过<%else%>取消<%end if%>" class=button onMouseOver="window.status='按这个按钮<%if rs("checkked")="0" then%>通过<%else%>取消<%end if%>审核文章“<%=trim(rs("Title"))%>”';return true;" onMouseOut="window.status='';return true;" title="按这个按钮<%if rs("checkked")="0" then%>通过<%else%>取消<%end if%>审核文章“<%=trim(rs("Title"))%>”"  style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
	</td>
	<td align=center ><a href=NewsEdit.asp?NewsID=<%=rs("newsid")%>>编辑</a></td>
	<td align=center ><a href=NewsDel.asp?NewsID=<%=rs("newsid")%>>删除</a></td>
</tr>
</form>
				<%
				rs.MoveNext
			end if
		next
		%>
<tr>
	<td colspan=8  align=center height=25>共 <%=total%> 条，当前第 <%=Mypage%>/<%=Maxpages%> 页，每页 <%=MyPageSize%> 条 
		<%
		url="checkNews1.asp?bigclassid="&bigclassid&"&"
		PageNextSize=int((MyPage-1)/PageShowSize)+1
		Pagetpage=int((total-1)/rs.PageSize)+1
		if PageNextSize >1 then
			PagePrev=PageShowSize*(PageNextSize-1)
			Response.write "<a target=_self class=black href='" & Url & "page=" & PagePrev & "' title='上" & PageShowSize & "页'>上一翻页</a> "
			Response.write "<a target=_self class=black href='" & Url & "page=1' title='第1页'>页首</a> "
		end if
		if MyPage-1 > 0 then
			Prev_Page = MyPage - 1
			Response.write "<a target=_self class=black href='" & Url & "page=" & Prev_Page & "' title='第" & Prev_Page & "页'>上一页</a> "
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
				Response.write "<a target=_self class=black href='" & Url & "page=" & PageLink & "'>[" & PageLink & "]</a> "
			else
				Response.Write "<B>["& PageLink &"]</B> "
			end if
			If PageLink = MaxPages Then Exit for
		Next
		if Mypage+1 <=Pagetpage  then
			Next_Page = MyPage + 1
			Response.write "<a target=_self target=_self class=black href='" & Url & "page=" & Next_Page & "' title='第" & Next_Page & "页'>下一页</A>"
		end if
		if MaxPages > PageShowSize*PageNextSize then
			PageNext = PageShowSize * PageNextSize + 1
			Response.write " <A target=_self class=black href='" & Url & "page=" & Pagetpage & "' title='第"& Pagetpage &"页'>页尾</A>"
			Response.write " <a target=_self class=black href='" & Url & "page=" & PageNext & "' title='下" & PageShowSize & "页'>下一翻页</a>"
		End if
		%>
	</td>
</tr>
</table>
</div>
</body>
</html>
<!--#include file=Admin_Bottom.asp-->
		<%
	else
		Show_Err("没有要审核的文章")
		response.end
	End If
	rs4.close
	set rs4=nothing
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
end if%>