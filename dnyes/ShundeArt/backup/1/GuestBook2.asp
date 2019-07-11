<!--#include file="include/conn.asp"-->
<!--#include file="include/config.asp"-->
<!--#include file="admin/char.inc"-->

<%
const MaxPerPage=10
dim totalPut
dim CurrentPage
dim TotalPages
dim a,j

if not isempty(request("page")) then
	currentPage=cint(request("page"))
else
	currentPage=1
end if

if request.cookies("key")="super" then
	aaas=1
	set urs=server.createobject("adodb.recordset")
	sql="select * from "& db_User_Table &" where "& db_User_Name &"='"&Request.cookies("username")&"'"
	urs.open sql,conn,1,3
	if urs.bof or urs.eof then
		aaas=0
	end if
	IF Request.cookies("passwd")<>urs(db_User_Password) THEN
		aaas=0
	END IF
	urs.close
	set urs=nothing
end if
%>

<html>
<head> 
<title>留言本_<%=jjgn%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="news.css" rel=stylesheet type=text/css></head>
<body bgcolor="#FFFFFF" text="#000000" topmargin="0">
<!--#include file="include/top.asp"-->
<table width="760" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
	<tr valign="top">
		<td>
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td valign="top">
						<table width="102%" border="0" cellspacing="0" cellpadding="0">
							<tr bgcolor="CDCCCC">
								<td width="100%" height="25" background="IMAGES/menu-guestbook.gif" >&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;栏目导航&nbsp;&nbsp;<a class=daohang href="./" >网站首页</a>&gt;&gt;<a href="guestbook.asp" class="daohang">雁过留声</a>&gt;&gt;<strong>查看留言</strong></td>
							</tr>
							<tr bgcolor="CDCCCC">
								<td background="IMAGES/menu-guest-l.gif" bgcolor="#FFFFFF">
									<div align="center">
										<a class=daohang href="./" >
											<script language=javascript src=./admin/zongg/ad.asp?i=13></script>
										</a>
									</div>
								</td>
							</tr>
							<tr bgcolor="CDCCCC">
								<td height="25" align="center" background="IMAGES/menu-l-guest.gif" bgcolor="#FFFFFF">
									<a href="guestadd.asp" class=class>我要留言</a> | <a class=class href="javascript:window.location.reload()">刷新</a>
								</td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr valign="top"> 
		<td>
			<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" background="IMAGES/menu-guest-l.gif">
				<tr> 
					<td> 
<%set rs=server.createobject("adodb.recordset")
sql="Select * From "& db_Review_Table &" where NewsId is NULL or NewsId<=0 Order by ReviewId DESC"
rs.open sql,conn,1,1
if rs.eof and rs.bof then
						response.write "<p align=center>还 没 有 任 何 留 言</p>"
else
	totalPut=rs.recordcount
	if currentpage<1 then
		currentpage=1
	end if
	if (currentpage-1)*MaxPerPage>totalput then
		if (totalPut mod MaxPerPage)=0 then
			currentpage=totalPut \ MaxPerPage
		else
			currentpage=totalPut \ MaxPerPage + 1
		end if
	end if
	if currentPage=1 then
		showContent
		showpage totalput,MaxPerPage,"guestbook.asp"
	else
		if (currentPage-1)*MaxPerPage<totalPut then
			rs.move (currentPage-1)*MaxPerPage
			dim bookmark
			bookmark=rs.bookmark
			showContent
			showpage totalput,MaxPerPage,"guestbook.asp"
		else
			currentPage=1
			showContent
			showpage totalput,MaxPerPage,"guestbook.asp"
		end if
	end if
	rs.close
end if
set rs=nothing
conn.close
set conn=nothing

sub showContent
do while not rs.eof
	author=rs("author")
	set urs=server.createobject("adodb.recordset")
	sql="select * from "& db_User_Table &" where "& db_User_Name &"='"&author&"'" 
	urs.open sql,conn,1,3
	if not urs.bof or not urs.eof then
		if UserTableType = "Dvbbs" then
			photowidth=urs(db_User_FaceWidth)	''取注册用户论坛设定的图片片宽度
			photoheight=urs(db_User_FaceHeight)	''取注册用户论坛设定的图片高度 
		end if
		user=true			''是已注册用户
	else 
		user=notreg			''是非注册用户
	end if
	urs.close
	set urs=nothing%>
					</td>	
				</tr>
				<tr>
					<td>
						<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" >
							<tr>
								<td width="20%" rowspan="2" valign="top">
									<table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
										<tr>
											<td width="10" height="105"></td>
											<td height="105" valign="top"><br>
											<img src="<%if trim(rs("face"))<>"" then%>
													<%=rs("face")%>
												<%else%>
													images/Image1.gif 
												<%end if%>" 
												<%if UserTableType = "Dvbbs" then	'取用户头像,如为注册用户,则定义图片显示宽、高分辨率%>
													<%if user=true then%>
														width="<%=photowidth%>" height="<%=photoheight%>" 
													<%end if%>
												<%end if%> >
												<br>
												<br>
												身份：<font color="red">
												<%if rs("shengfen")<>"" then%>
													<%=rs("shengfen")%> 
												<%else%>
													保密
												<%end if%>
												</font>
												<br>
												姓名：<%=rs("author")%>
												<br> 
												<%if rs("from")<>"" then %>
													来自：<%=rs("from")%> 
												<%end if%>
												<br> 
												<%if aaas=1 then %>
													IP：<%=rs("reviewip")%> 
												<%end if%>
											</td>
										</tr>
									</table>
								</td>
								<td valign="top">
									<table width="100%" border="0" cellspacing="0" cellpadding="0" >
										<tr>
											<td width="10"></td>
											<td> 
											<%if rs("email")<>"" then%>
												<a href="mailto:<%=rs("email")%>"><img src="images/EMAIL.GIF" border="0"></a>
											<%else%>
												<img src="images/EMAIL.GIF" border="0">
											<%end if%>
												&nbsp;
											<%if rs("homepage")<>"" then%>
												<a href="<%=rs("homepage")%>" target="_blank"><img src="images/homepage.gif" border="0"></a>
											<%else%>
												<img src="images/homepage.gif" border="0">
											<%end if%>
												&nbsp;
											<% if rs("oicq")<>"" then%>
												<a href="http://search.tencent.com/cgi-bin/friend/user_show_info?ln=<%=rs("oicq")%>" ><img src="images/OICQ.gif" border="0"></a>
											<%else%>
												<img src="images/OICQ.gif" border="0">
											<%end if%>
											
											<a href="GuestEdit.asp?ReviewId=<%=rs("reviewid")%>"><img src="images/EDIT.GIF" border="0"></a> 
											
											<%if aaas=1 then %>
												<a href="guestreply.asp?reviewid=<%=rs("reviewid")%>"><img src="images/reply.GIF" border="0"></a> 
												<a href="guestdel.asp?reviewid=<%=rs("reviewid")%>"><img src="images/delete.GIF" border="0"></a> 
											<%end if%>
										
											<%if rs("newsid")<>"" then%>
												<a target="_blank" href="readnews.asp?newsid=<%=rs("newsid")%>"><img src="images/reply1.GIF" width="0" height="0" border="0"></a> 
											<%end if%>
											&nbsp;
											</td>
										</tr>
									</table>
								</td>
							</tr>
							<tr>
								<td valign="top">
									<table width="98%" border="0" cellspacing="0" cellpadding="0" style="table-layout:fixed; word-break:break-all;border-collapse: collapse">
										<tr>
											<td width="10" ></td>
											<td >
												<%=rs("title")%> <p><b>留言内容：</b><br>
												
												
												<%if rs("NewsID")=-1 then	'-1 表示为悄悄话
													if aaas=1 then
														response.write rs("content")
													else
														response.write "<center><BR><font color=red>这是<B>悄悄话</B>,只有管理员才看的见...</font></center>"
													end if
												else
													response.write rs("content")
												end if%>
												</p>
												
												<%if rs("reply")<>"" then %>
													<p>
													<font color="#FF0000"><b><img src="IMAGES/base.gif" align="absmiddle">&nbsp;站长回复：</b><br>
													<%=rs("reply")%></font></p>
												<%end if%>
											</td>
										</tr>
									</table>
								</td>
							</tr>
							<tr>
								<td height="25" colspan="2" align="right" valign="top">&nbsp;&nbsp;发表于：<%=rs("updatetime")%>&nbsp;&nbsp;</td>
							</tr>
							<tr>
								<td height="25" colspan="2"><img src="IMAGES/menu-l-guest.gif" width="760" height="25"></td>
							</tr>
						</table>
					</td>
				</tr>
				<tr> 
					<td> 
	<%a=a+1
	if a>=MaxPerPage then exit do
	rs.movenext
loop
end sub%>
					</td>
				</tr>
				<tr>
					<td></td>
				</tr>
			</table>
		</td>
	</tr>
</table>
<table width="760" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="efefef">
	<tr> 
		<td align="center" background="IMAGES/menu-guest-l.gif" bgcolor="#FFFFFF">
			<a href="guestadd.asp" class=class>我要留言</a> | <a class=class href="javascript:window.location.reload()">刷新</a> 
			<%function showpage(totalnumber,maxperpage,filename)
			dim n
			if totalnumber mod maxperpage=0 then
				n=totalnumber \ maxperpage
			else
				n=totalnumber \ maxperpage + 1
			end if
			response.write "<form method=post action="&filename&">"
			response.write "<p align=center>"
			if CurrentPage<2 then
				response.write "<font color='#b9b9b9' style='font-family: 宋体; font-size: 9pt'>首页 上一页</font> "
			else
				response.write "<a class=class href="&filename&"?page=1>首页</a> "
				response.write "<a class=class href="&filename&"?page="&CurrentPage-1&">上一页</a> "
			end if
			if n-currentpage<1 then
				response.write "<font color='#b9b9b9' style='font-family: 宋体; font-size: 9pt'>下一页 尾页</font>"
			else
				response.write "<a class=class href="&filename&"?page="&(CurrentPage+1)&">下一页</a> "
				response.write "<a class=class href="&filename&"?page="&n&">尾页</a>"
			end if
				response.write "<font color='#000080' style='font-family: 宋体; font-size: 9pt'> 共<b>"&totalnumber&"</b>条留言 <b>"&maxperpage&"</b> 条留言/页</font> "
				response.write "<font color='#000080' style='font-family: 宋体; font-size: 9pt'>转到：</font><input class=smallInput type='text' name='page' size=4 maxlength=10 value="&Currentpage&">"
				response.write "<input class=buttonface type='submit' value='Go' name='cndok'></p></form>"
			end function%>
		</td>
	</tr>
	<tr>
		<td height="19" align="center" background="IMAGES/menu-guest-t.gif" bgcolor="#FFFFFF"></td>
	</tr>
</table>

<!--#include file="include/bottom.asp"-->
</body>
</html>