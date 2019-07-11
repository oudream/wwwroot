<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp" -->
<!--#include file="ChkManage.asp" -->
<%
if not(request.cookies(Forcast_SN)("ManageKEY")="super" or request.cookies(Forcast_SN)("ManageKEY")="typemaster" or request.cookies(Forcast_SN)("ManageKEY")="bigmaster") then
	Show_Err("对不起，您的后台管理权限不够操作！")
	response.end
else
	dim oskey
	dim rs,tsql
	dim rst
	if (request.form("action")="del") and (request.form("selectdel")<>"") then
		if (request.cookies(Forcast_SN)("ManageKEY")="super") then
			conn.execute "delete from "& db_User_Table &" where "& db_User_ID &" in ("&request.form("selectdel")&")"
			response.Redirect "UserManage.asp"
		else
			Show_Err("对不起，您不能进行用户的批理删除！")
			response.end
		end if
	end if
	oskey=request("oskey")
	set rst=server.CreateObject("ADODB.RecordSet")
	if request.cookies(Forcast_SN)("key")="super" then
		if oskey="" then
			rst.Source="select * from "& db_User_Table &" order by "& db_User_ID &""
		else
			rst.Source="select * from "& db_User_Table &" where oskey='"&oskey&"' order by "& db_User_ID &""
		end if
	else
		if oskey="" then
			rst.Source="select * from "& db_User_Table &" where adder='"&request.cookies(Forcast_SN)("username")&"' order by "& db_User_ID &""
		else
			rst.Source="select * from "& db_User_Table &" where adder='"&request.cookies(Forcast_SN)("username")&"' and oskey='"&oskey&"' order by "& db_User_ID &""
		end if
	end if

	PageShowSize = 10		'每页显示多少个页
	MyPageSize   = 10		'每页显示多少用户

	If Not IsNumeric(Request("page")) Or IsEmpty(Request("page")) Or Request("page") <=0 Then
		MyPage=1
	Else
		MyPage=Int(Abs(Request("page")))
	End if

	rst.Open rst.Source,ConnUser,3,1	
	If Not rst.eof then
		rst.PageSize     = MyPageSize
		MaxPages         = rst.PageCount
		rst.absolutepage = MyPage
		total            = rst.RecordCount
		%>
<html><head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=site.css rel=stylesheet>
<title><%=copyright%><%=version%>&nbsp;<%=ver%> - 用户管理</title>
<script language=javascript>
function CheckFormUserLogin()
{
	if(document.SeahUser.UserName.value=="")
	{
		alert("请输入用户名！");
		document.SeahUser.UserName.focus();
		return false;
	}
}
</script>
<script language="JavaScript">
<!--
function CheckAll(form){
	for (var i=0;i<form.elements.length;i++) {
		var e = form.elements[i];
		if (e.name != 'chkall') e.checked = form.chkall.checked; 
	}
}
//-->
</script>
</head>
<body topmargin="0">
<!--#include file=Admin_Top.asp-->
<table border="1" cellpadding="3" cellspacing="0" style="border-collapse: collapse" bordercolor="#C0C0C0" width="100%" id="AutoNumber1">
<tr align="center"  height=25>
	<td colspan="12" height=25 class="TDtop">┊ <B>用户管理</B> ┊ </td>
</tr>
<tr align="center"  height=25>
	<td colspan="12" height=25 >
		<p align=center>用户列表(<font color=red>红色</font>表示当前用户)，点击用户名可查看更多信息 <a href="useradd1.asp">添加用户</a>  <a href="UserManage.asp">全部用户</a>
	</td>
</tr>
<tr align="center" class="TDtop1"  height=20> 
	<td width="4%">ID号</td>
	<td width="11%">用户名</td>
	<td width="10%">权限</td>
	<td width="14%">最后登录地址</td>
	<td width="15%">最后登录时间</td>
	<td width="5%">登录</td>
	<td width="6%">发文数</td>
	<td width="6%">修改</td>
	<td width="6%">删除</td>
	<td width="9%">会员等级</td>
	<td width="10%">会员状态</td>
	<td width="4%" >选择</td>
  </tr> 
  <%response.write "<form name='formdel' method='post' action='UserManage.asp'>"
  for i=1 to rst.PageSize
		if not rst.eof then
			if rst("purview")<>"99999" or request.cookies(Forcast_SN)("purview")="99999" then
				response.write "<tr align='center' height=20>"& chr(10)
				if (request.cookies(Forcast_SN)("key")="bigmaster" and rst("oskey")="smallmaster") or (request.cookies(Forcast_SN)("key")="typemaster" and (rst("oskey")="bigmaster" or rst("oskey")="smallmaster")) or request.cookies(Forcast_SN)("key")="super" then
		
					response.write "<td align='center'>" &(rst(db_User_ID))& "</td>"
	
					response.write "<td align='center' bgcolor='#FFFFFF'>"
					response.write "<a href='list.asp?id=" &(rst(db_User_ID)) &"' title=" &(rst(db_User_Password)) &">"
					if rst(db_User_Name)=request.cookies(Forcast_SN)("username") then
						response.write "<font color=red>" &(rst("Username"))& "</font>"
					else
						response.write (rst("Username"))
					end if
					response.write "</a>"
					response.write "</td>"& chr(10)
	
					response.write "<td align='center' bgcolor='#FFFFFF'>"
					
					if rst("oskey")="super" and rst("purview")="99999" then
						response.write "<font color='red'>超级用户</font>"
					else
						oskey1=rst("oskey")
						select case oskey1
							case "super" response.write "系统用户"
							case "check" response.write "<a href='UserManage.asp?oskey="& oskey1 &"'>审核用户</a>"
							case "typemaster" response.write "<a href='UserManage.asp?oskey="& oskey1 &"'>总栏用户</a>"
							case "bigmaster" response.write "<a href='UserManage.asp?oskey="& oskey1 &"'>大类用户</a>"
							case "smallmaster" response.write "<a href='UserManage.asp?oskey="& oskey1 &"'>小类用户</a>"
							case "selfreg" response.write "<a href='UserManage.asp?oskey="& oskey1 &"'>注册用户</a>"
						end select
					end if
					response.write "</td>"& chr(10)
	
					response.write "<td align='center'>" &(rst(db_User_LastLoginIP))& "</td>"& chr(10)
	
					response.write "<td >" &(rst(db_User_LastLoginTime))& "</td>"& chr(10)
	
					response.write "<td align='center'>" &(rst(db_User_LoginTimes))& "</td>"& chr(10)
	
					response.write "<td align='center'>" &(rst("number"))& "</td>"& chr(10)
	
					response.write "<td align='center' bgcolor='#FFFFFF'>"
					response.write "<a href='UserEdit.asp?ID=" &(rst(db_User_ID))& "' title='修改用户“" &(rst(db_User_Name))& "”的属性'>修改</a>"
					response.write "</td>"& chr(10)
					
					response.write "<td align='center' bgcolor='#FFFFFF'>"
					if trim(request.cookies(Forcast_SN)("username"))=trim(rst(db_User_Name)) then
						response.write "----"
					else
						response.write "<a href='userdel.asp?id=" &(rst(db_User_ID))& "&amp;name=del' title='删除用户“" &(rst(db_User_Name))& "”'>删除</a>"
					end if
					response.write "</td>"& chr(10)
				end if
				response.write "<td>"
				if rst("oskey")="selfreg" then
					reglevel1=rst("reglevel")
					select case reglevel1
						case 1
							response.write "普通<br>"
							response.write "<a href=UserLevel.asp?ID=" &(rst(db_User_ID))& chr(38) &"reglevel=2 title='现等级为‘普通’，点击将置为‘高级’'>高级</a><br>"
							response.write "<a href=UserLevel.asp?ID=" &(rst(db_User_ID))& chr(38) &"reglevel=3 title='现等级为‘普通’，点击将置为‘特级’'>特级</a><br>"
						case 2
							response.write "<a href=UserLevel.asp?ID=" &(rst(db_User_ID))& chr(38) &"reglevel=1 title='现等级为‘高级’，点击将置为‘普通’'>普通</a><br>"
							response.write "高级<br>"
							response.write "<a href=UserLevel.asp?ID=" &(rst(db_User_ID))& chr(38) &"reglevel=3 title='现等级为‘高级’，点击将置为‘特级’'>特级</a><br>"
						case 3
							response.write "<a href=UserLevel.asp?ID=" &(rst(db_User_ID))& chr(38) &"reglevel=1 title='现等级为‘特级’，点击将置为‘普通’'>普通</a><br>"
							response.write "<a href=UserLevel.asp?ID=" &(rst(db_User_ID))& chr(38) &"reglevel=2 title='现等级为‘特级’，点击将置为‘高级’'>高级</a><br>"
							response.write "特级<br>"
					end select
				end if
				response.write "</td>"
				response.write "<td>"
				if rst("oskey")="selfreg" then
					jingyong1=rst("jingyong")
					response.write "<a href=UserJingYong.asp?ID=" &(rst(db_User_ID))&chr(38)& "jingyong="
					Select Case jingyong1
						case 0 response.write "1 title='现状态为‘启用’，点击将禁用该用户'>启用</a>"
						case 1 response.write "0 title='现状态为‘禁用’，点击将启用该用户'>禁用</a>"
					end Select
				end if
				response.write "</td>"
					response.write "<td align='center'>"
					if trim(request.cookies(Forcast_SN)("username"))=trim(rst(db_User_Name)) or rst("oskey")<>"selfreg" then
						response.write "--"
					else
						response.write "<input type=checkbox name=selectdel id=selectdel value=" &(rst(db_User_ID))& ">"
					end if
					response.write "</td>"
				response.write "</tr>"& chr(10)
			end if
			rst.MoveNext
		end if
	next
	response.write "<tr align='right' ><td colspan=13  height=20>"%>
	<input type=hidden name=action value='del'>
	<input type=checkbox name=chkall value=on onclick="CheckAll(this.form)" style="font-size: 9pt;  color: #000000; background-color: #ffff; solid #EAEAF4">选中所有显示用户&nbsp;
	<input type=submit name=action1 onclick="{if(confirm('删除选定的用户吗？')){this.document.check.submit();return true;}return false;}" value="删除用户" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
	<%response.write "</td></tr>"& chr(10)
	response.write "</form>"& chr(10)
	rst.close
	set rst=nothing%>
<tr height=20>
<td colspan=12 class="TDtop1" >
<p align=center height=25>共 <%=total%> 名，当前第 <%=Mypage%>/<%=Maxpages%> 页，每页 <%=MyPageSize%> 名 
	<%
	if oskey="" then
		url="UserManage.asp?"
	else
		url="UserManage.asp?oskey="& oskey &"&"
	end if
	PageNextSize=int((MyPage-1)/PageShowSize)+1
	Pagetpage=int((total-1)/rst.PageSize)+1
	
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
	else
		PageSizeShow = Maxpages-PageShowSize*(PageNextSize-1)
	end if
	
	if PageSizeShow < 1 Then PageSizeShow = 1
	for PageCounterSize=1 to PageSizeShow
		PageLink = (PageCounterSize+PageNextSize*PageShowSize)-PageShowSize
		if PageLink <> MyPage Then
			Response.write "<a class=black href='" & Url & "page=" & PageLink & "'>[" & PageLink & "]</a> "
		else
			Response.Write "<B>["& PageLink &"]</B> "
		end if
		if PageLink = MaxPages Then Exit for
	next
	
	if Mypage+1 <=Pagetpage  then
		Next_Page = MyPage + 1
		Response.write "<a class=black href='" & Url & "page=" & Next_Page & "' title='第" & Next_Page & "页'>下一页</A>"
	end if
	
	if MaxPages > PageShowSize*PageNextSize then
		PageNext = PageShowSize * PageNextSize + 1
		Response.write " <A class=black href='" & Url & "page=" & Pagetpage & "' title='第"& Pagetpage &"页'>页尾</A>"
		Response.write " <a class=black href='" & Url & "page=" & PageNext & "' title='下" & PageShowSize & "页'>下一翻页</a>"
	end if
	%>
</td>
</tr>
</table>
<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" bordercolor="#C0C0C0" width="50%" id="AutoNumber1">
<tr>
<td>
<font color="#FF0000">
<li><p>总栏用户设置：在添加总栏用户后请在“总栏管理”再设置一下。</li>
<li><p>大类用户设置：在添加大类用户后请在“大类管理”再设置一下。</li>
<li><p>小类用户设置：在添加小类用户后请在“小类设置”设置。</li>
<li><p>系统用户具有所有权限。</li>
<li><p>文章审核员可管理所有大类文章，不必在大类中添加。</li>
<li><p>注册用户可在所有类别中发表文章。</li></font>
</td>
</tr>
</table>
	<%dim SeachUserName,SeachModel,Sea_Mod
	SeachUserName=CheckStr(request.form("UserName"))
	SeachModel=CheckStr(request.form("SeachModel"))
	%>
<table border="1" cellpadding="3" cellspacing="0" style="border-collapse: collapse" bordercolor="#C0C0C0" width="100%" id="AutoNumber1">
<form method="POST" action="usermanage.asp" name="SeahUser"  onSubmit="return CheckFormUserLogin();">
<tr align="center" height=20>
<td colspan=13>
<input type=hidden name="page" value=<%=MyPage%>>
用户名：<input name="UserName" <%if SeachUserName<>"" then%>value=<%=SeachUserName%><%end if%> size="15" font face="宋体" style="font-size: 9pt; background-color:#EAEAF4">
<input type="submit" name="Submit" value="查找" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'" title="查找用户">&nbsp;&nbsp;
	<%if SeachModel="" or SeachModel=1 then
		Response.write "<input type=radio name=SeachModel checked  value=1>精确&nbsp;"
		Response.write "<input type=radio name=SeachModel value=0>模糊"
		rstsql="select * from "& db_User_Table &" where "& db_User_Name &" like '"& SeachUserName &"' order by "& db_User_ID
		Sea_Mod="精确查找："
	else
		Response.write "<input type=radio name=SeachModel value=1>精确&nbsp;"
		Response.write "<input type=radio name=SeachModel checked value=0>模糊"
		rstsql="select * from "& db_User_Table &" where "& db_User_Name &" like '%"& SeachUserName &"%' order by "& db_User_ID
		Sea_Mod="模糊查找："
	end if%>
</td>
</tr>
</form>
	<%
	if SeachUserName<>"" then
		set rst1=server.CreateObject("ADODB.RecordSet")
		rst1.Open rstsql,ConnUser,3,1
		if not rst1.eof then
			%>
			<form method="POST" action="UserManage.asp" name="SeahUsertest" >
			<tr align="center" class="TDtop">
				<td width="4%">ID号</td>
				<td width="11%">用户名</td>
				<td width="10%">权限</td>
				<td width="14%">最后登录地址</td>
				<td width="15%">最后登录时间</td>
				<td width="5%">登录</td>
				<td width="6%">发文数</td>
				<td width="6%">修改</td>
				<td width="6%">删除</td>
				<td width="9%">会员等级</td>
				<td width="10%">会员状态</td>
				<td width="4%" >选择</td>
			</tr> 
		<%do while not rst1.eof
			if rst1("purview")<>"99999" or request.cookies(Forcast_SN)("purview")="99999" then
				response.write "<tr align='center' height=20>"& chr(10)
				if (request.cookies(Forcast_SN)("key")="bigmaster" and rst1("oskey")="smallmaster") or (request.cookies(Forcast_SN)("key")="typemaster" and (rst1("oskey")="bigmaster" or rst1("oskey")="smallmaster")) or request.cookies(Forcast_SN)("key")="super" then
					response.write "<td align='center'>" &(rst1(db_User_ID))& "</td>"
					response.write "<td align='center' bgcolor='#FFFFFF'>"
					response.write "<a href='list.asp?id=" &(rst1(db_User_ID)) &"' title=" &(rst1(db_User_Password)) &">"
					if rst1(db_User_Name)=request.cookies(Forcast_SN)("username") then
						response.write "<font color=red>" &(rst1("Username"))& "</font>"
					else
						response.write (rst1("Username"))
					end if
					response.write "</a>"
					response.write "</td>"& chr(10)
		
					response.write "<td align='center' bgcolor='#FFFFFF'>"
					if rst1("oskey")="super" and rst1("purview")="99999" then
						response.write "<font color='red'>超级用户</font>"
					else
						oskey2=rst1("oskey")
						select case oskey2
							case "super" response.write "系统用户"
							case "check" response.write "审核用户"
							case "typemaster" response.write "总栏用户"
							case "bigmaster" response.write "大类用户"
							case "smallmaster" response.write "小类用户"
							case "selfreg" response.write "注册用户"
						end select
					end if
					response.write "</td>"& chr(10)
		
					response.write "<td align='center'>" &(rst1(db_User_LastLoginIP))& "</td>"& chr(10)
		
					response.write "<td>" &(rst1(db_User_LastLoginTime))& "</td>"& chr(10)
		
					response.write "<td align='center'>" &(rst1(db_User_LoginTimes))& "</td>"& chr(10)
		
					response.write "<td align='center'>" &(rst1("number"))& "</td>"& chr(10)
		
					response.write "<td align='center' bgcolor='#FFFFFF'>"
					response.write "<a href='UserEdit.asp?ID=" &(rst1(db_User_ID))& "' title='修改用户“" &(rst1(db_User_Name))& "”的属性'>修改</a>"
					response.write "</td>"& chr(10)
					
					response.write "<td align='center' bgcolor='#FFFFFF'>"
					if trim(request.cookies(Forcast_SN)("username"))=trim(rst1(db_User_Name)) then
						response.write "----"
					else
						response.write "<a href='userdel.asp?id=" &(rst1(db_User_ID))& "&amp;name=del' title='删除用户“" &(rst1(db_User_Name))& "”'>删除</a>"
					end if
					response.write "</td>"& chr(10)
				end if
		
				response.write "<td>"
				if rst1("oskey")="selfreg" then
					reglevel2=rst1("reglevel")
					select case reglevel2
						case 1
							response.write "普通<br>"
							response.write "<a href=UserLevel.asp?ID=" &(rst1(db_User_ID))& chr(38) &"reglevel=2 title='现等级为‘普通’，点击将置为‘高级’'>高级</a><br>"
							response.write "<a href=UserLevel.asp?ID=" &(rst1(db_User_ID))& chr(38) &"reglevel=3 title='现等级为‘普通’，点击将置为‘特级’'>特级</a><br>"
						case 2
							response.write "<a href=UserLevel.asp?ID=" &(rst1(db_User_ID))& chr(38) &"reglevel=1 title='现等级为‘高级’，点击将置为‘普通’'>普通</a><br>"
							response.write "高级<br>"
							response.write "<a href=UserLevel.asp?ID=" &(rst1(db_User_ID))& chr(38) &"reglevel=3 title='现等级为‘高级’，点击将置为‘特级’'>特级</a><br>"
						case 3
							response.write "<a href=UserLevel.asp?ID=" &(rst1(db_User_ID))& chr(38) &"reglevel=1 title='现等级为‘特级’，点击将置为‘普通’'>普通</a><br>"
							response.write "<a href=UserLevel.asp?ID=" &(rst1(db_User_ID))& chr(38) &"reglevel=2 title='现等级为‘特级’，点击将置为‘高级’'>高级</a><br>"
							response.write "特级<br>"
					end select
				end if
				response.write "</td>"
				response.write "<td>"
				if rst1("oskey")="selfreg" then
					jingyong2=rst1("jingyong")
					response.write "<a href=UserJingYong.asp?ID=" &(rst1(db_User_ID))&chr(38)& "jingyong="
					Select Case jingyong2
						case 0 response.write "1 title='现状态为‘启用’，点击将禁用该用户'>启用</a>"
						case 1 response.write "0 title='现状态为‘禁用’，点击将启用该用户'>禁用</a>"
					end Select
				end if
				response.write "</td>"
					response.write "<td align='center'>"
					if trim(request.cookies(Forcast_SN)("username"))=trim(rst1(db_User_Name)) or rst1("oskey")<>"selfreg" then
						response.write "--"
					else
						response.write "<input type=checkbox name=selectdel id=selectdel value=" &(rst1(db_User_ID))& ">"
					end if
					response.write "</td>"
				response.write "</tr>"& chr(10)
			end if
			rst1.MoveNext
			loop
			response.write "<tr align='right' ><td colspan=13  height=20>"%>
		<input type=hidden name=action value='del'>
		<input type=checkbox name=chkall value=on onclick="CheckAll(this.form)" style="font-size: 9pt;  color: #000000; background-color: #ffff; solid #EAEAF4">选中所有显示用户&nbsp;
		<input type=submit name=action1 onclick="{if(confirm('删除选定的用户吗？')){this.document.check.submit();return true;}return false;}" value="删除用户" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
			<%response.write "</td></tr>"& chr(10)
			response.write "</form>"& chr(10)
		else
			response.write "<tr align='center'><td colspan=13 bgcolor='ffb100' height=90><font size=4>"& Sea_Mod &"未找到用户["& SeachUserName &"]</font></td></tr>"& chr(10)
		end if
		rst1.close
		set rst1=nothing
	end if%>
	</table>
</body>
</html>
<!--#include file=Admin_Bottom.asp-->
	<%else
		Show_Err("<p>当前没有您可管理的用户！</p><p><a href='UserAdd1.asp'>请添加用户</a></p>")
		response.end
	end if
end if
conn.close
set conn=nothing
ConnUser.close
set ConnUser=nothing%>