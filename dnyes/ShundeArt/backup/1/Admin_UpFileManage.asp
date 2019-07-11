<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp" -->
<!--#include file="ChkManage.asp" -->
<%
if request.cookies(Forcast_SN)("ManageKEY")<>"super" or request.cookies(Forcast_SN)("ManagePurview")<>"99999" then
	Show_Err("对不起，您的后台管理权限不够！")
	response.end
else
	if system="1" or request.cookies(Forcast_SN)("ManagePurview")="99999" then
		
		If Not IsNumeric(Request("page")) Or IsEmpty(Request("page")) Or Request("page") <=0 Then
			MyPage=1
		Else
			MyPage=Int(Abs(Request("page")))
		End if
		
		if Request.form("Check_Save")="ok" then
			response.cookies(Forcast_SN)("C_Attach")=Request.form("Check_Attach")
			response.cookies(Forcast_SN)("C_News")=Request.form("Check_News")
			response.cookies(Forcast_SN)("C_UserFace")=Request.form("Check_UserFace")
		end if

		Check_Attach=Request.cookies(Forcast_SN)("C_Attach")
		Check_News=Request.cookies(Forcast_SN)("C_News")
		Check_UserFace=Request.cookies(Forcast_SN)("C_UserFace")

		Function Get_Attribute(int_AttribCode,str_ListSeparator)
		  '--- 传入参数 ---
		  'int_AttribCode:属性代码参数-整型
		  'str_ListSeparator:分隔符号-字符型
		  'Get_Attribute=""
		  'int_AttribCode1=int_AttribCode
		  '"0 普通文件|128 压缩文件|64 链接或快捷方式|32 存档文件|16 目录|8 磁盘驱动器卷标|4 系统文件|2 隐藏文件|1 只读文件", "|", -1, 1)

		  str_AttribType = Split("普通文件|压缩文件|链接或快捷方式|存档文件|目录|磁盘驱动器卷标|系统文件|隐藏文件|只读文件", "|", -1, 1)
		  int_AttribCode=CInt(int_AttribCode)

		  if int_AttribCode=0 or int_AttribCode>256 then
				Get_Attribute=str_AttribType(0)
			Else
			  int_Gene=128
			  ii=1
			  do until int_Gene<1
			  	if int_AttribCode-int_Gene>=0 then
						if int_Gene/2<1 or int_AttribCode-int_Gene=0 then
							str_ListSeparator=""
						end if
						Get_Attribute=Get_Attribute & str_AttribType(ii) & str_ListSeparator
			  		int_AttribCode=int_AttribCode-int_Gene
			  	end if
		  		int_Gene=int_Gene/2
		  		ii=ii+1
				loop			  
			end if
			'Get_Attribute=int_AttribCode & Get_Attribute
		End function

		'FileUploadPath	'实际图片目录
		ThisNameFile="Admin_UpFileManage.asp"   '定义本文件的文件名
		urlpath="http://"&Request.ServerVariables("SERVER_NAME")
		
		dim cpath,lpath,r_path,l_path
		r_path=Request("path")
		l_path=left(Request.ServerVariables("PATH_INFO"),instrrev(Request.ServerVariables("PATH_INFO"),"/"))&FileUploadPath

		set fsoBrowse=CreateObject("Scripting.FileSystemObject")
	
		if InStr(1,r_path, l_path,1)<>1 then
			r_path=""
		end if

		'检查路径中是否含有“.”，防止以“..”方式返回上级目录信息。 
		if InStr(1,r_path, ".",1)>0 or InStr(1,r_path, "%2e",1)>0 or InStr(1,r_path, "%2E",1)>0 then
			r_path=""
			response.write "<script language=javascript>alert('警告！路径中含有非法字符。');</script>"
		end if
		
		if r_path="" then
			lpath=l_path
		else
			if right(r_path,1)<>"/" then  '在目录后加(/)
				lpath=r_path&"/"
			else
				lpath=r_path
			end if
		end if
		if Request("attrib")="true" then
			cpath=lpath
			attrib="true"
		else
			cpath=Server.MapPath(lpath)
			attrib=""
 			if request.form("pathfile")<>"" then
				for i = 1 to request.form("pathfile").count
					pathfile=request.form("pathfile")(i)
					whichfile=Server.MapPath(pathfile)
				  Set fs = CreateObject("Scripting.FileSystemObject")
				  Set thisfile = fs.GetFile(whichfile)
				  thisfile.delete true
				next
				Response.redirect ThisNameFile &"?path="& r_path &"&page="& MyPage
		  end if
		end if
		%>
<html><head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="site.css">
<title><%=copyright%><%=version%>&nbsp;<%=ver%> - 文件管理</title>
<script>
function IsDigit()
{
  return ((event.keyCode >= 48) && (event.keyCode <= 57));
}
</script>
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
<body topmargin="0" oncontextmenu="self.event.returnValue=false">
<!--#include file=Admin_Top.asp-->
<div align="center">
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#C0C0C0" width="100%" id="AutoNumber1">
<form action='Admin_UpFileManage.asp' method=post name=check>
<tr class="TDtop">
	<td height="25" colspan="5" align="center" valign="middle" >┊ <B>所 有 文 件 列 表 <font color=red >/ 慎 重 操 作 /</font></B> ┊ </td>
</tr>
</table>
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#C0C0C0" width="100%" id="AutoNumber1">
<tr>
	<td width='30%' align='center' height=20 bgcolor='#eeeeee'>文件名</td>
	<td width='30%' align='center' bgcolor='#eeeeee'>引用验证</td>
	<td width='15%' align='center' bgcolor='#eeeeee'>文件大小</td>
	<td width='15%' align='center' bgcolor='#eeeeee'>文件属性</td>
	<td width='10%' align='center' bgcolor='#eeeeee'>功能</td>
</tr>
</table>
		<%
		dim theFolder
		dim theFiles
		dim theSubFolders
		dim contain
		dim cont_search
		dim rs
		contain=""
		cont_search=0
		
		if fsoBrowse.FolderExists(cpath) then
			Set theFolder=fsoBrowse.GetFolder(cpath)
			Set theFiles=theFolder.files
			Set theSubFolders=theFolder.SubFolders

			Set rs = Server.CreateObject("ADODB.Recordset")

			PageShowSize		= 10		'每页显示多少个页链接
			MyPageSize			= 15		'每页显示多少个文件
			total						= theFiles.count
			MaxPages				= int((total-1)/MyPageSize)+1

			i = 1

			for each x in theFiles
				if i>(MyPageSize*(Mypage-1)) and i<(MyPageSize*Mypage+1) then
					if Request("attrib")="true" then
						showstring="<strong>"&x.Name&"</strong>"
					else
						showstring="<a href='"&urlpath&lpath&x.Name&"' target='_blank'><strong>"&x.Name&"</strong></a>"
					end if
					
					if Check_Attach="1" then
						sql ="select * from "& db_Attach_Table &" where filename like '%"& x.name &"%'" 
						rs.open sql,conn,1,1
						if not rs.eof then
							do until rs.eof
								'在上传附件表中搜索是否有上传记录
								if rs("filename")=x.name then
									contain=contain&"<a href='ReadNews.asp?NewsID="& rs("newsid") &"' target='_blank' title='在上传附件表中显示该附件在ID号为"& rs("newsid") &"的文章中被引用'>上传记录：" & rs("newsid") & "号文章</a><br>"
								end if
								rs.MoveNext
							loop
						end if
						rs.close
					end if
					
					if Check_News="1" then					
						sql ="select * from "& db_News_Table &" where Content like '%"& x.name &"%'"
						rs.open sql,conn,1,1
						if not rs.eof then
							do until rs.eof
								'在新闻的具体内容中搜索是否被引用
								cont_search = Instr(1, rs("Content"),x.name, 1)
								if cont_search>0 then
									contain=contain&"<a href='ReadNews.asp?NewsID="& rs("newsid") &"' target='_blank' title='该附件在ID号为"& rs("newsid") &"的文章内容中被引用'><font color='green'>新闻内容：" & rs("newsid") & "号文章第"& cont_search &"字处</font></a><br>"
								end if
								rs.MoveNext
							loop
						end if
						rs.close
					end if

					if Check_UserFace="1" then
						sql ="select * from "& db_User_Table &" where "& db_User_Face &" like '%"& x.name &"%'"
						rs.open sql,ConnUser,1,1
						if not rs.eof then
							do until rs.eof
								'在用户表中搜索文件是否是头象引用
								cont_search = Instr(1, rs(db_User_Face),x.name, 1)
								if cont_search>0 then
									contain=contain&"<a href='User.asp?User="& rs(db_User_Name) &"' target='_blank' title='该附件为用户"& rs(db_User_Name) &"的头象文件'><font color='green'>用户"& rs(db_User_Name) &"的头象文件</font></a><br>"
								end if
								rs.MoveNext
							loop
						end if
						rs.close
					end if
					
					if contain="" then
						contain="无引用或不在选择的验证项中"
					end if

					response.write "<table border='1' cellpadding='0' cellspacing='0' style='border-collapse: collapse' bordercolor='#C0C0C0' width='100%' id='AutoNumber1'>"
					response.write "<tr>"
					response.write "<td width='30%' bgcolor=#ffffff height=20 >□"&showstring&"</td>"
					response.write "<td width='30%' align='center' bgcolor=#ffffff>"&contain&"</td>"
					response.write "<td width='15%' align='right' bgcolor=#ffffff>"&x.size&"字节&nbsp;&nbsp;</td>"
					response.write "<td width='15%' align='center' bgcolor=#ffffff><a href='#' title='"&"文件类型："&x.type&chr(10)&"文件属性："&Get_Attribute(x.Attributes,"、")&chr(10)&"创建(上传)："&x.DateCreated&chr(10)&"最近访问："&x.DateLastAccessed&chr(10)&"最后修改："&x.DateLastModified&"'>属性</a></td>"
					response.write "<td width='10%' align='center' bgcolor=#ffffff><input name='pathfile' type='checkbox' id='pathfile' value='"&lpath&x.Name&"'></td>"
					response.write "</tr>"
					response.write "</table>"
					contain=""
					cont_search=0
				end if
				if i>(MyPageSize*Mypage) then
					exit for
				end if
				i=i+1
			next
		end if
		%>
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#C0C0C0" width="100%" id="AutoNumber1">
<tr>
	<td height="20" colspan="5">当前文件目录：<%=urlpath%>/<%
		Dim Path_String, Path_Array,Path_Url
		Path_Url=""
		Path_String = Split(lpath, "/", -1, 1)
		for each Path_Array in Path_String
			if Path_Array<>"" then
				Path_Url=Path_Url &"/"& Path_Array
				response.write "<a href='"& ThisNameFile &"?path="& Path_Url &"' title='点击查看 ["& urlpath & Path_Url &"] 目录'><font color=green>"& Path_Array &"/</font></a>"
			end if
		next
		%>
	</td>
</tr>
<tr>
	<td height="20" colspan="5">子目录：<%
		for each SubF in theSubFolders%>
		<a href='<%=ThisNameFile&"?path="& lpath & SubF.name %>'  title='点击查看 [<%=lpath & SubF.name%>] 目录'>〈<%= SubF.name %>〉</a>
		<%next%>
	</td>
</tr>
<tr>
	<td height="20" colspan="5" bgcolor=#eeeeee align=right>
		<input type=hidden name=path value=<%=r_path%>>
		<input type=hidden name=page value=<%=MyPage%>>
		<input type=hidden name=attrib value=<%=attrib%>>
		<input type=checkbox name=chkall value=on onclick="CheckAll(this.form)">选中所有显示的文件&nbsp;
		<input type=submit name=action onclick="{if(confirm('删除选定的上传文件吗？')){this.document.check.submit();return true;}return false;}" value="删除文件" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
	</td>
</tr>
</form>
<tr>
	<td height="25" colspan="5" align=center>
		<%
		Response.write "共 "& total &" 个文件，当前第 "& Mypage &"/"& Maxpages &" 页，每页 "& MyPageSize &" 个文件 <br>"
		url="Admin_UpFileManage.asp?path="& lpath &"&order="&order									
		PageNextSize=int((MyPage-1)/PageShowSize)+1
		Pagetpage=Maxpages
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
		next
		
		if Mypage+1 <=Pagetpage  then
			Next_Page = MyPage + 1
			Response.write "<a class=black href='" & Url & "&page=" & Next_Page & "' title='第" & Next_Page & "页'>下一页</A>"
		end if
		
		if MaxPages > PageShowSize*PageNextSize then
			PageNext = PageShowSize * PageNextSize + 1
			Response.write " <A class=black href='" & Url & "&page=" & Pagetpage & "' title='第"& Pagetpage &"页'>页尾</A>"
			Response.write " <a class=black href='" & Url & "&page=" & PageNext & "' title='下" & PageShowSize & "页'>下一翻页</a>"
		end if
		%>
	</td>
</tr>
<form action="Admin_UpFileManage.asp" method=post name=select>
<tr>
	<td height="25" colspan="5" bgcolor=#ffffff align=center>
		查找验证项：
		<input type=hidden name=path value=<%=r_path%>>
		<input type=hidden name=page value=<%=MyPage%>>
		<input type=hidden name=attrib value=<%=attrib%>>
		<input type=hidden name=Check_Save value="ok">
		<input type=checkbox name=Check_Attach value="1" <%if Check_Attach="1" then%>checked<%end if%>>上传附件表&nbsp;&nbsp;
		<input type=checkbox name=Check_News value="1" <%if Check_News="1" then%>checked<%end if%>>新闻内容&nbsp;&nbsp;
		<input type=checkbox name=Check_UserFace value="1" <%if Check_UserFace="1" then%>checked<%end if%>>用户头象&nbsp;&nbsp;
		<input type=submit name=Check_Sel value="设定" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
	</td>
</tr>
</form>
<tr>
	<td height="25" colspan="5" bgcolor=#ffffff align=center>
		<font color=green>系统中的文章较多时可能需耗费更多的时间，请耐心等待！</font>
	</td>
</tr>
</table>
</div>
<!--#include file="Admin_Bottom.asp"-->
		<%
		set rs=nothing
		conn.close
		set conn=nothing
		ConnUser.close
		set ConnUser=nothing
	else
		Show_Err("对不起，该功能已经被超级系统管理员关闭，您没有权限操作！")
		response.end
	end if
end if%>