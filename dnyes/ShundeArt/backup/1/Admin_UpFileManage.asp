<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp" -->
<!--#include file="ChkManage.asp" -->
<%
if request.cookies(Forcast_SN)("ManageKEY")<>"super" or request.cookies(Forcast_SN)("ManagePurview")<>"99999" then
	Show_Err("�Բ������ĺ�̨����Ȩ�޲�����")
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
		  '--- ������� ---
		  'int_AttribCode:���Դ������-����
		  'str_ListSeparator:�ָ�����-�ַ���
		  'Get_Attribute=""
		  'int_AttribCode1=int_AttribCode
		  '"0 ��ͨ�ļ�|128 ѹ���ļ�|64 ���ӻ��ݷ�ʽ|32 �浵�ļ�|16 Ŀ¼|8 �������������|4 ϵͳ�ļ�|2 �����ļ�|1 ֻ���ļ�", "|", -1, 1)

		  str_AttribType = Split("��ͨ�ļ�|ѹ���ļ�|���ӻ��ݷ�ʽ|�浵�ļ�|Ŀ¼|�������������|ϵͳ�ļ�|�����ļ�|ֻ���ļ�", "|", -1, 1)
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

		'FileUploadPath	'ʵ��ͼƬĿ¼
		ThisNameFile="Admin_UpFileManage.asp"   '���屾�ļ����ļ���
		urlpath="http://"&Request.ServerVariables("SERVER_NAME")
		
		dim cpath,lpath,r_path,l_path
		r_path=Request("path")
		l_path=left(Request.ServerVariables("PATH_INFO"),instrrev(Request.ServerVariables("PATH_INFO"),"/"))&FileUploadPath

		set fsoBrowse=CreateObject("Scripting.FileSystemObject")
	
		if InStr(1,r_path, l_path,1)<>1 then
			r_path=""
		end if

		'���·�����Ƿ��С�.������ֹ�ԡ�..����ʽ�����ϼ�Ŀ¼��Ϣ�� 
		if InStr(1,r_path, ".",1)>0 or InStr(1,r_path, "%2e",1)>0 or InStr(1,r_path, "%2E",1)>0 then
			r_path=""
			response.write "<script language=javascript>alert('���棡·���к��зǷ��ַ���');</script>"
		end if
		
		if r_path="" then
			lpath=l_path
		else
			if right(r_path,1)<>"/" then  '��Ŀ¼���(/)
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
<title><%=copyright%><%=version%>&nbsp;<%=ver%> - �ļ�����</title>
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
	<td height="25" colspan="5" align="center" valign="middle" >�� <B>�� �� �� �� �� �� <font color=red >/ �� �� �� �� /</font></B> �� </td>
</tr>
</table>
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#C0C0C0" width="100%" id="AutoNumber1">
<tr>
	<td width='30%' align='center' height=20 bgcolor='#eeeeee'>�ļ���</td>
	<td width='30%' align='center' bgcolor='#eeeeee'>������֤</td>
	<td width='15%' align='center' bgcolor='#eeeeee'>�ļ���С</td>
	<td width='15%' align='center' bgcolor='#eeeeee'>�ļ�����</td>
	<td width='10%' align='center' bgcolor='#eeeeee'>����</td>
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

			PageShowSize		= 10		'ÿҳ��ʾ���ٸ�ҳ����
			MyPageSize			= 15		'ÿҳ��ʾ���ٸ��ļ�
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
								'���ϴ��������������Ƿ����ϴ���¼
								if rs("filename")=x.name then
									contain=contain&"<a href='ReadNews.asp?NewsID="& rs("newsid") &"' target='_blank' title='���ϴ�����������ʾ�ø�����ID��Ϊ"& rs("newsid") &"�������б�����'>�ϴ���¼��" & rs("newsid") & "������</a><br>"
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
								'�����ŵľ��������������Ƿ�����
								cont_search = Instr(1, rs("Content"),x.name, 1)
								if cont_search>0 then
									contain=contain&"<a href='ReadNews.asp?NewsID="& rs("newsid") &"' target='_blank' title='�ø�����ID��Ϊ"& rs("newsid") &"�����������б�����'><font color='green'>�������ݣ�" & rs("newsid") & "�����µ�"& cont_search &"�ִ�</font></a><br>"
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
								'���û����������ļ��Ƿ���ͷ������
								cont_search = Instr(1, rs(db_User_Face),x.name, 1)
								if cont_search>0 then
									contain=contain&"<a href='User.asp?User="& rs(db_User_Name) &"' target='_blank' title='�ø���Ϊ�û�"& rs(db_User_Name) &"��ͷ���ļ�'><font color='green'>�û�"& rs(db_User_Name) &"��ͷ���ļ�</font></a><br>"
								end if
								rs.MoveNext
							loop
						end if
						rs.close
					end if
					
					if contain="" then
						contain="�����û���ѡ�����֤����"
					end if

					response.write "<table border='1' cellpadding='0' cellspacing='0' style='border-collapse: collapse' bordercolor='#C0C0C0' width='100%' id='AutoNumber1'>"
					response.write "<tr>"
					response.write "<td width='30%' bgcolor=#ffffff height=20 >��"&showstring&"</td>"
					response.write "<td width='30%' align='center' bgcolor=#ffffff>"&contain&"</td>"
					response.write "<td width='15%' align='right' bgcolor=#ffffff>"&x.size&"�ֽ�&nbsp;&nbsp;</td>"
					response.write "<td width='15%' align='center' bgcolor=#ffffff><a href='#' title='"&"�ļ����ͣ�"&x.type&chr(10)&"�ļ����ԣ�"&Get_Attribute(x.Attributes,"��")&chr(10)&"����(�ϴ�)��"&x.DateCreated&chr(10)&"������ʣ�"&x.DateLastAccessed&chr(10)&"����޸ģ�"&x.DateLastModified&"'>����</a></td>"
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
	<td height="20" colspan="5">��ǰ�ļ�Ŀ¼��<%=urlpath%>/<%
		Dim Path_String, Path_Array,Path_Url
		Path_Url=""
		Path_String = Split(lpath, "/", -1, 1)
		for each Path_Array in Path_String
			if Path_Array<>"" then
				Path_Url=Path_Url &"/"& Path_Array
				response.write "<a href='"& ThisNameFile &"?path="& Path_Url &"' title='����鿴 ["& urlpath & Path_Url &"] Ŀ¼'><font color=green>"& Path_Array &"/</font></a>"
			end if
		next
		%>
	</td>
</tr>
<tr>
	<td height="20" colspan="5">��Ŀ¼��<%
		for each SubF in theSubFolders%>
		<a href='<%=ThisNameFile&"?path="& lpath & SubF.name %>'  title='����鿴 [<%=lpath & SubF.name%>] Ŀ¼'>��<%= SubF.name %>��</a>
		<%next%>
	</td>
</tr>
<tr>
	<td height="20" colspan="5" bgcolor=#eeeeee align=right>
		<input type=hidden name=path value=<%=r_path%>>
		<input type=hidden name=page value=<%=MyPage%>>
		<input type=hidden name=attrib value=<%=attrib%>>
		<input type=checkbox name=chkall value=on onclick="CheckAll(this.form)">ѡ��������ʾ���ļ�&nbsp;
		<input type=submit name=action onclick="{if(confirm('ɾ��ѡ�����ϴ��ļ���')){this.document.check.submit();return true;}return false;}" value="ɾ���ļ�" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
	</td>
</tr>
</form>
<tr>
	<td height="25" colspan="5" align=center>
		<%
		Response.write "�� "& total &" ���ļ�����ǰ�� "& Mypage &"/"& Maxpages &" ҳ��ÿҳ "& MyPageSize &" ���ļ� <br>"
		url="Admin_UpFileManage.asp?path="& lpath &"&order="&order									
		PageNextSize=int((MyPage-1)/PageShowSize)+1
		Pagetpage=Maxpages
		if PageNextSize >1 then
			PagePrev=PageShowSize*(PageNextSize-1)
			Response.write "<a class=black href='" & Url & "&page=" & PagePrev & "' title='��" & PageShowSize & "ҳ'>��һ��ҳ</a> "
			Response.write "<a class=black href='" & Url & "&page=1' title='��1ҳ'>ҳ��</a> "
		end if
		if MyPage-1 > 0 then
			Prev_Page = MyPage - 1
			Response.write "<a class=black href='" & Url & "&page=" & Prev_Page & "' title='��" & Prev_Page & "ҳ'>��һҳ</a> "
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
			Response.write "<a class=black href='" & Url & "&page=" & Next_Page & "' title='��" & Next_Page & "ҳ'>��һҳ</A>"
		end if
		
		if MaxPages > PageShowSize*PageNextSize then
			PageNext = PageShowSize * PageNextSize + 1
			Response.write " <A class=black href='" & Url & "&page=" & Pagetpage & "' title='��"& Pagetpage &"ҳ'>ҳβ</A>"
			Response.write " <a class=black href='" & Url & "&page=" & PageNext & "' title='��" & PageShowSize & "ҳ'>��һ��ҳ</a>"
		end if
		%>
	</td>
</tr>
<form action="Admin_UpFileManage.asp" method=post name=select>
<tr>
	<td height="25" colspan="5" bgcolor=#ffffff align=center>
		������֤�
		<input type=hidden name=path value=<%=r_path%>>
		<input type=hidden name=page value=<%=MyPage%>>
		<input type=hidden name=attrib value=<%=attrib%>>
		<input type=hidden name=Check_Save value="ok">
		<input type=checkbox name=Check_Attach value="1" <%if Check_Attach="1" then%>checked<%end if%>>�ϴ�������&nbsp;&nbsp;
		<input type=checkbox name=Check_News value="1" <%if Check_News="1" then%>checked<%end if%>>��������&nbsp;&nbsp;
		<input type=checkbox name=Check_UserFace value="1" <%if Check_UserFace="1" then%>checked<%end if%>>�û�ͷ��&nbsp;&nbsp;
		<input type=submit name=Check_Sel value="�趨" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
	</td>
</tr>
</form>
<tr>
	<td height="25" colspan="5" bgcolor=#ffffff align=center>
		<font color=green>ϵͳ�е����½϶�ʱ������ķѸ����ʱ�䣬�����ĵȴ���</font>
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
		Show_Err("�Բ��𣬸ù����Ѿ�������ϵͳ����Ա�رգ���û��Ȩ�޲�����")
		response.end
	end if
end if%>