<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp" -->
<!--#include file="ChkManage.asp" -->
<%
if not(request.cookies(Forcast_SN)("ManageKEY")="super" or request.cookies(Forcast_SN)("ManageKEY")="typemaster" or request.cookies(Forcast_SN)("ManageKEY")="bigmaster")  THEN
	Show_Err("�Բ������ĺ�̨����Ȩ�޲���������")
	response.end
else
	if (request.form("action")="del") and (request.form("selectdel")<>"") then
		if (request.cookies(Forcast_SN)("ManageKEY")="super") then
			conn.execute "delete from "& db_Manage_Table &" where "& db_ManageUser_ID &" in ("&request.form("selectdel")&")"
			response.Redirect "AdminManage.asp"
		else
			Show_Err("�Բ��������ܽ��й����û�������ɾ����")
			response.end
		end if
	end if
	
	dim oskey
	dim rs,tsql
	dim rst
	oskey=request("oskey")
	set rst=server.CreateObject("ADODB.RecordSet")
	if request.cookies(Forcast_SN)("ManageKEY")="super" then
		if oskey="" then
			rst.Source="select * from "& db_Manage_Table &" order by "& db_ManageUser_ID &""
		else
			rst.Source="select * from "& db_Manage_Table &" where oskey='"&oskey&"' order by "& db_ManageUser_ID &""
		end if
	else
		if oskey="" then
			rst.Source="select * from "& db_Manage_Table &" where "& db_ManageUser_Name &"='"&request.cookies(Forcast_SN)("username")&"' order by "& db_ManageUser_ID &""
		else
			rst.Source="select * from "& db_Manage_Table &" where "& db_ManageUser_Name &"='"&request.cookies(Forcast_SN)("username")&"' and oskey='"&oskey&"' order by "& db_ManageUser_ID &""
		end if
	end if

	rst.Open rst.Source,Conn,3,1	
	PageShowSize = 10		'ÿҳ��ʾ���ٸ�ҳ
	MyPageSize   = 10		'ÿҳ��ʾ�����û�

	If Not IsNumeric(Request("page")) Or IsEmpty(Request("page")) Or Request("page") <=0 Then
		MyPage=1
	Else
		MyPage=Int(Abs(Request("page")))
	End if

	If Not rst.eof then
		rst.PageSize     = MyPageSize
		MaxPages         = rst.PageCount
		rst.absolutepage = MyPage
		total            = rst.RecordCount
		%>
<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=site.css rel=stylesheet>
<title><%=copyright%><%=version%>&nbsp;<%=ver%> - ��̨�����û�����</title>
<script language=javascript>
function CheckFormUserLogin()
{
	if(document.SeahUser.UserName.value=="")
	{
		alert("����������û�����");
		document.SeahUser.UserName.focus();
		return false;
	}
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
<body topmargin="0">
<!--#include file=Admin_Top.asp-->
<p align=center>�����û��б�(<font color=red>��ɫ</font>��ʾ��ǰ�����û�)����������û����ɲ鿴������Ϣ <a href="AdminAdd1.asp">��ӹ����û�</a>  <a href="AdminManage.asp">ȫ�������û�</a><br>
<table border="1" cellpadding="3" cellspacing="0" style="border-collapse: collapse" bordercolor="#C0C0C0" width="100%" id="AutoNumber1">
  <tr align="center"  height=25>
		  <td colspan="12" height=25 class="TDtop">�� <B>��̨�����û�����</B> �� </td>
  </tr>
  <tr align="center" class="TDtop1"  height=20> 
	<td width="4%">ID��</td>
	<td width="11%">������</td>
	<td width="10%">Ȩ��</td>
	<td width="14%">����¼��ַ</td>
	<td width="15%">����¼ʱ��</td>
	<td width="5%">��¼</td>
	<td width="6%">������</td>
	<td width="6%">�޸�</td>
	<td width="6%">ɾ��</td>
	<td width="9%">��Ա�ȼ�</td>
	<td width="10%">��Ա״̬</td>
	<td width="4%" >ѡ��</td>
</tr>
		<%response.write "<form name='formdel' method='post' action='AdminManage.asp'>"
			for i=1 to rst.PageSize
				if not rst.EOF then
					if rst("Purview")<>"99999" or request.cookies(Forcast_SN)("ManagePurview")="99999" then
						response.write "<tr align='center' height=20>"& chr(10)
							response.write "<td align='center'>" &(rst(db_ManageUser_ID))& "</td>"
							response.write "<td align='center' bgcolor='#FFFFFF'>"
							response.write "<a href='AdminList.asp?id=" &(rst(db_ManageUser_ID)) &"' title=" &(rst(db_ManageUser_Password)) &">"
							if rst(db_ManageUser_Name)=request.cookies(Forcast_SN)("username") then
								response.write "<font color=red>" &(rst("Username"))& "</font>"
							else
								response.write (rst("UserName"))
							end if
							response.write "</a>"
							response.write "</td>"& chr(10)

							response.write "<td align='center' bgcolor='#FFFFFF'>"
							if rst("oskey")="super" and rst("purview")=99999 then
								response.write "<font color='red'>��������Ա</font>"
							else
								if rst("oskey")="super" and rst("purview")=1 then
									response.write "ϵͳ����Ա"
								else
									response.write "�ǹ����û�"
								end if
							end if
							response.write "</td>"& chr(10)
			
							response.write "<td align='center'>" &(rst(db_ManageUser_LastLoginIP))& "</td>"& chr(10)
			
							response.write "<td >" &(rst(db_ManageUser_LastLoginTime))& "</td>"& chr(10)
			
							response.write "<td align='center'>" &(rst(db_ManageUser_LoginTimes))& "</td>"& chr(10)
			
							response.write "<td align='center'>" &(rst("number"))& "</td>"& chr(10)
			
							response.write "<td align='center' bgcolor='#FFFFFF'>"
							response.write "<a href='AdminEdit.asp?id=" &(rst(db_ManageUser_ID))& "' title='�޸��û���" &(rst(db_ManageUser_Name))& "��������'>�޸�</a>"
							response.write "</td>"& chr(10)
							
							response.write "<td align='center' bgcolor='#FFFFFF'>"
							if trim(request.cookies(Forcast_SN)("username"))=trim(rst(db_ManageUser_Name)) then
								response.write "----"
							else
								response.write "<a href='AdminDel.asp?id=" &(rst(db_ManageUser_ID))& "&amp;name=del' title='ɾ���û���" &(rst(db_ManageUser_Name))& "��'>ɾ��</a>"
							end if
							response.write "</td>"& chr(10)

						response.write "<td>"
						if rst("oskey")="selfreg" then
							reglevel1=rst("reglevel")
							select case reglevel1
								case 1
									response.write "��ͨ<br>"
								case 2
									response.write "�߼�<br>"
								case 3
									response.write "�ؼ�<br>"
							end select
						end if
						response.write "</td>"

						response.write "<td>"
						if rst("oskey")="selfreg" then
							jingyong1=rst("jingyong")
							Select Case jingyong1
								case 0 response.write "����"
								case 1 response.write "����"
							end Select
						end if
						response.write "</td>"
						response.write "<td align='center'>"
						if trim(request.cookies(Forcast_SN)("username"))=trim(rst(db_ManageUser_Name)) then
							response.write "--"
						else
							response.write "<input type=checkbox name=selectdel id=selectdel value=" &(rst(db_ManageUser_ID))& ">"
						end if
						response.write "</td>"
						response.write "</tr>"& chr(10)
					end if
					rst.MoveNext
				end if
			next
			response.write "<tr align='right' ><td colspan=13  height=20>"%>
	<input type=hidden name=action value='del'>
	<input type=checkbox name=chkall value=on onclick="CheckAll(this.form)" style="font-size: 9pt;  color: #000000; background-color: #ffff; solid #EAEAF4">ѡ��������ʾ�����û�&nbsp;
	<input type=submit name=action1 onclick="{if(confirm('ɾ��ѡ���Ĺ����û���')){this.document.check.submit();return true;}return false;}" value="ɾ�������û�" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
			<%response.write "</td></tr>"& chr(10)
			response.write "</form>"& chr(10)
			rst.close
			set rst=nothing%>
<tr height=20>
<td colspan=12 class="TDtop1" >
<p align=center height=25>�� <%=total%> ������ǰ�� <%=Mypage%>/<%=Maxpages%> ҳ��ÿҳ <%=MyPageSize%> �� 
			<%
			if oskey="" then
				url="AdminManage.asp?" 
			else
				url="AdminManage.asp?oskey="&oskey&"&" 
			end if
			PageNextSize=int((MyPage-1)/PageShowSize)+1
			Pagetpage=int((total-1)/rst.PageSize)+1
			
			if PageNextSize >1 then
				PagePrev=PageShowSize*(PageNextSize-1)
				Response.write "<a class=black href='" & Url & "page=" & PagePrev & "' title='��" & PageShowSize & "ҳ'>��һ��ҳ</a> "
				Response.write "<a class=black href='" & Url & "page=1' title='��1ҳ'>ҳ��</a> "
			end if
			if MyPage-1 > 0 then
				Prev_Page = MyPage - 1
				Response.write "<a class=black href='" & Url & "page=" & Prev_Page & "' title='��" & Prev_Page & "ҳ'>��һҳ</a> "
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
				Response.write "<a class=black href='" & Url & "page=" & Next_Page & "' title='��" & Next_Page & "ҳ'>��һҳ</A>"
			end if
			
			if MaxPages > PageShowSize*PageNextSize then
				PageNext = PageShowSize * PageNextSize + 1
				Response.write " <A class=black href='" & Url & "page=" & Pagetpage & "' title='��"& Pagetpage &"ҳ'>ҳβ</A>"
				Response.write " <a class=black href='" & Url & "page=" & PageNext & "' title='��" & PageShowSize & "ҳ'>��һ��ҳ</a>"
			end if
			%>
</td>
</tr>
</table>
<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" bordercolor="#C0C0C0" width="50%" id="AutoNumber1">
<tr>
<td>
<font color="#FF0000">
<li><p>ע�⣺�뽫����ӵĹ���Ա���Ӧ��Ȩ�ޱ���Ϊ�������Ա���ϡ�</li>
</font>
</td>
</tr>
</table>
			<%dim SeachUserName,SeachModel,Sea_Mod
			SeachUserName=CheckStr(request.form("UserName"))
			SeachModel=CheckStr(request.form("SeachModel"))
			%>
<table border="1" cellpadding="3" cellspacing="0" style="border-collapse: collapse" bordercolor="#C0C0C0" width="100%" id="AutoNumber1">
<form method="POST" action="AdminManage.asp" name="SeahUser"  onSubmit="return CheckFormUserLogin();">
<tr align="center" height=20>
<td colspan=13>
			<%if oskey<>"" then
				Response.write "<input type=hidden name='oskey' value="& oskey &">"
			end if%>
<input type=hidden name="page" value=<%=MyPage%>>
�����û�����<input name="UserName" <%if SeachUserName<>"" then%>value=<%=SeachUserName%><%end if%> size="15" font face="����" style="font-size: 9pt; background-color:#EAEAF4">
<input type="submit" name="Submit" value="����" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'" title="���ҹ����û�">&nbsp;&nbsp;
			<%if SeachModel="" or SeachModel=1 then
				Response.write "<input type=radio name=SeachModel checked  value=1>��ȷ&nbsp;"
				Response.write "<input type=radio name=SeachModel value=0>ģ��"
				rstsql="select * from "& db_Manage_Table &" where "& db_ManageUser_Name &" like '"& SeachUserName &"' order by "& db_ManageUser_ID
				Sea_Mod="��ȷ���ң�"
			else
				Response.write "<input type=radio name=SeachModel value=1>��ȷ&nbsp;"
				Response.write "<input type=radio name=SeachModel checked value=0>ģ��"
				rstsql="select * from "& db_Manage_Table &" where "& db_ManageUser_Name &" like '%"& SeachUserName &"%' order by "& db_ManageUser_ID
				Sea_Mod="ģ�����ң�"
			end if%>
</td>
</tr>
</form>
			<%
			if SeachUserName<>"" then
				set rst1=server.CreateObject("ADODB.RecordSet")
				rst1.Open rstsql,Conn,3,1
				if not rst1.eof then
					%>
					<form method="POST" action="AdminManage.asp" name="SeahUsertest" >
					<tr align="center" class="TDtop">
						<td width="4%">ID��</td>
						<td width="11%">������</td>
						<td width="10%">Ȩ��</td>
						<td width="14%">����¼��ַ</td>
						<td width="15%">����¼ʱ��</td>
						<td width="5%">��¼</td>
						<td width="6%">������</td>
						<td width="6%">�޸�</td>
						<td width="6%">ɾ��</td>
						<td width="9%">��Ա�ȼ�</td>
						<td width="10%">��Ա״̬</td>
						<td width="4%" >ѡ��</td>
					</tr> 
					<%response.write "<form name='formdel1' method='post' action='AdminManage.asp'>"
					do while not rst1.eof
					if rst1("purview")<>"99999" or request.cookies(Forcast_SN)("ManagePurview")="99999" then
						response.write "<tr align='center' height=20>"& chr(10)
						response.write "<td align='center'>" &(rst1(db_ManageUser_ID))& "</td>"
						response.write "<td align='center' bgcolor='#FFFFFF'>"
						response.write "<a href='AdminList.asp?id=" &(rst1(db_ManageUser_ID)) &"' title=" &(rst1(db_ManageUser_Password)) &">"
						if rst1(db_ManageUser_Name)=request.cookies(Forcast_SN)("username") then
							response.write "<font color=red>" &(rst1("Username"))& "</font>"
						else
							response.write (rst1("Username"))
						end if
						response.write "</a>"
						response.write "</td>"& chr(10)
			
						response.write "<td align='center' bgcolor='#FFFFFF'>"
							if rst1("oskey")="super" and rst1("purview")="99999" then
								response.write "<font color='red'>�����û�</font>"
							else
								oskey2=rst1("oskey")
								select case oskey2
									case "super" response.write "ϵͳ�û�"
									case "check" response.write "<a href='AdminManage.asp?oskey="& oskey2 &"'>����û�</a>"
									case "typemaster" response.write "<a href='AdminManage.asp?oskey="& oskey2 &"'>�����û�</a>"
									case "bigmaster" response.write "<a href='AdminManage.asp?oskey="& oskey2 &"'>�����û�</a>"
									case "smallmaster" response.write "<a href='AdminManage.asp?oskey="& oskey2 &"'>С���û�</a>"
									case "selfreg" response.write "<a href='AdminManage.asp?oskey="& oskey2 &"'>ע���û�</a>"
								end select
							end if
						response.write "</td>"& chr(10)
						response.write "<td align='center'>" &(rst1(db_ManageUser_LastLoginIP))& "</td>"& chr(10)
						response.write "<td>" &(rst1(db_ManageUser_LastLoginTime))& "</td>"& chr(10)
						response.write "<td align='center'>" &(rst1(db_ManageUser_LoginTimes))& "</td>"& chr(10)
						response.write "<td align='center'>" &(rst1("number"))& "</td>"& chr(10)
						response.write "<td align='center' bgcolor='#FFFFFF'>"
						response.write "<a href='AdminEdit.asp?id=" &(rst1(db_ManageUser_ID))& "' title='�޸��û���" &(rst1(db_ManageUser_Name))& "��������'>�޸�</a>"
						response.write "</td>"& chr(10)
						response.write "<td align='center' bgcolor='#FFFFFF'>"
						if trim(request.cookies(Forcast_SN)("username"))=trim(rst1(db_ManageUser_Name)) then
							response.write "----"
						else
							response.write "<a href='AdminDel.asp?id=" &(rst1(db_ManageUser_ID))& "&amp;name=del' title='ɾ���û���" &(rst1(db_ManageUser_Name))& "��'>ɾ��</a>"
						end if
						response.write "</td>"& chr(10)

						response.write "<td>"
						if rst1("oskey")="selfreg" then
							reglevel2=rst1("reglevel")
							select case reglevel2
								case 1
									response.write "��ͨ<br>"
								case 2
									response.write "�߼�<br>"
								case 3
									response.write "�ؼ�<br>"
							end select
						end if
						response.write "</td>"

						response.write "<td>"
						if rst1("oskey")="selfreg" then
							jingyong2=rst1("jingyong")
							Select Case jingyong2
								case 0 response.write "����"
								case 1 response.write "����"
							end Select
						end if
						response.write "</td>"
						response.write "<td align='center'>"
						if trim(request.cookies(Forcast_SN)("username"))=trim(rst1(db_ManageUser_Name)) then
							response.write "--"
						else
							response.write "<input type=checkbox name=selectdel id=selectdel value=" &(rst1(db_ManageUser_ID))& ">"
						end if
						response.write "</td></tr>"& chr(10)
					end if
					rst1.MoveNext
					loop
					response.write "<tr align='right' ><td colspan=13  height=20>"%>
	<input type=hidden name=action value='del'>
	<input type=checkbox name=chkall value=on onclick="CheckAll(this.form)" style="font-size: 9pt;  color: #000000; background-color: #ffff; solid #EAEAF4">ѡ��������ʾ�����û�&nbsp;
	<input type=submit name=action1 onclick="{if(confirm('ɾ��ѡ���Ĺ����û���')){this.document.check.submit();return true;}return false;}" value="ɾ�������û�" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
					<%response.write "</td></tr>"& chr(10)
				else
					response.write "<tr align='center'><td colspan=13 bgcolor='ffb100' height=90><font size=4>"& Sea_Mod &"δ�ҵ������û�["& SeachUserName &"]</font></td></tr>"& chr(10)
				end if
				rst1.close
				set rst1=nothing
			end if%>
</form>
</table>
</body>
</html>
<!--#include file=Admin_Bottom.asp-->
		<%else
			Show_Err("<p>��ǰ�޹����û���</p><p>��<a href='AdminAdd1.asp'>��ӹ����û�</a></p>")
			response.end
		end if
	END IF
%>