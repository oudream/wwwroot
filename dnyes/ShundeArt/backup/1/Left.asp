<!--#include file="Conn.ASP"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp" -->

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
%>

<HTML><HEAD><TITLE></TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<META content="MSHTML 6.00.3790.118" name=GENERATOR>
<STYLE type=text/css>
BODY {
	background:url('images/leftbr.gif'); MARGIN: 0px; FONT: 9pt ����)
}
TABLE {
	BORDER-RIGHT: 0px; BORDER-TOP: 0px; BORDER-LEFT: 0px; BORDER-BOTTOM: 0px
}
TD {
	FONT: 12px ����
}
IMG {
	BORDER-RIGHT: 0px; BORDER-TOP: 0px; VERTICAL-ALIGN: bottom; BORDER-LEFT: 0px; BORDER-BOTTOM: 0px
}
A {
	FONT: 12px ����; COLOR: #000000; TEXT-DECORATION: none
}
A:hover {
	COLOR: #428eff; TEXT-DECORATION: none
}
.sec_menu {
	BORDER-RIGHT: white 0px solid; BACKGROUND:images/leftbr1.gif; OVERFLOW: hidden; BORDER-LEFT: white 1px solid; BORDER-BOTTOM: white 0px solid
}
.menu_title {
	
}
.menu_title SPAN {
	FONT-WEIGHT: bold; LEFT: 8px; COLOR: #215dc6; POSITION: relative; TOP: 2px
}
.menu_title2 {
	CURSOR: hand
}
.menu_title2 SPAN {
	FONT-WEIGHT: bold; LEFT: 8px; COLOR: #428eff; POSITION: relative; TOP: 2px
}
</STYLE>

<SCRIPT src="inc/js.js" type=text/javascript></SCRIPT>
<SCRIPT src="inc/exit.js" type=text/javascript></SCRIPT>
<SCRIPT language=JavaScript>
<!--
/*for ie and ns*/
document.onclick=function(evt){
var evt=evt?evt:(window.event)?window.event:""
var e=evt.target?evt.target:evt.srcElement
evt.cancelBubble=true
if(e.tagName=="A"&&evt.shiftKey)return false
}
//-->
</SCRIPT>

<SCRIPT language=javascript>
  var whichOpen="";
  var whichContinue='';
</SCRIPT>

</HEAD>
<BODY oncontextmenu=window.event.returnValue=false onkeydown="if(event.keyCode==78&amp;&amp;event.ctrlKey)return false;" onresize=maxWin() leftMargin=0 topMargin=0 onload=maxWin() marginwidth="0" marginheight="0">
<SCRIPT language=JavaScript1.2 src="inc/menu.js"></SCRIPT>

<TABLE cellSpacing=0 cellPadding=0 width="100%" align=left border=0>
<TBODY>
	<TR>
	<TD vAlign=top>
		<TABLE cellSpacing=0 cellPadding=0 width=135 align=center>
		<TBODY>
			<TR>
			<TD background="images/title_bg_quit.gif" height=40></TD>
			</TR>
		</TBODY>
		</TABLE>
		<TABLE cellSpacing=0 cellPadding=0 width=135 align=center>
		<TBODY>
			<TR >
			<TD align=middle height=35>&nbsp;&nbsp;&nbsp;&nbsp;����չ������ϵͳ</TD>
			</TR>
		</TBODY>
		</TABLE>
	<TABLE cellSpacing=0 cellPadding=0 width="100%" align=left border=0>
	<TBODY>
	<TR>
	<TD vAlign=top>
		<TABLE cellSpacing=0 cellPadding=0 width=135 align=center>
		<TBODY>
			<TR align=middle>
			<TD background="images/title_bg_hide1.gif" height=25>
				<SPAN><A href="javascript:window.location.reload()" target=list><B>ˢ��</B></A></SPAN>&nbsp;&nbsp;<SPAN><A href="Admin_Exit.asp" target=list><B>�˳�</B></SPAN>
			</TD>
			</TR>
		</TBODY>
		</TABLE>
		<DIV style="WIDTH: 135px">
		<TABLE cellSpacing=0 cellPadding=0 width=135 align=center>
		<TBODY>
			<TR>
			<TD height=10></TD>
			</TR>
		</TBODY>
		</TABLE>
		</DIV>

<%if (request.cookies(Forcast_SN)("KEY")="super" or request.cookies(Forcast_SN)("KEY")="bigmaster" or request.cookies(Forcast_SN)("KEY")="check" or request.cookies(Forcast_SN)("key")="typemaster") then%>
		<TABLE cellSpacing=0 cellPadding=0 width=135 align=center>
		<TBODY>
			<TR align=middle>
			<TD class=menu_title id=menuTitle1 onmouseover="this.className='menu_title2';" onclick=menuChange(menu1,<%if Request.cookies(Forcast_SN)("ManageUserName")<>"" then%>180<%else%>20<%end if%>,menuTitle1); onmouseout="this.className='menu_title';" background=images/title_bg_hide.gif height=25>
				<SPAN>ϵͳ����</SPAN>
			</TD>
			</TR>
			<TR>
			<TD>
				<DIV class=sec_menu id=menu1 style="DISPLAY: none; FILTER: alpha(Opacity=0); WIDTH: 135px; HEIGHT: 0px">
				<TABLE cellSpacing=0 cellPadding=0 width=135 align=center background=images/leftbr1.gif>
				<TBODY>
	<%if Request.cookies(Forcast_SN)("ManageUserName")<>"" then%>
					<TR>
					<TD height=20>
						&nbsp;&nbsp;&nbsp;&nbsp;<A href="System.asp" target=list>�� ��վ���� ��</A>
					</TD>
					</TR>
					<TR>
					<TD  height=20>
						&nbsp;&nbsp;&nbsp;&nbsp;<A href="System1.asp" target=list>�� �������� ��</A>
					</TD>
					</TR>
					<TR>
					<TD  height=20>
						&nbsp;&nbsp;&nbsp;&nbsp;<A href="CssEdit.asp" target=list>�� ģ��༭ ��</A>
					</TD>
					</TR>
					<TR>
					<TD  height=20>
						&nbsp;&nbsp;&nbsp;&nbsp;<A href="DepManage.asp" target=list>�� ���Ź��� ��</A>
					</TD>
					</TR>
					<TR>
					<TD  height=20>
						&nbsp;&nbsp;&nbsp;&nbsp;<A href="UserManage.asp" target=list>�� �û����� ��</A>
					</TD>
					</TR>
					</TD>
					</TR>
					<TR>
					<TD  height=20>
						&nbsp;&nbsp;&nbsp;&nbsp;<A href="AdminManage.asp" target=list>�� ���ܹ��� ��</A>
					</TD>
					</TR>					
					<TR>
					<TD  height=20>
						&nbsp;&nbsp;&nbsp;&nbsp;<A href="new.asp" target=list>�� ϵͳ��ʼ ��</A>
					</TD>
					</TR>
					<TR>
					<TD  height=20>
						&nbsp;&nbsp;&nbsp;&nbsp;<A href="Manage_Tool.asp" target=list>�� ������ ��</A>
					</TD>
					</TR>
					<TR>
					<TD  height=20>
						&nbsp;&nbsp;&nbsp;&nbsp;<A href="ExitManage.asp" target=list>�� �˳����� ��</A>
					</TD>
					</TR>
	<%else%>
					<TR>
					<TD  height=20>
						&nbsp;&nbsp;&nbsp;&nbsp;<A href="Manage.asp" target=list>�� �����¼ ��</A>
					</TD>
					</TR>
	<%end if%>
					</TBODY>
				</TABLE>
				</DIV>
				<DIV style="WIDTH: 135px">
				<TABLE cellSpacing=0 cellPadding=0 width=135 align=center>
				<TBODY>
					<TR>
					<TD height=10></TD>
					</TR>
				</TBODY>
				</TABLE>
				</DIV>
			</TD>
			</TR>
		</TBODY>
		</TABLE>
		<TABLE cellSpacing=0 cellPadding=0 width=135 align=center>
		<TBODY>
			<TR align=middle>
			<TD class=menu_title id=menuTitle9 onmouseover="this.className='menu_title2';" onclick=menuChange(menu9,<%if Request.cookies(Forcast_SN)("ManageUserName")<>"" then%>220<%else%>20<%end if%>,menuTitle9); onmouseout="this.className='menu_title';" background=images/title_bg_hide.gif height=25>
				<SPAN>���ӹ���</SPAN> 
			</TD>
			</TR>
			<TR>
			<TD>
				<DIV class=sec_menu id=menu9 style="DISPLAY: none; FILTER: alpha(Opacity=0); WIDTH: 135px; HEIGHT: 0px">
				<TABLE cellSpacing=0 cellPadding=0 width=135 align=center background=images/leftbr1.gif>
				<TBODY>
	<%if Request.cookies(Forcast_SN)("ManageUserName")<>"" then%>
					<TR>
					<TD height=20>
						&nbsp;&nbsp;&nbsp;&nbsp;<A href="SpecialManage.asp" target=list>�� ר����� ��</A>
					</TD>
					</TR>
					<TR>
					<TD height=20>
						&nbsp;&nbsp;&nbsp;&nbsp;<A href="ReViewManage.asp" target=list>�� ���۹��� ��</A>
					</TD>
					</TR>	
					<TR>
					<TD height=20>
						&nbsp;&nbsp;&nbsp;&nbsp;<A href="GuestReview.asp" target=list>�� ���Թ��� ��</A>
					</TD>
					</TR>	
					<TR>
					<TD height=20>
						&nbsp;&nbsp;&nbsp;&nbsp;<A href="CheckNews1.asp" target=list>�� ������� ��</A>
					</TD>
					</TR>
					<TR>
						<TD  height=20>
							&nbsp;&nbsp;&nbsp;&nbsp;<A href="VoteManage.asp" target=list>�� ͶƱ���� ��</A>
						</TD>
					</TR>
					<TR>
					<TD  height=20>
						&nbsp;&nbsp;&nbsp;&nbsp;<A href="BoardManage.asp" target=list>�� ������� ��</A>
					</TD>
					</TR>
					<TR>
					<TD  height=20>
						&nbsp;&nbsp;&nbsp;&nbsp;<A href="LinkManage.asp" target=list>�� �������� ��</A>
					</TD>
					</TR>
					<TR>
					<TD  height=20>
						&nbsp;&nbsp;&nbsp;&nbsp;<A href="zongg/index.asp" target=list>�� ������ ��</A>
					</TD>
					</TR>
					<TR>
					<TD  height=20>
						&nbsp;&nbsp;&nbsp;&nbsp;<%if DbType = "MSSQL" then%><A href="SQLBackRest.asp" target=list>�� ���ݻָ� ��</A><%elseif DbType = "ACCESS" then%><A href="Db_compact.asp" target=list>�� ����ѹ�� ��</A><%end if%>
					</TD>
					</TR>
					<TR>
					<TD  height=20>
						&nbsp;&nbsp;&nbsp;&nbsp;<A href="AspCheck.asp" target=list>�� ����̽�� ��</A>
					</TD>
					</TR>
					<TR>
					<TD  height=20>
						&nbsp;&nbsp;&nbsp;&nbsp;<A href="Admin_UpFileManage.asp" target=list>�� �������� ��</A>
					</TD>
					</TR>
	<%else%>
					<TR>
					<TD  height=20>
						&nbsp;&nbsp;&nbsp;&nbsp;<A href="Manage.asp" target=list>�� �����¼ ��</A>
					</TD>
					</TR>
	<%end if%>
				</TBODY>
				</TABLE>
				</DIV>
				<DIV style="WIDTH: 135px">
				<TABLE cellSpacing=0 cellPadding=0 width=135 align=center>
				<TBODY>
					<TR>
					<TD height=10></TD>
					</TR>
				</TBODY>
				</TABLE>
				</DIV>
			</TD>
			</TR>
		</TBODY>
		</TABLE>
<%end if%>
		<TABLE cellSpacing=0 cellPadding=0 width=135 align=center>
		<TBODY>
			<TR align=middle>
			<TD class=menu_title id=menuTitle2 onmouseover="this.className='menu_title2';" onclick=menuChange(menu2,320,menuTitle2); onmouseout="this.className='menu_title';" background=images/title_bg_hide.gif height=25>
				<SPAN>ͼ�Ĺ���</SPAN> 
			</TD>
			</TR>
			<TR>
			<TD>
				<DIV class=sec_menu id=menu2 style="DISPLAY: none; FILTER: alpha(Opacity=0); WIDTH: 135px; HEIGHT: 0px">
				<TABLE cellSpacing=0 cellPadding=0 width=135 align=center background=images/leftbr1.gif>
				<TBODY>
					<TR>
					<TD height=5>
					&nbsp;&nbsp;</TD>
					</TR>
					<TR>
					<TD  height=20>
						&nbsp;&nbsp;&nbsp;&nbsp;<A title=չ��/�����˵� onclick=showsubmenu(10) href="TypeManage.asp" target=list>�� ��Ŀ���� ��</A>
					</TD>
					</TR>
					<TR>
					<TD id=submenu10 style="DISPLAY: none" vAlign=center height=20>
<%
response.write "	<!--==========================-->"& chr(10)
response.write "	<!--      ��Ŀ�˵���ʼ        -->"& chr(10)
response.write "	<!--==========================-->"
sqltype="select * from "& db_Type_Table & " order by typeorder"
Set rstype= Server.CreateObject("ADODB.Recordset")
rstype.open sqltype,conn,1,1

dim Type_Content,BigClass_Content
Type_Content=""
BigClass_Content=""

typeNum=rstype.recordcount
rstype.movefirst

if rstype.bof and rstype.eof then 
	response.write "&nbsp;&nbsp;&nbsp;&nbsp;��Ŀ���ڽ����У�"
else
	menu_pic1="images/menu1.gif"
	menu_pic2="images/menu2.gif"	
	menu_spacer_pic="images/spacer.gif"
	
	do while not rstype.eof
		SqlBigClass="select * from "& db_BigClass_Table & "  where typeid="& rstype("typeid") &" Order by bigclassorder"	'ѡ�������
		Set rsBigClass= Server.CreateObject("ADODB.Recordset")
		rsBigClass.open SqlBigClass,conn,1,1
		if rsBigClass.bof and rsBigClass.eof then
			'�������޴�����
			Type_Content=Type_Content & chr(10)
			Type_Content=Type_Content &"<DIV class=parent id=KB"& rstype("typeid") &"Parent>"& chr(10)
			Type_Content=Type_Content &"&nbsp;&nbsp;<IMG src="& menu_pic1 &">"& chr(10)
			Type_Content=Type_Content &"	<A href='BigClassManage.asp?TypeID="& rstype("typeid") &"' target=list>"& rstype("typeName") &"</A>"& chr(10)
			Type_Content=Type_Content &"</DIV>"& chr(10)
			response.write Type_Content
			Type_Content=""
		else
			'�������д�����
			Type_Content=Type_Content & chr(10)
			Type_Content=Type_Content &"<DIV class=parent id=KB"& rstype("typeid") &"Parent>"& chr(10)
			Type_Content=Type_Content &"&nbsp;&nbsp;<IMG title='չ��/�����˵�' style='CURSOR: hand' onclick='expandIt("& chr(34) &"KB"& rstype("typeid") &""& chr(34) &")' src="& menu_pic2 &">"& chr(10)
			Type_Content=Type_Content &"	<A href='BigClassManage.asp?TypeID="& rstype("typeid") &"' target=list>"& rstype("typeName") &"</A>"& chr(10)
			Type_Content=Type_Content &"</DIV>"& chr(10)
			Type_Content=Type_Content &"<DIV class=child id=KB"& rstype("typeid") &"Child>"& chr(10)
			response.write Type_Content
			Type_Content=""

			do while not rsBigClass.eof
				SqlSmallClass="select * from "& db_SmallClass_Table & "  where BigClassId="& rsBigClass("BigClassId") &" Order by smallclassorder"	'ѡ��С����
				Set rsSmallClass= Server.CreateObject("ADODB.Recordset")
				rsSmallClass.open SqlSmallClass,conn,1,1
				if rsSmallClass.bof and rsSmallClass.eof then
					'��������С����
					BigClass_Content=BigClass_Content &"	&nbsp;&nbsp;"& chr(10)
					BigClass_Content=BigClass_Content &"	<IMG src="& menu_pic1 &">"& chr(10)
					BigClass_Content=BigClass_Content &"	<A href='SmallClassManage.asp?BigClassID="& rsBigClass("BigClassId") &"' target=list>"& rsBigClass("BigClassName") &"</A>"& chr(10)
					BigClass_Content=BigClass_Content &"	<BR>"& chr(10)
					response.write BigClass_Content
					BigClass_Content=""
				else
					'��������С����
					BigClass_Content=BigClass_Content &"	&nbsp;&nbsp;"& chr(10)
					BigClass_Content=BigClass_Content &"	<IMG title=չ��/�����˵� style='CURSOR: hand' onclick=showsubmenu("& rstype("typeid") & rsBigClass("BigClassId") &") src="& menu_pic2 &">"& chr(10)
					BigClass_Content=BigClass_Content &"	<A href='SmallClassManage.asp?BigClassID="& rsBigClass("BigClassId") &"' target=list>"& rsBigClass("BigClassName") &"</A>"& chr(10)
					BigClass_Content=BigClass_Content &"	<BR>"& chr(10)
					BigClass_Content=BigClass_Content &"	<DIV id=submenu"& rstype("typeid") & rsBigClass("BigClassId") &" style='DISPLAY: none'>"& chr(10)
					response.write BigClass_Content
					BigClass_Content=""
					
					rsSmallClass.movefirst
					do while not rsSmallClass.eof
						response.write "		&nbsp;&nbsp;"
						response.write "&nbsp;&nbsp;"& chr(10)
						response.write "		<IMG src="& menu_pic1 &">"& chr(10)
						response.write "		<A href='ListNews.asp?SmallClassID="& rsSmallClass("SmallClassId") &"' target=list>"& rsSmallClass("SmallClassName") &"</A>"& chr(10)
						response.write "		<BR>"& chr(10)
						rsSmallClass.movenext
					loop
					response.write "	</DIV>"& chr(10)
				end if
				rsBigClass.movenext
			loop
			response.write "</DIV>"& chr(10)
		end if
		rstype.movenext
	loop
end if

rsSmallClass.close
set rsSmallClass=nothing

rsBigClass.close
set rsBigClass=nothing

rstype.close
set rstype=nothing

conn.close
set conn=nothing

response.write "	<!--==========================-->"& chr(10)
response.write "	<!--      ��Ŀ�˵�����        -->"& chr(10)
response.write "	<!--==========================-->"
%>
					</TD>
					</TR>				
					<TR>
					<TD height=20>
						&nbsp;&nbsp;&nbsp;&nbsp;<A href="MyNews.asp" target=list>�� �ҵ����� ��</A>
					</TD>
					</TR>
				</TBODY>
				</TABLE>
				</DIV>
				<DIV style="WIDTH: 135px">
				<TABLE cellSpacing=0 cellPadding=0 width=135 align=center>
				<TBODY>
					<TR>
					<TD height=10></TD>
					</TR>
				</TBODY>
				</TABLE>
				</DIV>
			</TD>
			</TR>
		</TBODY>
		</TABLE>
		<TABLE cellSpacing=0 cellPadding=0 width=135 align=center>
		<TBODY>
			<TR align=middle>
			<TD class=menu_title id=menuTitle4 onmouseover="this.className='menu_title2';" onclick=menuChange(menu4,40,menuTitle4); onmouseout="this.className='menu_title';" background=images/title_bg_hide.gif height=25>
				<SPAN>��������</SPAN>
			</TD>
			</TR>
			<TR>
			<TD>
				<DIV class=sec_menu id=menu4 style="DISPLAY: none; FILTER: alpha(Opacity=0); WIDTH: 135px; HEIGHT: 0px">
				<TABLE cellSpacing=0 cellPadding=0 width=135 align=center background=images/leftbr1.gif>
				<TBODY>
					<TR>
					<TD  height=20>
						&nbsp;&nbsp;&nbsp;&nbsp;<A href="Edit.asp" target=list>�� �������� ��</A>
					</TD>
					</TR>
					<TR>
					<TD  height=20>
						&nbsp;&nbsp;&nbsp;&nbsp;<A href="wnl.asp" target=list>�� �������� ��</A>
					</TD>
					</TR>
				</TBODY>
				</TABLE>
				</DIV>
				<DIV style="WIDTH: 135px">
				<TABLE cellSpacing=0 cellPadding=0 width=135 align=center>
				<TBODY>
					<TR>
					<TD height=10></TD>
					</TR>
				</TBODY>
				</TABLE>
				</DIV>
			</TD>
			</TR>
		</TBODY>
		</TABLE>
		<TABLE cellSpacing=0 cellPadding=0 width=135 align=center>
		<TBODY>
			<TR align=middle>
			<TD class=menu_title id=menuTitle10 onmouseover="this.className='menu_title2';" onclick=menuChange(menu10,100,menuTitle10); onmouseout="this.className='menu_title';" background=images/title_bg_hide.gif height=25>
				<SPAN>ϵͳ��Ϣ</SPAN> 
			</TD>
			</TR>
			<TR>
			<TD>
				<DIV class=sec_menu id=menu10 style="DISPLAY: none; FILTER: alpha(Opacity=0); WIDTH: 135px; HEIGHT: 0px">
				<TABLE cellSpacing=0 cellPadding=0 width=135 heigh="100%" align=center background=images/leftbr1.gif>
				<TBODY>
					<TR>
					<TD  height=20>
						&nbsp;&nbsp;&nbsp;&nbsp;<A href="http://feitium.yeah.net" title="����չ���ٷ���վ" target=_blank>�� ���ڹٷ� ��</A>
					</TD>
					</TR>
					<TR>
					<TD  height=20>
						&nbsp;&nbsp;&nbsp;&nbsp;<A href="help.asp" target=list>�� ������· ��</A>
					</TD>
					</TR>
					<TR>
					<TD  height=20>
						&nbsp;&nbsp;&nbsp;&nbsp;<A onclick="checkclick('���Ƿ�Ҫ���µ�¼��')" href="Login.asp" target=list>�� ���µ�¼ ��</A>
					</TD>
					</TR>
					<TR>
					<TD  height=20>
						&nbsp;&nbsp;&nbsp;&nbsp;<A onclick="checkclick('������˳���վ��')" href="Admin_Exit.asp" target=list>�� �˳����� ��</A>
					</TD>
					</TR>
					<TR>
					<TD  height=20>
						&nbsp;&nbsp;&nbsp;&nbsp;<A href="javascript:about()">�� ����ϵͳ ��</A>
					</TD>
					</TR>
				</TBODY>
				</TABLE>
				</DIV>
			</TD>
			</TR>
		</TBODY>
		</TABLE>
	</TD>
	</TR>
</TBODY>
</TABLE>
</BODY>
</HTML>
<%else%>
	<script language=javascript>
		history.back()
		alert("�Բ��𣬹���Ա��δ��ͨ�����ʺţ����Ժ�")
	</script>
<%end if%>