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
	MARGIN: 0px; FONT: 9pt ����)
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
	BORDER-RIGHT: white 0px solid; OVERFLOW: hidden; BORDER-LEFT: white 1px solid; BORDER-BOTTOM: white 0px solid
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
		<TABLE cellSpacing=0 cellPadding=0 width=179 align=center>
		<TBODY>
			<TR>
			<TD height=40></TD>
			</TR>
		</TBODY>
		</TABLE>
		<TABLE cellSpacing=0 cellPadding=0 width=179 align=center>
		<TBODY>
			<TR >
			
      <TD align=middle height=35>&nbsp;</TD>
			</TR>
		</TBODY>
		</TABLE>
	<TABLE cellSpacing=0 cellPadding=0 width="100%" align=left border=0>
		<TABLE cellSpacing=0 cellPadding=0 width=179 align=center>
		<TBODY>
			<TR>
			<TD height=25>
				<SPAN><A href="javascript:window.location.reload()" target=list><B>ˢ��</B></A></SPAN>&nbsp;&nbsp;<SPAN><A href="Admin_Exit.asp" target=list><B>�˳�</B></SPAN>
			</TD>
			</TR>
		</TBODY>
		</TABLE>
		<DIV style="WIDTH: 179px">
		<TABLE cellSpacing=0 cellPadding=0 width=179 align=center>
		<TBODY>
			<TR>
			<TD height=10></TD>
			</TR>
		</TBODY>
		</TABLE>
		</DIV>

<%if (request.cookies(Forcast_SN)("KEY")="super" or request.cookies(Forcast_SN)("KEY")="bigmaster" or request.cookies(Forcast_SN)("KEY")="check" or request.cookies(Forcast_SN)("key")="typemaster") then%>
		<TABLE cellSpacing=0 cellPadding=0 width=179 align=center>
		<TBODY>
			<TR>
			
            
      <TD class=menu_title id=menuTitle1 onmouseover="this.className='menu_title2';" onclick=menuChange(menu1,<%if Request.cookies(Forcast_SN)("ManageUserName")<>"" then%>180<%else%>20<%end if%>,menuTitle1); onmouseout="this.className='menu_title';"  height=25> 
        <SPAN> + ϵͳ����</SPAN> </TD>
			</TR>
			<TR>
			<TD>
				<DIV class=sec_menu id=menu1 style="DISPLAY: none; FILTER: alpha(Opacity=0); WIDTH: 179px; HEIGHT: 0px">
				<TABLE cellSpacing=0 cellPadding=0 width=179 align=center >
				<TBODY>
	<%if Request.cookies(Forcast_SN)("ManageUserName")<>"" then%>
					<TR>
					<TD height=20>
						<A href="System.asp" target=list>�� ��վ���� ��</A>
					</TD>
					</TR>
					<TR>
					<TD  height=20>
						<A href="System1.asp" target=list>�� �������� ��</A>
					</TD>
					</TR>
					<TR>
					<TD  height=20>
						<A href="CssEdit.asp" target=list>�� ģ��༭ ��</A>
					</TD>
					</TR>
					<TR>
					<TD  height=20>
						<A href="DepManage.asp" target=list>�� ���Ź��� ��</A>
					</TD>
					</TR>
					<TR>
					<TD  height=20>
						<A href="UserManage.asp" target=list>�� �û����� ��</A>
					</TD>
					</TR>
					</TD>
					</TR>
					<TR>
					<TD  height=20>
						<A href="AdminManage.asp" target=list>�� ���ܹ��� ��</A>
					</TD>
					</TR>					
					<TR>
					<TD  height=20>
						<A href="new.asp" target=list>�� ϵͳ��ʼ ��</A>
					</TD>
					</TR>
					<TR>
					<TD  height=20>
						<A href="Manage_Tool.asp" target=list>�� ������ ��</A>
					</TD>
					</TR>
					<TR>
					<TD  height=20>
						<A href="ExitManage.asp" target=list>�� �˳����� ��</A>
					</TD>
					</TR>
	<%else%>
					<TR>
					<TD  height=20>
						<A href="Manage.asp" target=list>�� �����¼ ��</A>
					</TD>
					</TR>
	<%end if%>
					</TBODY>
				</TABLE>
				</DIV>
				<DIV style="WIDTH: 179px">
				<TABLE cellSpacing=0 cellPadding=0 width=179 align=center>
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
		<TABLE cellSpacing=0 cellPadding=0 width=179 align=center>
		<TBODY>
			<TR>
			
              
        <TD class=menu_title id=menuTitle9 onmouseover="this.className='menu_title2';" onclick=menuChange(menu9,<%if Request.cookies(Forcast_SN)("ManageUserName")<>"" then%>220<%else%>20<%end if%>,menuTitle9); onmouseout="this.className='menu_title';" height=25> 
          <SPAN> + ���ӹ���</SPAN> </TD>
			</TR>
			<TR>
			<TD>
				<DIV class=sec_menu id=menu9 style="DISPLAY: none; FILTER: alpha(Opacity=0); WIDTH: 179px; HEIGHT: 0px">
				<TABLE cellSpacing=0 cellPadding=0 width=179 align=center >
				<TBODY>
	<%if Request.cookies(Forcast_SN)("ManageUserName")<>"" then%>
					<TR>
					<TD height=20>
						<A href="SpecialManage.asp" target=list>�� ר����� ��</A>
					</TD>
					</TR>
					<TR>
					<TD height=20>
						<A href="ReViewManage.asp" target=list>�� ���۹��� ��</A>
					</TD>
					</TR>	
					<TR>
					<TD height=20>
						<A href="GuestReview.asp" target=list>�� ���Թ��� ��</A>
					</TD>
					</TR>	
					<TR>
					<TD height=20>
						<A href="CheckNews1.asp" target=list>�� ������� ��</A>
					</TD>
					</TR>
					<TR>
						<TD  height=20>
							<A href="VoteManage.asp" target=list>�� ͶƱ���� ��</A>
						</TD>
					</TR>
					<TR>
					<TD  height=20>
						<A href="BoardManage.asp" target=list>�� ������� ��</A>
					</TD>
					</TR>
					<TR>
					<TD  height=20>
						<A href="LinkManage.asp" target=list>�� �������� ��</A>
					</TD>
					</TR>
					<TR>
					<TD  height=20>
						<A href="zongg/index.asp" target=list>�� ������ ��</A>
					</TD>
					</TR>
					<TR>
					<TD  height=20>
						<%if DbType = "MSSQL" then%><A href="SQLBackRest.asp" target=list>�� ���ݻָ� ��</A><%elseif DbType = "ACCESS" then%><A href="Db_compact.asp" target=list>�� ����ѹ�� ��</A><%end if%>
					</TD>
					</TR>
					<TR>
					<TD  height=20>
						<A href="AspCheck.asp" target=list>�� ����̽�� ��</A>
					</TD>
					</TR>
					<TR>
					<TD  height=20>
						<A href="Admin_UpFileManage.asp" target=list>�� �������� ��</A>
					</TD>
					</TR>
	<%else%>
					<TR>
					<TD  height=20>
						<A href="Manage.asp" target=list>�� �����¼ ��</A>
					</TD>
					</TR>
	<%end if%>
				</TBODY>
				</TABLE>
				</DIV>
				<DIV style="WIDTH: 179px">
				<TABLE cellSpacing=0 cellPadding=0 width=179 align=center>
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
		<TABLE cellSpacing=0 cellPadding=0 width=179 align=center>
		<TBODY>
			<TR>
			
              
        <TD class=menu_title id=menuTitle2 onmouseover="this.className='menu_title2';" onclick=menuChange(menu2,80,menuTitle2); onmouseout="this.className='menu_title';" height=25> 
          <SPAN> + ͼ�Ĺ���</SPAN> </TD>
			</TR>
			<TR>
			<TD>
				<DIV class=sec_menu id=menu2 style="DISPLAY: none; FILTER: alpha(Opacity=0); WIDTH: 179px; HEIGHT: 0px">
				<TABLE cellSpacing=0 cellPadding=0 width=179 align=center >
				<TBODY>
					<TR>
					<TD height=20>
						<A href="NewsAddd1.asp" target=list>�� ������� ��</A>
					</TD>
					</TR>
					<TR>
					<TD height=20>
						<A href="MyNews.asp" target=list>�� �ҵ����� ��</A>
					</TD>
					</TR>
				</TBODY>
				</TABLE>
				</DIV>
				<DIV style="WIDTH: 179px">
				<TABLE cellSpacing=0 cellPadding=0 width=179 align=center>
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
		<TABLE cellSpacing=0 cellPadding=0 width=179 align=center>
		<TBODY>
			<TR>
			
              
        <TD class=menu_title id=menuTitle4 onmouseover="this.className='menu_title2';" onclick=menuChange(menu4,60,menuTitle4); onmouseout="this.className='menu_title';" height=25> 
          <SPAN> + ��������</SPAN> </TD>
			</TR>
			<TR>
			<TD>
				<DIV class=sec_menu id=menu4 style="DISPLAY: none; FILTER: alpha(Opacity=0); WIDTH: 179px; HEIGHT: 0px">
				<TABLE cellSpacing=0 cellPadding=0 width=179 align=center >
				<TBODY>
					<TR>
					<TD  height=20>
						<A href="Edit.asp" target=list>�� �������� ��</A>
					</TD>
					</TR>
				</TBODY>
				</TABLE>
				</DIV>
			</TD>
			</TR>
		</TBODY>
		</TABLE>
		<TABLE cellSpacing=0 cellPadding=0 width=179 align=center>
		<TBODY>
			<TR>
			
              
        <TD class=menu_title id=menuTitle10 onmouseover="this.className='menu_title2';" onclick=menuChange(menu10,600,menuTitle10); onmouseout="this.className='menu_title';" height=25> 
          <SPAN> + ��Э�ڲ�����</SPAN> </TD>
			</TR>
			<TR>
			<TD>
				<DIV class=sec_menu id=menu10 style="DISPLAY: none; FILTER: alpha(Opacity=0); WIDTH: 179px; HEIGHT: 0px">
				<TABLE cellSpacing=0 cellPadding=0 width=179 heigh="100%" align=center >
				<TBODY>
                  <TR> 
                    <TD height="20">&nbsp;</TD>
                  </TR>
                  <TR> 
                    <TD><a href="art_artist_add.asp" target=list>�� ������ ��� ��</a></TD>
                  </TR>
                  <TR> 
                    <TD><a href="art_artist_view.asp" target=list> �� ������ �鿴�޸� ��</a></TD>
                  </TR>
                  <TR> 
                    <TD height="15">&nbsp;</TD>
                  </TR>
                  <TR> 
                    <TD><a href="art_subs_add.asp" target=list> �� ��Ʒ��� ��</a></TD>
                  </TR>
                  <TR> 
                    <TD><a href="art_subs_view.asp" target=list> �� ��Ʒ�鿴�޸� ��</a></TD>
                  </TR>
                  <TR> 
                    <TD height="15">&nbsp;</TD>
                  </TR>
                  <TR> 
                    <TD><a href="art_collecter_society_add.asp" target=list> �� �ղؼ�Э�� ��� ��</a></TD>
                  </TR>
                  <TR> 
                    <TD><a href="art_collecter_society_view.asp" target=list> �� �ղؼ�Э�� �鿴�޸� ��</a></TD>
                  </TR>
                  <TR> 
                    <TD height="15">&nbsp;</TD>
                  </TR>
                  <TR> 
                    <TD><a href="art_shopper_govermente_add.asp" target=list>�� �ط�������� ��� ��</a></TD>
                  </TR>
                  <TR> 
                    <TD><a href="art_shopper_govermente_view.asp" target=list>�� �ط�������� �鿴�޸� ��</a></TD>
                  </TR>
                  <TR> 
                    <TD height="15">&nbsp;</TD>
                  </TR>
                  <TR> 
                    <TD><a href="art_business_add.asp" target=list>�� ��ҵ ��� ��</a></TD>
                  </TR>
                  <TR> 
                    <TD><a href="art_business_view.asp" target=list>�� ��ҵ �鿴�޸� ��</a></TD>
                  </TR>
                  <TR> 
                    <TD height="15">&nbsp;</TD>
                  </TR>
                  <TR> 
                    <TD><a href="art_artist_society_add.asp" target=list>�� ������Э�� ��� ��</a></TD>
                  </TR>
                  <TR> 
                    <TD><a href="art_artist_society_view.asp" target=list>�� ������Э�� �鿴�޸� ��</a></TD>
                  </TR>
                  <TR> 
                    <TD height="15">&nbsp;</TD>
                  </TR>
                  <TR> 
                    <TD><a href="art_auction_add.asp" target=list>�� ����Ʒ������ ��� ��</a></TD>
                  </TR>
                  <TR> 
                    <TD><a href="art_auction_view.asp" target=list>�� ����Ʒ������ �鿴�޸� ��</a></TD>
                  </TR>
                  <TR> 
                    <TD height="15">&nbsp;</TD>
                  </TR>
                  <TR> 
                    <TD><a href="art_museum_add.asp" target=list>�� �й������ 
                      ��� ��</a></TD>
                  </TR>
                  <TR> 
                    <TD><a href="art_museum_view.asp" target=list>�� �й������ 
                      �鿴�޸� ��</a></TD>
                  </TR>
                  <TR> 
                    <TD height="15">&nbsp;</TD>
                  </TR>
                  <TR> 
                    <TD><a href="art_gallery_add.asp" target=list>�� �й����� ��� ��</a></TD>
                  </TR>
                  <TR> 
                    <TD><a href="art_gallery_view.asp" target=list>�� �й����� �鿴�޸� ��</a></TD>
                  </TR>
                  <TR> 
                    <TD height="15">&nbsp;</TD>
                  </TR>
                  <TR> 
                    <TD><a href="art_shopper_add.asp" target=list>�� ���˿ͻ���� ��� ��</a></TD>
                  </TR>
                  <TR> 
                    <TD><a href="art_shopper_view.asp" target=list>�� ���˿ͻ���� �鿴�޸� ��</a></TD>
                  </TR>
                  <TR> 
                    <TD height="15">&nbsp;</TD>
                  </TR>
                  <TR> 
                    <TD><a href="art_note_old_add.asp" target=list>�� ��Э���� ��� ��</a></TD>
                  </TR>
                  <TR> 
                    <TD><a href="art_note_old_view.asp" target=list>�� ��Э���� �鿴�޸� ��</a></TD>
                  </TR>
                  <TR> 
                    <TD height="15">&nbsp;</TD>
                  </TR>
                  <TR> 
                    <TD><a href="art_note_view.asp" target=list>�� ��˾���� ��</a></TD>
                  </TR>
                  <TR> 
                    <TD><a href="art_message_view.asp" target=list>�� ����Ա���� ��</a></TD>
                  </TR>
                  <TR> 
                    <TD height="15">&nbsp;</TD>
                  </TR>
					<TR>
					<TD  height=20>
						<A onclick="checkclick('���Ƿ�Ҫ���µ�¼��')" href="Login.asp" target=list> �� ���µ�¼ ��</A>
					</TD>
					</TR>
					<TR>
					<TD  height=20>
						<A onclick="checkclick('������˳���վ��')" href="Admin_Exit.asp" target=list>�� �˳����� ��</A>
					</TD>
					</TR>
				</TBODY>
				</TABLE>
				</DIV>
			</TD>
			</TR>
		</TBODY>
		</TABLE>
</TABLE>
</BODY>
</HTML>
<%else%>
	<script language=javascript>
		history.back()
		alert("�Բ��𣬹���Ա��δ��ͨ�����ʺţ����Ժ�")
	</script>
<%end if%>