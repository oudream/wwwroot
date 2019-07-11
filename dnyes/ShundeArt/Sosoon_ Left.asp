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
	MARGIN: 0px; FONT: 9pt 宋体)
}
TABLE {
	BORDER-RIGHT: 0px; BORDER-TOP: 0px; BORDER-LEFT: 0px; BORDER-BOTTOM: 0px
}
TD {
	FONT: 12px 宋体
}
IMG {
	BORDER-RIGHT: 0px; BORDER-TOP: 0px; VERTICAL-ALIGN: bottom; BORDER-LEFT: 0px; BORDER-BOTTOM: 0px
}
A {
	FONT: 12px 宋体; COLOR: #000000; TEXT-DECORATION: none
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
				<SPAN><A href="javascript:window.location.reload()" target=list><B>刷新</B></A></SPAN>&nbsp;&nbsp;<SPAN><A href="Admin_Exit.asp" target=list><B>退出</B></SPAN>
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
        <SPAN> + 系统管理</SPAN> </TD>
			</TR>
			<TR>
			<TD>
				<DIV class=sec_menu id=menu1 style="DISPLAY: none; FILTER: alpha(Opacity=0); WIDTH: 179px; HEIGHT: 0px">
				<TABLE cellSpacing=0 cellPadding=0 width=179 align=center >
				<TBODY>
	<%if Request.cookies(Forcast_SN)("ManageUserName")<>"" then%>
					<TR>
					<TD height=20>
						<A href="System.asp" target=list>┊ 网站属性 ┊</A>
					</TD>
					</TR>
					<TR>
					<TD  height=20>
						<A href="System1.asp" target=list>┊ 功能设置 ┊</A>
					</TD>
					</TR>
					<TR>
					<TD  height=20>
						<A href="CssEdit.asp" target=list>┊ 模版编辑 ┊</A>
					</TD>
					</TR>
					<TR>
					<TD  height=20>
						<A href="DepManage.asp" target=list>┊ 部门管理 ┊</A>
					</TD>
					</TR>
					<TR>
					<TD  height=20>
						<A href="UserManage.asp" target=list>┊ 用户管理 ┊</A>
					</TD>
					</TR>
					</TD>
					</TR>
					<TR>
					<TD  height=20>
						<A href="AdminManage.asp" target=list>┊ 超管管理 ┊</A>
					</TD>
					</TR>					
					<TR>
					<TD  height=20>
						<A href="new.asp" target=list>┊ 系统初始 ┊</A>
					</TD>
					</TR>
					<TR>
					<TD  height=20>
						<A href="Manage_Tool.asp" target=list>┊ 管理工具 ┊</A>
					</TD>
					</TR>
					<TR>
					<TD  height=20>
						<A href="ExitManage.asp" target=list>┊ 退出管理 ┊</A>
					</TD>
					</TR>
	<%else%>
					<TR>
					<TD  height=20>
						<A href="Manage.asp" target=list>┊ 管理登录 ┊</A>
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
          <SPAN> + 附加管理</SPAN> </TD>
			</TR>
			<TR>
			<TD>
				<DIV class=sec_menu id=menu9 style="DISPLAY: none; FILTER: alpha(Opacity=0); WIDTH: 179px; HEIGHT: 0px">
				<TABLE cellSpacing=0 cellPadding=0 width=179 align=center >
				<TBODY>
	<%if Request.cookies(Forcast_SN)("ManageUserName")<>"" then%>
					<TR>
					<TD height=20>
						<A href="SpecialManage.asp" target=list>┊ 专题管理 ┊</A>
					</TD>
					</TR>
					<TR>
					<TD height=20>
						<A href="ReViewManage.asp" target=list>┊ 评论管理 ┊</A>
					</TD>
					</TR>	
					<TR>
					<TD height=20>
						<A href="GuestReview.asp" target=list>┊ 留言管理 ┊</A>
					</TD>
					</TR>	
					<TR>
					<TD height=20>
						<A href="CheckNews1.asp" target=list>┊ 文章审核 ┊</A>
					</TD>
					</TR>
					<TR>
						<TD  height=20>
							<A href="VoteManage.asp" target=list>┊ 投票管理 ┊</A>
						</TD>
					</TR>
					<TR>
					<TD  height=20>
						<A href="BoardManage.asp" target=list>┊ 公告管理 ┊</A>
					</TD>
					</TR>
					<TR>
					<TD  height=20>
						<A href="LinkManage.asp" target=list>┊ 友情链接 ┊</A>
					</TD>
					</TR>
					<TR>
					<TD  height=20>
						<A href="zongg/index.asp" target=list>┊ 广告管理 ┊</A>
					</TD>
					</TR>
					<TR>
					<TD  height=20>
						<%if DbType = "MSSQL" then%><A href="SQLBackRest.asp" target=list>┊ 备份恢复 ┊</A><%elseif DbType = "ACCESS" then%><A href="Db_compact.asp" target=list>┊ 备份压缩 ┊</A><%end if%>
					</TD>
					</TR>
					<TR>
					<TD  height=20>
						<A href="AspCheck.asp" target=list>┊ 阿江探针 ┊</A>
					</TD>
					</TR>
					<TR>
					<TD  height=20>
						<A href="Admin_UpFileManage.asp" target=list>┊ 附件管理 ┊</A>
					</TD>
					</TR>
	<%else%>
					<TR>
					<TD  height=20>
						<A href="Manage.asp" target=list>┊ 管理登录 ┊</A>
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
			
              
        <TD class=menu_title id=menuTitle2 onmouseover="this.className='menu_title2';" onclick=menuChange(menu2,320,menuTitle2); onmouseout="this.className='menu_title';" height=25> 
          <SPAN> + 图文管理</SPAN> </TD>
			</TR>
			<TR>
			<TD>
				<DIV class=sec_menu id=menu2 style="DISPLAY: none; FILTER: alpha(Opacity=0); WIDTH: 179px; HEIGHT: 0px">
				<TABLE cellSpacing=0 cellPadding=0 width=179 align=center >
				<TBODY>
					<TR>
					<TD height=5>
					&nbsp;&nbsp;</TD>
					</TR>
					<TR>
					<TD  height=20>
						<A title=展开/收缩菜单 onclick=showsubmenu(10) href="TypeManage.asp" target=list>┊ 栏目管理 ┊</A>
					</TD>
					</TR>
					<TR>
					<TD id=submenu10 style="DISPLAY: none" vAlign=center height=20>
<%
response.write "	<!--==========================-->"& chr(10)
response.write "	<!--      栏目菜单开始        -->"& chr(10)
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
	response.write "栏目正在建设中！"
else
	menu_pic1=""
	menu_pic2=""	
	menu_spacer_pic=""
	
	do while not rstype.eof
		SqlBigClass="select * from "& db_BigClass_Table & "  where typeid="& rstype("typeid") &" Order by bigclassorder"	'选择大组类
		Set rsBigClass= Server.CreateObject("ADODB.Recordset")
		rsBigClass.open SqlBigClass,conn,1,1
		if rsBigClass.bof and rsBigClass.eof then
			'总栏下无大组栏
			Type_Content=Type_Content & chr(10)
			Type_Content=Type_Content &"<DIV class=parent id=KB"& rstype("typeid") &"Parent>"& chr(10)
			Type_Content=Type_Content &"&nbsp;&nbsp;<IMG src="& menu_pic1 &">"& chr(10)
			Type_Content=Type_Content &"	<A href='BigClassManage.asp?TypeID="& rstype("typeid") &"' target=list>"& rstype("typeName") &"</A>"& chr(10)
			Type_Content=Type_Content &"</DIV>"& chr(10)
			response.write Type_Content
			Type_Content=""
		else
			'总栏下有大组栏
			Type_Content=Type_Content & chr(10)
			Type_Content=Type_Content &"<DIV class=parent id=KB"& rstype("typeid") &"Parent>"& chr(10)
			Type_Content=Type_Content &"&nbsp;&nbsp;<IMG title='展开/收缩菜单' style='CURSOR: hand' onclick='expandIt("& chr(34) &"KB"& rstype("typeid") &""& chr(34) &")' src="& menu_pic2 &">"& chr(10)
			Type_Content=Type_Content &"	<A href='BigClassManage.asp?TypeID="& rstype("typeid") &"' target=list>"& rstype("typeName") &"</A>"& chr(10)
			Type_Content=Type_Content &"</DIV>"& chr(10)
			Type_Content=Type_Content &"<DIV class=child id=KB"& rstype("typeid") &"Child>"& chr(10)
			response.write Type_Content
			Type_Content=""

			do while not rsBigClass.eof
				SqlSmallClass="select * from "& db_SmallClass_Table & "  where BigClassId="& rsBigClass("BigClassId") &" Order by smallclassorder"	'选择小组类
				Set rsSmallClass= Server.CreateObject("ADODB.Recordset")
				rsSmallClass.open SqlSmallClass,conn,1,1
				if rsSmallClass.bof and rsSmallClass.eof then
					'大组下无小组栏
					BigClass_Content=BigClass_Content &"	&nbsp;&nbsp;"& chr(10)
					BigClass_Content=BigClass_Content &"	<IMG src="& menu_pic1 &">"& chr(10)
					BigClass_Content=BigClass_Content &"	<A href='SmallClassManage.asp?BigClassID="& rsBigClass("BigClassId") &"' target=list>"& rsBigClass("BigClassName") &"</A>"& chr(10)
					BigClass_Content=BigClass_Content &"	<BR>"& chr(10)
					response.write BigClass_Content
					BigClass_Content=""
				else
					'大组下有小组栏
					BigClass_Content=BigClass_Content &"	&nbsp;&nbsp;"& chr(10)
					BigClass_Content=BigClass_Content &"	<IMG title=展开/收缩菜单 style='CURSOR: hand' onclick=showsubmenu("& rstype("typeid") & rsBigClass("BigClassId") &") src="& menu_pic2 &">"& chr(10)
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
response.write "	<!--      栏目菜单结束        -->"& chr(10)
response.write "	<!--==========================-->"
%>
					</TD>
					</TR>				
					<TR>
					<TD height=20>
						<A href="MyNews.asp" target=list>┊ 我的文章 ┊</A>
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
			
              
        <TD class=menu_title id=menuTitle4 onmouseover="this.className='menu_title2';" onclick=menuChange(menu4,320,menuTitle4); onmouseout="this.className='menu_title';" height=25> 
          <SPAN> + 个人事务</SPAN> </TD>
			</TR>
			<TR>
			<TD>
				<DIV class=sec_menu id=menu4 style="DISPLAY: none; FILTER: alpha(Opacity=0); WIDTH: 179px; HEIGHT: 0px">
				<TABLE cellSpacing=0 cellPadding=0 width=179 align=center >
				<TBODY>
					<TR>
					<TD  height=20>
						<A href="Edit.asp" target=list>┊ 个人资料 ┊</A>
					</TD>
					</TR>
					<TR>
					<TD  height=20>
						<A href="wnl.asp" target=list>┊ 超级年历 ┊</A>
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
			
              
        <TD class=menu_title id=menuTitle10 onmouseover="this.className='menu_title2';" onclick=menuChange(menu10,600,menuTitle10); onmouseout="this.className='menu_title';" height=25> 
          <SPAN> + 美协内部管理</SPAN> </TD>
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
                    <TD><a href="art_artist_add.asp" target=list>┊ 美术家 添加 ┊</a></TD>
                  </TR>
                  <TR> 
                    <TD><a href="art_artist_view.asp" target=list> ┊ 美术家 查看修改 ┊</a></TD>
                  </TR>
                  <TR> 
                    <TD height="15">&nbsp;</TD>
                  </TR>
                  <TR> 
                    <TD><a href="art_subs_add.asp" target=list> ┊ 作品添加 ┊</a></TD>
                  </TR>
                  <TR> 
                    <TD><a href="art_subs_view.asp" target=list> ┊ 作品查看修改 ┊</a></TD>
                  </TR>
                  <TR> 
                    <TD height="15">&nbsp;</TD>
                  </TR>
                  <TR> 
                    <TD><a href="art_collecter_society_add.asp" target=list> ┊ 收藏家协会 添加 ┊</a></TD>
                  </TR>
                  <TR> 
                    <TD><a href="art_collecter_society_view.asp" target=list> ┊ 收藏家协会 查看修改 ┊</a></TD>
                  </TR>
                  <TR> 
                    <TD height="15">&nbsp;</TD>
                  </TR>
                  <TR> 
                    <TD><a href="art_shopper_govermente_add.asp" target=list>┊ 地方政府买家 添加 ┊</a></TD>
                  </TR>
                  <TR> 
                    <TD><a href="art_shopper_govermente_view.asp" target=list>┊ 地方政府买家 查看修改 ┊</a></TD>
                  </TR>
                  <TR> 
                    <TD height="15">&nbsp;</TD>
                  </TR>
                  <TR> 
                    <TD><a href="art_business_add.asp" target=list>┊ 企业 添加 ┊</a></TD>
                  </TR>
                  <TR> 
                    <TD><a href="art_business_view.asp" target=list>┊ 企业 查看修改 ┊</a></TD>
                  </TR>
                  <TR> 
                    <TD height="15">&nbsp;</TD>
                  </TR>
                  <TR> 
                    <TD><a href="art_artist_society_add.asp" target=list>┊ 美术家协会 添加 ┊</a></TD>
                  </TR>
                  <TR> 
                    <TD><a href="art_artist_society_view.asp" target=list>┊ 美术家协会 查看修改 ┊</a></TD>
                  </TR>
                  <TR> 
                    <TD height="15">&nbsp;</TD>
                  </TR>
                  <TR> 
                    <TD><a href="art_auction_add.asp" target=list>┊ 美术品拍卖行 添加 ┊</a></TD>
                  </TR>
                  <TR> 
                    <TD><a href="art_auction_view.asp" target=list>┊ 美术品拍卖行 查看修改 ┊</a></TD>
                  </TR>
                  <TR> 
                    <TD height="15">&nbsp;</TD>
                  </TR>
                  <TR> 
                    <TD><a href="art_museum_add.asp" target=list>┊ 中国博物馆 
                      添加 ┊</a></TD>
                  </TR>
                  <TR> 
                    <TD><a href="art_museum_view.asp" target=list>┊ 中国博物馆 
                      查看修改 ┊</a></TD>
                  </TR>
                  <TR> 
                    <TD height="15">&nbsp;</TD>
                  </TR>
                  <TR> 
                    <TD><a href="art_gallery_add.asp" target=list>┊ 中国画廊 添加 ┊</a></TD>
                  </TR>
                  <TR> 
                    <TD><a href="art_gallery_view.asp" target=list>┊ 中国画廊 查看修改 ┊</a></TD>
                  </TR>
                  <TR> 
                    <TD height="15">&nbsp;</TD>
                  </TR>
                  <TR> 
                    <TD><a href="art_artist_add.asp" target=list>┊ 个人客户买家 添加 ┊</a></TD>
                  </TR>
                  <TR> 
                    <TD><a href="art_artist_view.asp" target=list>┊ 个人客户买家 查看修改 ┊</a></TD>
                  </TR>
                  <TR> 
                    <TD height="15">&nbsp;</TD>
                  </TR>
                  <TR> 
                    <TD><a href="art_note_old_add.asp" target=list>┊ 美协纪事 添加 ┊</a></TD>
                  </TR>
                  <TR> 
                    <TD><a href="art_note_old_view.asp" target=list>┊ 美协纪事 查看修改 ┊</a></TD>
                  </TR>
                  <TR> 
                    <TD height="15">&nbsp;</TD>
                  </TR>
                  <TR> 
                    <TD><a href="art_note_view.asp" target=list>┊ 公司备忘 ┊</a></TD>
                  </TR>
                  <TR> 
                    <TD><a href="art_message_view.asp" target=list>┊ 管理员留言 ┊</a></TD>
                  </TR>
                  <TR> 
                    <TD height="15">&nbsp;</TD>
                  </TR>
					<TR>
					<TD>
						<A href="help.asp" target=list>┊ 新手上路 ┊</A>
					</TD>
					</TR>
					<TR>
					<TD  height=20>
						<A onclick="checkclick('您是否要重新登录？')" href="Login.asp" target=list> ┊ 重新登录 ┊</A>
					</TD>
					</TR>
					<TR>
					<TD  height=20>
						<A onclick="checkclick('您真的退出网站？')" href="Admin_Exit.asp" target=list>┊ 退出管理 ┊</A>
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
		alert("对不起，管理员尚未开通您的帐号，请稍候！")
	</script>
<%end if%>