<!--#include file="Conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->

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
	list=request("action")
	%>
<HTML><HEAD>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<META content="MSHTML 6.00.3790.118" name=GENERATOR>
<LINK href="site.css" type=text/css rel=stylesheet>
<title>后台</title>
<script language="JavaScript">
<!--
if (self != top) top.location.href = window.location.href
//-->
</script>
<SCRIPT language=javascript id=clientEventHandlersJS>
function line_onclick() {
    if (pic.value == 0){
        menu.style.display = 'none';
        pic.src = 'images/jt_right.gif';
        pic.title = '展开';
        pic.value = 1;
    }
    else {
        menu.style.display = '';
        pic.src = 'images/jt_left.gif';
        pic.title = '隐藏';
        pic.value = 0;
    }
}
</SCRIPT>
</HEAD>
<BODY oncontextmenu=window.event.returnValue=false style="MARGIN: 0px" scroll=no>
<TABLE height="100%" cellSpacing=0 cellPadding=0 width="100%" border=0>
<TBODY>
	<TR>
	<TD id=menu vAlign=center noWrap align=middle name="menu">
	<div align="center">
		<IFRAME id=mainmenu style="Z-INDEX: 2; VISIBILITY: inherit; WIDTH: 200px; HEIGHT: 100%" name=mainmenu src="left.asp" frameBorder=0 noResize></IFRAME>
	</TD>
	<TD width=8>
		<TABLE height="100%" cellSpacing=0 cellPadding=0 border=0>
		<TBODY>
			<TR>
			<TD language=javascript id=line title=点击展开/隐藏左列菜单 style="WIDTH: 8px; CURSOR: hand; HEIGHT: 100%" onclick="return line_onclick()" align=middle background=images/jt_br.gif>
				<IMG id=pic title=隐藏 height=50 src="images/jt_left.gif" width=6 value="0">
			</TD>
			</TR>
		</TBODY>
		</TABLE>
	</TD>
	<TD style="WIDTH: 100%">
		<IFRAME id=list style="Z-INDEX: 1; VISIBILITY: inherit; WIDTH: 100%; HEIGHT: 100%" name=list src="<%if list="" then%>right.asp<%else%><%=list%><%end if%>" frameBorder=0 scrolling=yes></IFRAME>
	</TD>
	</TR>
</TBODY>
</TABLE>
</BODY>
</HTML>
<%else%>
	<script language=javascript>
		history.back()
		alert("对不起，管理员尚未开通您的帐号，请稍候！")
	</script>
<%end if%>