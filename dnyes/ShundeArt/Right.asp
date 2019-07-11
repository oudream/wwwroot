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
<HTML><HEAD><TITLE><%=copyright%><%=version%><%=ver%> - 管理首页</TITLE>
<META http-equiv=Content-Language content=zh-cn>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<LINK href="site.css" rel=stylesheet>
<META content="MSHTML 6.00.3790.118" name=GENERATOR>
</HEAD>
<SCRIPT src="inc/menu.js" type=text/javascript></SCRIPT>
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
<body topmargin="0">
<!--#include file=Admin_Top.asp-->
<table width="100%" height="100%" border="1" cellpadding="3" cellspacing="0" bordercolor="#C0C0C0" id="AutoNumber1" style="border-collapse: collapse">
<TBODY>
<TR class="TDtop" height=25>
<TD height="25" ><div align="center"></div></TD>
</TR>
<TR height=25>
<TD width="100%" height="84" align="center" >成功进入
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
<!--#include file=Admin_Bottom.asp-->