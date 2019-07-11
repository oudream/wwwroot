<% @language="vbscript" %>
<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp" -->
<!--#include file="ChkManage.asp" -->
<!--#include file="ChkURL.asp"-->

<%
IF request.cookies(Forcast_SN)("ManageKEY")<>"super" or request.cookies(Forcast_SN)("ManagePurview")<>"99999" then
	Show_Err("对不起，您的后台管理权限不够！")
	response.end
else
	topfont=request("topfont")
	topbg=request("topbg")
	t_bg=request("t_bg")
	usecheck=request("usecheck")
	zs=request("zs")
	zs1=request("zs1")
	zs2=request("zs2")
	reg=request("reg")
	inputmodpwd=request("inputmodpwd")
	fabiao=request("fabiao")
	fabiaomod=request("fabiaomod")
	fabiaocheck=request("fabiaocheck")
	checkmod=request("checkmod")
	checkdel=request("checkdel")
	smallgl=request("smallgl")
	smallmana=request("smallmana")
	specialgl=request("specialgl")
	system=request("system")
	css=request("css")
	init=request("init")
	delreview=request("delreview")
	shsmallgl=request("shsmallgl")
	shspecialgl=request("shspecialgl")
	shdelreview=request("shdelreview")
	reviewable=request("reviewable")
	M_MAIN=request("M_MAIN")
	modnewsable=request("modnewsable")
	checkshow=request("checkshow")
	showuser=request("showuser")
	search=request("search")
	showsearch=request("showsearch")
	showauthor=request("showauthor")
	showclub=request("showclub")
	showlink=request("showlink")
	showlinkmap=request("showlinkmap")
	showvote=request("showvote")
	showcount=request("showcount")
	linkreg=request("linkreg")
	linkpass=request("linkpass")
	linkmana=request("linkmana")
	powermana=request("powermana")
	votemana=request("votemana")
	showip=request("showip")
	showclick=request("showclick")
	showtime=request("showtime")
	showyear=request("showyear")
	showdata=request("showdata")
	showspecial=request("showspecial")
	reviewcheck=request("reviewcheck")
	reviewcheckshow=request("reviewcheckshow")
	freetime=request("freetime")
	uselevel=request("uselevel")
	
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql="update "& db_System_Table &" set zs='"&zs&"',topfont='"&topfont&"',topbg='"&topbg&"',t_bg='"&t_bg&"',zs1='"&zs1&"',inputmodpwd='"&inputmodpwd&"',showclick='"&showclick&"',showyear='"&showyear&"',showtime='"&showtime&"',showdata='"&showdata&"',showspecial='"&showspecial&"',showsearch='"&showsearch&"',reviewcheckshow='"&reviewcheckshow&"',reviewcheck='"&reviewcheck&"',showip='"&showip&"',linkmana='"&linkmana&"',powermana='"&powermana&"',votemana='"&votemana&"',search='"&search&"',linkpass='"&linkpass&"',showauthor='"&showauthor&"',checkshow='"&checkshow&"',showuser='"&showuser&"',zs2='"&zs2&"',smallgl='"&smallgl&"',specialgl='"&specialgl&"',shsmallgl='"&shsmallgl&"',shspecialgl='"&shspecialgl&"',system='"&system&"',delreview='"&delreview&"',shdelreview='"&shdelreview&"',css='"&css&"',init='"&init&"',usecheck='"&usecheck&"',reviewable='"&reviewable&"',M_MAIN='"&M_MAIN&"',modnewsable='"&modnewsable&"',reg='"&reg&"',fabiaocheck='"&fabiaocheck&"',fabiao='"&fabiao&"',fabiaomod='"&fabiaomod&"',checkmod='"&checkmod&"',showclub='"&showclub&"',showlink='"&showlink&"',showlinkmap='"&showlinkmap&"',showvote='"&showvote&"',linkreg='"&linkreg&"',showcount='"&showcount&"',freetime='"&freetime&"',uselevel='"&uselevel&"',smallmana='"&smallmana&"',checkdel='"&checkdel&"'"
	conn.execute(sql)
	conn.close
	set conn=nothing
	response.write "<meta http-equiv=""refresh"" content="""&freetime&";url=System1.asp"">"
	Show_Message("<p align=center><font color=red>已经成功添加到数据库!<br><br>"&freetime&"秒钟后返回上页!</font>")
	response.end
end if%>