<% @language="vbscript" %>
<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp"-->
<!--#include file="ChkManage.asp" -->
<!--#include file="ChkURL.asp"-->
<%
IF request.cookies(Forcast_SN)("ManageKEY")<>"super" or request.cookies(Forcast_SN)("ManagePurview")<>"99999" then
	Show_Err("对不起，您的后台管理权限不够！")
	response.end
else
	xpurl=request("xpurl")
	email=request("email")
	copyright=request("copyright")
	version=request("version")
	ver=request("ver")
	gd1=request("gd1")
	gd2=request("gd2")
	B_BG=request("B_BG")
	
	top_news=request("top_news")
	bigclassshownum=request("bigclassshownum")
	top_sp=request("top_sp")
	top_txt=request("top_txt")
	top_img=request("top_img")
	newsnum=request("newsnum")
	picnum=request("picnum")
	topuser=request("topuser")
	linkshownum=request("linkshownum")
	reviewnum=request("reviewnum")
	
	logo=request("logo")
	logourl=request("logourl")
	banner=request("banner")
	bannerurl=request("bannerurl")
	M_BG=request("M_BG")


	
	gg1=server.htmlencode(request("gg1"))
	gg1=CheckStr(gg1)
	name=server.htmlencode(request("name"))
	name=CheckStr(name)
	copyright="新闻系统"
	
	if DbType="ACCESS" then
		version="V1.1 Access版"
	end if
	
	if DbType="MSSQL" then
		version="V1.1 MSSQL版"
	end if
	
	if DbType="MYSQL" then
		version="V1.1 MYSQL版"
	end if
	
	ver="Build1"
	R_BG=request("R_BG")
	R_TOP=request("R_TOP")
	R_MAIN=request("R_MAIN")
	L_MAIN=request("L_MAIN")
	
	If name=""  or xpurl="" or email="" Then
		Show_Err("对不起,看看哪里忘记填写了?<br><br><a href='javascript:history.back()'>回去重来</a>")
		response.end
	end if
	
	if len(gg1)>=65535 then
		Show_Err("本站公告太长了，将会破坏页面的整体风格<br><br><a href='javascript:history.back()'>回去重来</a>")
		response.end
	end if
	
	if len(url)>=100 then
		Show_Err("你的主页地址太长了<br><br><a href='javascript:history.back()'>回去重来</a>")
		response.end
	end if
	
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql="update "& db_system_Table &" set name='"&name&"',xpurl='"&xpurl&"',bigclassshownum='"&bigclassshownum&"',reviewnum='"&reviewnum&"',linkshownum='"&linkshownum&"',topuser='"&topuser&"',email='"&email&"',copyright='"&copyright&"',gd1='"&gd1&"',gd2='"&gd2&"',B_BG='"&B_BG&"',top_news="&top_news&",top_sp="&top_sp&",R_BG="&R_BG&",top_txt="&top_txt&",top_img="&top_img&",gg1='"&gg1&"',R_TOP='"&R_TOP&"',R_MAIN='"&R_MAIN&"',L_MAIN='"&L_MAIN&"',newsnum='"&newsnum&"',picnum='"&picnum&"',ver='"&ver&"',version='"&version&"',logo='"&logo&"',logourl='"&logourl&"',banner='"&banner&"',bannerurl='"&bannerurl&"',M_BG='"&M_BG&"'"
	conn.execute(sql)
	conn.close
	
	set conn=nothing
	set fso=Server.CreateObject("Scripting.FileSystemObject")
	set hf=fso.CreateTextFile(Server.mappath("inc/config.asp"),true)
	hf.write "<" & "%" & vbcrlf
	hf.write "Const MaxFileSize=" & chr(34) & trim(request("MaxFileSize")) & chr(34) & "        '上传文件大小限制" & vbcrlf
	hf.write "Const SaveUpFilesPath=" & chr(34) & "UploadFile" & chr(34) & "        		'存放上传文件的目录" & vbcrlf
	hf.write "Const UpFileType=" & chr(34) & trim(request("UpFileType")) & chr(34) & "        '允许的上传文件类型" & vbcrlf
	hf.write "Const byteType=" & chr(34) & trim(request("byteType")) & chr(34) & "        '留言本屏蔽词语" & vbcrlf
	hf.write "Const basemenu=" & chr(34) & trim(request("basemenu")) & chr(34) & "        '自定义菜单" & vbcrlf
	hf.write "Const BOTTOMmenu=" & chr(34) & trim(request("BOTTOMmenu")) & chr(34) & "	'自定义底部菜单" & vbcrlf
	hf.write "%" & ">"
	hf.close
	set hf=nothing
	set fso=nothing
	response.write "<meta http-equiv=""refresh"" content="""&freetime&";url=System.asp"">"
	Show_Message("<p align=center><font color=red>已经成功添加到数据库!<br><br>"&freetime&"秒钟后返回上页!</font>")
	response.end
end if%>