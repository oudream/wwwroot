<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="function.asp"-->

<%
IF request.cookies(Forcast_SN)("KEY")="" THEN

else
	usernamecookie=CheckStr(request.cookies(Forcast_SN)("UserName"))
	passwdcookie=CheckStr(trim(Request.cookies(Forcast_SN)("passwd")))
	KEYcookie=CheckStr(trim(request.cookies(Forcast_SN)("KEY")))
	if usernamecookie="" or passwdcookie="" then
		response.cookies(Forcast_SN)("UserName")=""
		response.cookies(Forcast_SN)("KEY")=""
		response.cookies(Forcast_SN)("purview")=""
		response.cookies(Forcast_SN)("fullname")=""
		response.cookies(Forcast_SN)("reglevel")=""
	else
		'判断用户的合法性
		set rs=server.createobject("adodb.recordset")
		sql="select * from "& db_User_Table &" where "& db_User_Name &"='"&usernamecookie&"'"
		rs.open sql,ConnUser,1,1
		if rs.eof and rs.bof then
			response.cookies(Forcast_SN)("UserName")=""
			response.cookies(Forcast_SN)("KEY")=""
			response.cookies(Forcast_SN)("purview")=""
			response.cookies(Forcast_SN)("fullname")=""
			response.cookies(Forcast_SN)("reglevel")=""
		end if
		IF passwdcookie<>rs(db_User_Password) THEN
			response.cookies(Forcast_SN)("UserName")=""
			response.cookies(Forcast_SN)("KEY")=""
			response.cookies(Forcast_SN)("purview")=""
			response.cookies(Forcast_SN)("fullname")=""
			response.cookies(Forcast_SN)("reglevel")=""
		END IF
		'下面判断用户级别实际在有用户级别是都应该判断
		if KEYcookie<>rs("OSKEY") then
			response.cookies(Forcast_SN)("UserName")=""
			response.cookies(Forcast_SN)("KEY")=""
			response.cookies(Forcast_SN)("purview")=""
			response.cookies(Forcast_SN)("fullname")=""
			response.cookies(Forcast_SN)("reglevel")=""
		end if
		rs.close
		set rs=nothing
	END IF
END IF



'该文件需要进行调整和设置
dim typename
NewsID=Request.QueryString("NewsID")
Page=Request.QueryString("page")

if page="" then
	page=1
	elseif not IsNumeric(page) then
		Show_Err("非法参数！<br><br><a href='javascript:history.back()'>返回</a>")
		response.end
	end if
	page=int(page)
	if newsid="" then
		Show_Err("未指定参数！<br><br><a href='javascript:history.back()'>返回</a>")
		response.end
	elseif not IsNumeric(newsid) then
		Show_Err("非法参数！<br><br><a href='javascript:history.back()'>返回</a>")
		response.end
	else
		'判断该篇文章是否审核
		set rs=server.createobject("adodb.recordset")
		sql="select * from "& db_News_Table &" where NewsId="&NewsId
		rs.open sql,conn,1,1
		if rs.eof and rs.bof then
			rs.close
			set rs=nothing
			Show_Err("指定的文章不存在！<br><br><a href='javascript:history.back()'>返回</a>")
			response.end
		else
			checked=rs("checkked")
			rs.close
			set rs=nothing
		end if
		
		if checked=1 or Request.cookies(Forcast_SN)("key")="super" or Request.cookies(Forcast_SN)("key")="typemaster" or Request.cookies(Forcast_SN)("key")="bigmaster" or Request.cookies(Forcast_SN)("key")="smallmaster" then	'文章已审核并且是相关管理员
			conn.execute("update "& db_News_Table &" Set Click=click+1 where NewsID=" & NewsID )
		end if
		set rs=server.CreateObject("ADODB.RecordSet")
		if uselevel=1 then
			if Request.cookies(Forcast_SN)("key")="" then
				rs.Source="select * from "& db_News_Table &" where checkked=1 and newslevel=0 and newsid="&newsid
			end if
			if Request.cookies(Forcast_SN)("key")="selfreg" then
				if Request.cookies(Forcast_SN)("reglevel")=3 then
					rs.Source="select * from "& db_News_Table &" where checkked=1 and newslevel<=3 and newsid="&newsid
				end if
				if Request.cookies(Forcast_SN)("reglevel")=2 then
					rs.Source="select * from "& db_News_Table &" where checkked=1 and newslevel<=2 and newsid="&newsid
				end if
				if Request.cookies(Forcast_SN)("reglevel")=1 then
					rs.Source="select * from "& db_News_Table &" where checkked=1 and newslevel<=1 and newsid="&newsid
				end if
			end if
			if Request.cookies(Forcast_SN)("key")="super" or Request.cookies(Forcast_SN)("key")="typemaster" or Request.cookies(Forcast_SN)("key")="bigmaster" or Request.cookies(Forcast_SN)("key")="smallmaster" or Request.cookies(Forcast_SN)("key")="check" then
				rs.Source="select * from "& db_News_Table &" where newsid="&newsid
			else
				rs.Source="select * from "& db_News_Table &" where newsid="&newsid
			end if
			rs.Open rs.Source,conn,1,1
			bigclassid=rs("bigclassid")
			smallclassid=rs("smallclassid")
			title=htmlencode4(trim(rs("title")))
			title1=htmlencode4(trim(rs("title")))
			about=htmlencode4(trim(rs("about")))
			Author=htmlencode4(trim(rs("Author")))
			editor=htmlencode4(trim(rs("editor")))
			Original=htmlencode4(trim(rs("Original")))
			image=rs("image")
			UpdateTime=trim(rs("UpdateTime"))
			News_Content=rs("Content")
			SpecialID=rs("SpecialID")
			SpecialID2=rs("SpecialID2")
			click=rs("click")
			EnCode=trim(rs("EnCode"))
			typeid=rs("typeid")
			titletype=rs("titletype")
			titlecolor=rs("titlecolor")
			titleface=rs("titleface")
			editor=rs("editor")
			wzdj=rs("newslevel")
			backtype=rs("backtype")
			rs.Close
			set rs=nothing

			set rs=server.CreateObject("ADODB.RecordSet")
			rs.Source="select * from "& db_Type_Table &" where typeID=" & typeID
			rs.Open rs.Source,conn,1,1
			typename=rs("typename")
			rs.Close
			set rs=nothing
			set rs=server.CreateObject("ADODB.RecordSet")
			rs.Source="select * from "& db_BigClass_Table &" Where BigClassid=" & BigClassid
			rs.Open rs.Source,conn,1,1
			bigclassname=rs("bigclassname")
			rs.close
			set rs=nothing
			if smallclassid<>"" then
				set rs=server.CreateObject("ADODB.RecordSet")
				rs.Source="select * from "& db_SmallClass_Table &" Where smallClassid=" & smallClassid
				rs.Open rs.Source,conn,1,1
				smallclassname=rs("smallclassname")
				rs.close
				set rs=nothing
			end if%>
<html><head><meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=title%><%if smallclassid<>"" then%>_<%=SmallClassName%><%end if%>_<%=BigClassName%>_<%=typename%>_<%=jjgn%></title>

<%if backtype=0 then %>
<LINK href=news.css rel=stylesheet>
<% else %>
<LINK href=News_1.css rel=stylesheet>
<% end if %>

<style type="text/css">
.newstitle  {COLOR: #000000; FONT-FAMILY:"Verdana, Arial, 宋体"; FONT-SIZE: 14px;line-height:1.5}
</style>

<script language="JavaScript" type="text/JavaScript">

function validateMode()
{
  if (!bTextMode) return true;
  alert("请先取消查看HTML源代码，进入“编辑”状态，然后再使用系统编辑功能!");
  message.focus();
  return false;
}
function validateModea()
{
  if (!bTextMode) return true;
  alert("请先取消查看HTML源代码!");
  message.focus();
  return false;
}

function sign()
{if (!validateMode()) return;
message.focus();

var range =message.document.selection.createRange();
str1="<font color='red'><b>|签收|</b>文件已经阅读</font>"
   range.pasteHTML(str1);
}

function img_zoom(e, o)		//图片鼠标滚轮缩放
{
	var zoom = parseInt(o.style.zoom, 10) || 100;
	zoom += event.wheelDelta / 12;
	if (zoom > 0) o.style.zoom = zoom + '%';
		return false;
}
</script>

<script language=javascript>
function user_login()
{
	document.UserLogin.UserName.focus();
	return false;
}
</script>

<script language="JavaScript" type="text/JavaScript">
<!--
function MM_preloadImages() { //v3.0
	var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
	var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
	if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}
//-->
</script>

<script language="JavaScript" type="text/JavaScript">

var size=14;			//字体大小

function fontZoom(fontsize){	//设置字体
	size=fontsize;
	document.getElementById('fontZoom').style.fontSize=size+'px';
}

function fontMax(){	//字体放大
	size=size+2;
	document.getElementById('fontZoom').style.fontSize=size+'px';
}

function fontMin(){	//字体缩小
	size=size-2;
	if (size < 2 ){
		size = 2
	}
	document.getElementById('fontZoom').style.fontSize=size+'px';
}

</SCRIPT>

<script language="javascript">
<!--
function changedata() {
	if ( document.AddReview.none.checked ) {
		document.AddReview.Author.value = "游客";
	} 
}
function changedata1() {
	if ( document.AddReview.none1.checked ) {
		document.AddReview.email.value = "guest@feitium.net";
	} 
}

//-->
</script>

<script language=javascript>
function CheckFormAddReview()	//检测评论发表栏填写项目是否为空
{
	if(document.AddReview.Author.value=="")
	{
		alert("请输入用户名！");
		document.AddReview.Author.focus();
		return false;
	}
	if(document.AddReview.email.value == "")
	{
		alert("请输入您的EMAIL！");
		document.AddReview.email.focus();
		return false;
	}
	if(document.AddReview.content.value == "")
	{
		alert("请输入评论内容！");
		return false;
	}
}
</script>
</head>
<%if ScrollClick = "double" then%>
	<SCRIPT language=JavaScript>
	//双击滚屏代码
	var currentpos,timer;
	
	function initialize()
	{
	timer=setInterval("scrollwindow()",50);
	}
	
	function sc(){
	clearInterval(timer);
	}
	
	function scrollwindow()
	{
	currentpos=document.body.scrollTop;
	window.scroll(0,++currentpos);
	if (currentpos != document.body.scrollTop)
		sc();
	}
	document.onmousedown=sc
	document.ondblclick=initialize
	</script>
<%else%>
	<SCRIPT>
	//单击拖动屏幕方式
	var old_y=0;  //记录鼠标按下时的Ｙ轴位置
	var yn=0;  //用于记录鼠标状态
	function moveit()
	{
	if(yn==1 &&  event.button==1)  //如果鼠标左键按下就移动页面
	document.body.scrollTop=(old_y-event.clientY); //移动页面
	}
	function downit()
	{old_y=event.clientY+document.body.scrollTop; //记住鼠标按下时的Ｙ轴位置
	yn=1; //用于记住鼠标已按下
	}
	
	function upit()
	{yn=0;}  //记住鼠标已放开
	
	document.onmouseup=upit; //鼠标放开时执行upit()函数
	document.onmousedown=downit; //鼠标按下时执行downit()函数
	document.onmousemove =moveit; //鼠标移动时执行moveit()函数
	</SCRIPT>
<%end if%>
<body topmargin="0" marginheight="0">
<SCRIPT language=javascript>
<!--
  function do_color(vobject,vvar)
  { document.getElementById(vobject).style.color=vvar; }
  function do_zooms(vobject,vvar)
  { document.getElementById(vobject).style.fontsize=vvar+'px'; }
-->
</SCRIPT>
<!--#include file="top.asp"-->
<table width="760" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
	<tr valign="top"> 
		<td>
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr> 
					<td valign="top">
						<table width="102%" border="0" cellspacing="0" cellpadding="0">
							<tr bgcolor="#FFFFFF"> 
								<td height="3" bgcolor="#FFFFFF"><img src="images/kb.gif" width="9" height="3"></td>
							</tr>
							<tr bgcolor="#FFFFFF"> 
								<td width="100%" height="25" background="images/menu-guestbook.gif" >&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;栏目导航<a class=daohang href="./" >&nbsp;</a><strong><a class=daohang href="./" >网站首页</a>&gt;&gt;<a class=daohang href="Type.asp?TypeId=<%=typeid%>"><%=typename%></a>&gt;&gt;<a class=daohang href="BigClass.asp?TypeId=<%=typeid%>&BigClassid=<%=BigClassid%>" ><%=BigClassName%></a> 
<%if smallclassid<>"" then%>
									&gt;&gt;<a class=daohang href='SmallClass.asp?typeid=<%=typeid%>&BigClassID=<%=BigClassID%>&SmallClassID=<%=SmallClassID%>'><%=SmallClassName%></a> 
<%end if%>
<%title1=htmlencode4(title1)%>
									</strong>
								</td>
							</tr>
							<tr bgcolor="#FFFFFF"> 
	                					<td height="25" background="images/menu-guest-l.gif" bgcolor="#FFFFFF" align="right">
	                 						&nbsp;&nbsp;共有 <font color=red><%=click%></font> 位读者读过此文&nbsp;&nbsp;
		字体颜色：<select name=do_color_frm size=1 onchange="if(this.options[this.selectedIndex].value!='')
		{do_color('fontZoom',this.options[this.selectedIndex].value);}">
		<option>选择颜色</option>
		<option style="background-color:Black;color:Black" value=Black>黑  色</option>
		<option style="background-color:Red;color:Red" value=Red>红  色</option>
		<option style="background-color:Yellow;color:Yellow" value=Yellow>黄  色</option>
		<option style="background-color:Green;color:Green" value=Green>绿  色</option>
		<option style="background-color:Orange;color:Orange" value=Orange>橙  色</option>
		<option style="background-color:Purple;color:Purple" value=Purple>紫  色</option>
		<option style="background-color:Blue;color:Blue" value=Blue>蓝  色</option>
		<option style="background-color:Brown;color:Brown" value=Brown>褐  色</option>
		<option style="background-color:Teal;color:Teal" value=Teal>墨  绿</option>
		<option style="background-color:Navy;color:Navy" value=Navy>深  蓝</option>
		<option style="background-color:Maroon;color:Maroon" value=Maroon>赭  石</option>
		<option style="background-color:#00FFFF;color: #00FFFF" value="#00FFFF">粉 绿</option>
		<option style="background-color:#7FFFD4;color: #7FFFD4" value="#7FFFD4">淡 绿</option>
		<option style="background-color:#FFE4C4;color: #FFE4C4" value="#FFE4C4">黄 灰</option>
		<option style="background-color:#7FFF00;color: #7FFF00" value="#7FFF00">翠 绿</option>
		<option style="background-color:#D2691E;color: #D2691E" value="#D2691E">综 红</option>
		<option style="background-color:#FF7F50;color: #FF7F50" value="#FF7F50">砖 红</option>
		<option style="background-color:#6495ED;color: #6495ED" value="#6495ED">淡 蓝</option>
		<option style="background-color:#DC143C;color: #DC143C" value="#DC143C">暗 红</option>
		<option style="background-color:#FF1493;color: #FF1493" value="#FF1493">玫瑰红</option>
		<option style="background-color:#FF00FF;color: #FF00FF" value="#FF00FF">紫 红</option>
		<option style="background-color:#FFD700;color: #FFD700" value="#FFD700">桔 黄</option>
		<option style="background-color:#DAA520;color: #DAA520" value="#DAA520">军 黄</option>
		<option style="background-color:#808080;color: #808080" value="#808080">烟 灰</option>
		<option style="background-color:#778899;color: #778899" value="#778899">深 灰</option>
		<option style="background-color:#B0C4DE;color: #B0C4DE" value="#B0C4DE">灰 蓝</option>
 		</select>&nbsp;&nbsp;
									【字体：<A class=black href="javascript:fontMax()">放大</A>&nbsp;<A class=black href="javascript:fontZoom(14)">正常</A>&nbsp;<A class=black href="javascript:fontMin()">缩小</A>】&nbsp;&nbsp;&nbsp;&nbsp;
								</td>
							</tr>
	              					<tr bgcolor="#FFFFFF">
								<td height="25" align="right" background="images/menu-l-guest.gif" bgcolor="#FFFFFF"><font color="#000000"><%if ScrollClick = "double" then%>【双击鼠标左键自动滚屏】<%else%>【单击鼠标左键拖动滚屏】<%end if%>【图片上滚动鼠标滚轮变焦图片】&nbsp;&nbsp;&nbsp;&nbsp;</font></td>
							</tr>
						</table>
					</td>
				</tr>
				<tr> 
					<td valign="top" background="images/menu-guest-l.gif">&nbsp; </td>
				</tr>
			</table>
		</td>
	</tr>
	<tr valign="top"> 
		<td> 
			<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
				<tr> 
          				<td>
          					<table border="0" cellpadding="0" cellspacing="0" width="98%" align="center">
							<tr> 
								<td width="100%" align=center>
									<br>
									<font color="<%=titlecolor%>" size="+2" face="楷体_GB2312"><strong><%=title1%></strong></font>
									<HR>
									<br>
									&nbsp;&nbsp;发表日期：<%=year(updateTime)%>年<%=month(updateTime)%>月<%=day(updateTime)%>日&nbsp;&nbsp; 
	<%if Original<>"" then%>
									出处：<%=Original%> 
	<%end if%>
									&nbsp;&nbsp; 
	<%if Author<>"" then%>
									作者：<%=Author%> 
	<%end if%>
									&nbsp;&nbsp;&nbsp;&nbsp;【编辑录入：<a href="User.asp?user=<%=editor%>"><%=editor%></a>】
								</td>
							</tr>
	<%
set rs=server.CreateObject("ADODB.RecordSet")
rs.Source="select * from "& db_News_Table &" where NewsID=" & NewsID
rs.Open rs.source,conn,1,1
typeid=rs("typeid")
Title=trim(rs("Title"))
image=rs("image")

dim mode
    
set rs11=server.CreateObject("ADODB.RecordSet")
rs11.Source="select * from "& db_Type_Table &" where typeid="&typeid&" order by typeid"
rs11.Open rs11.Source,conn,1,1
mode=rs11("mode")
rs11.close
set rs11=nothing

''添加图片鼠标滚轮缩放效果
if mouse_wheel_zoom="on" then
	News_Content=replace(News_Content,"<IMG","<IMG onmousewheel='return img_zoom(event,this)' onload='javascript:if(this.width>screen.width-333)this.width=screen.width-333'",1,-1,1) 
end if
''图片上传路径还原为 config.asp 中设定的 [FileUploadPath] 值
News_Content=replace(News_Content,"="&chr(34)&"uploadfile/","="&chr(34)&FileUploadPath,1,-1,1)

arr_Content=split(News_Content,"[---分页---]")
MaxPages=ubound(arr_Content)
%>
							<tr> 
								<td width="100%"  align="center">
									<table border="0" cellspacing="0" cellpadding="0" align="center" style="TABLE-LAYOUT: fixed">
										<tr> 
											<td width="100%" align=center></td>
										</tr>
										<tr> 
											<TD class=newstitle id=fontzoom vAlign=top><br>
                        
<%if M_BG=1 and rs("picnews")=0 and Not Instr(rs("Content"),"TD")>0 then%>
												<table border=0 align="left" cellpadding=3>
													<tr> 
														<td>
															<script language=javascript src="zongg/ad.asp?i=15"></script>
														</td>
													</tr>
												</table>
<%end if%>

<%if (checked<>1) and ((Request.cookies(Forcast_SN)("key")<>"super") and (Request.cookies(Forcast_SN)("key")<>"typemaster") and (Request.cookies(Forcast_SN)("key")<>"bigmaster") and (Request.cookies(Forcast_SN)("key")<>"smallmaster")) then	'文章未审核,并且是非相关管理员
	response.write "<P><CENTER><strong><font color='0000ff' size='+2' face='隶书'>文章还未经过审核<br>请等待或者通知管理员审核才能阅览！</font></strong></CENTER>"
	response.write "<P><CENTER><IMG SRC='" &ReadNews_CopyRight_Logo& "' BORDER='0' ALT=''></CENTER>"
else	'文章已审核
	if checked<>1 then
		response.write "<P><CENTER><strong><font color='ff00ff' size='+2' face='隶书'>提醒：该文章还未通过审核</font></strong></CENTER>"
	end if
	if cINT(wzdj)<1 then
		Response.Write arr_Content(Page-1)%>
		<CENTER><IMG SRC="<%=ReadNews_CopyRight_Logo%>" BORDER="0" ALT=""></CENTER>
	<% else %>
		<%if Request.cookies(Forcast_SN)("key")="super" or Request.cookies(Forcast_SN)("key")="typemaster" or Request.cookies(Forcast_SN)("key")="bigmaster" or Request.cookies(Forcast_SN)("key")="smallmaster" or Request.cookies(Forcast_SN)("key")="check" then %>
			<%Response.Write arr_Content(Page-1)%>
			<CENTER><IMG SRC="<%=ReadNews_CopyRight_Logo%>" BORDER="0" ALT=""></CENTER>
		<% else %>
			<%if  Request.cookies(Forcast_SN)("key")="" then%>
				<br><font color="#cc0000"><b>内容简介：</b></font><br><br>
				<%=CutStr(nohtml(rs("Content")),150)%>... 
				<br><br>
				<br><font color="#cc0000"><b>友情提醒：</b></font><br><br>
				你还没注册？或者没有登录？或者权限不够？这篇文章要求是本站符合权限要求的注册用户才能阅读！<br><br>
				<%
				response.write "文章级别："
				response.write cINT(wzdj)
				response.write "级"
				%>
				<br>
				<%
				response.write "您的权限："
				response.write "未注册"
				%>
				<br><br>
				如果你还没注册，请赶紧点此<a href="#" OnClick="adduser()"  class=bottom><font color="#cc0000">注册</font></a>吧！<br>
				<br>
				如果你已经注册但还没登录，请赶紧点此<a href="#" OnClick="user_login()"  class=bottom><font color="#cc0000">登录</font></a>吧！<br>
				<CENTER><IMG SRC="<%=ReadNews_CopyRight_Logo%>" BORDER="0" ALT=""></CENTER>
			<% else %>
				<%if  Request.cookies(Forcast_SN)("key")="selfreg" then
					if cINT(Request.cookies(Forcast_SN)("reglevel"))>=cINT(wzdj) then%>
						<%Response.Write arr_Content(Page-1)%>
						<CENTER><IMG SRC="<%=ReadNews_CopyRight_Logo%>" BORDER="0" ALT=""></CENTER>
					<% else %>
						<br><font color="#cc0000"><b>内容简介：</b></font>
						<br><br>
						<%=CutStr(nohtml(rs("Content")),150)%>... 
						<br><br>
						<br><font color="#cc0000"><b>友情提醒：</b></font><br><br>
						你还没注册？或者没有登录？或者权限不够？这篇文章要求是本站符合权限要求的注册用户才能阅读！<br><br>
						<%
						response.write "文章级别："
						response.write cINT(wzdj)
						response.write "级"
						%><br>
						<%
						response.write "您的权限："
						response.write (Request.cookies(Forcast_SN)("reglevel"))
						response.write "级"
						%>
						<br><br>
						如果你还没注册，请赶紧点此<a href="#" OnClick="adduser()" class=bottom><font color="#cc0000">注册</font></a>吧！<br>
						<br>
						如果你已经注册但还没登录，请赶紧点此<a href="#" OnClick="user_login()"  class=bottom><font color="#cc0000">登录</font></a>吧！<br>
						<CENTER><IMG SRC="<%=ReadNews_CopyRight_Logo%>" BORDER="0" ALT=""></CENTER>
						<br>
					<%end if%>
				<%end if%>
			<%end if%>
		<%end if%>
	<%end if%>
<%end if%>
												<br>
												<div align="right">
<%
url="ReadNews.asp?NewsId="&newsid
if MaxPages >0 then
	Response.write "<a class=black href='"& Url &"&page=1' title='第1页'>首页</a> "
	if Page-1 > 0 then
		Prev_Page = Page - 1
		Response.write "<a class=black href='"& Url &"&page="& Prev_Page &"' title='第"& Prev_Page &"页'>上一页</a> "
	end if

	for PageCounter=0 to MaxPages
		PageLink = PageCounter+1
		if PageLink <> Page Then
			Response.write "<a class=black href='"& Url &"&page="& PageLink &"'>["& PageLink &"]</a> "
		else
			Response.Write "<font color='#FF0000'><B>["& PageLink &"]</B></font> "
		end if
		If PageLink = MaxPages+1 Then Exit for
	Next
	if page <= Maxpages then
		bdd_Page = Page + 1
		Response.write "<a class=black href='" & Url & "&page=" & bdd_Page & "' title='第" & bdd_Page & "页'>下一页</A>"
	end if
	Response.write " <A class=black href='" & Url & "&page=" & Maxpages+1 & "' title='第"& Maxpages+1 &"页'>尾页</A>"
end if
%>
												</div>
											</td>
										</tr>
									</table>
								</td>
							</tr>
							<tr> 
								<td width="100%"  height="25">
									<div align="center"> 
									<!--#include file=attach.asp -->
									</div>
								</td>
							</tr>
							<tr> 
								<td width=100% ><hr size=1></td>
							</tr>
							<tr> 
								<td height=8></td>
							</tr>
							<tr> 
								<td width=100% height=18><B>相关专题：</b> 
<%set rs4=server.CreateObject("ADODB.RecordSet")
if SpecialID<>0 then
	rs4.Source="select * from "& db_Special_Table &" where SpecialID=" & SpecialID
	rs4.Open,conn,1,3
	if not rs4.eof then%>
		<a class=class href='Special_News.asp?SpecialID=<%=SpecialID%>'><%=trim(rs4("specialname"))%></a> 
	<%end if
	rs4.Close
	set rs4=nothing

	set rs4=server.CreateObject("ADODB.RecordSet")
	if specialid2<>0 then
		rs4.Source="select * from "& db_Special_Table &" where SpecialID=" & SpecialID2
		rs4.Open,conn,1,3
		if not rs4.eof then %>
			<a class=class href='Special_News.asp?SpecialID=<%=SpecialID2%>'><%=trim(rs4("specialname"))%></a> 
		<%end if
		rs4.Close
		set rs4=nothing
	end if%>
								</td>
							</tr>
							<tr> 
								<td width=100% height=18><B>专题信息：</b></td>
							</tr>
							<tr> 
								<td height=8></td>
							</tr>
	<%
	set rs=server.CreateObject("ADODB.RecordSet")
	rs.Source="select top 5 * from "& db_News_Table &" where NewsID<>"& NewsID & " and (SpecialID=" & SpecialID & " or SpecialID2=" & SpecialID & ") and checkked=1 order by NewsID DESC"
	rs.Open,conn,1,1
	
	if rs.EOF then
		Response.Write "<tr><td width=100% >&nbsp;没有专题信息</td></tr>"
	else
		while not rs.EOF
			title=htmlencode4((rs("title")))
			Response.Write "<tr><td width=100% height=18>&nbsp;<img src=images/goxp.gif>&nbsp;<a class=middle target=_top href='ReadNews.asp?NewsID=" & rs("NewsID") & "'>" & Title & "</a><font color=#666666>(" & trim(rs("UpdateTime")) &")[<font color=#ff0000>" & rs("click") & "</font>]</font></td></tr>"
			rs.MoveNext
		wend
		Response.Write "<tr><td width=98% align=right height=18><a class=lift href='Special_News.asp?SpecialID=" & SpecialID &"'><img src='images/more.gif' border='0'></a></td></tr>"
			
	end if

	else
	Response.Write "<font color=red ><b>专题信息无</b></font>"

	
	rs.Close
	set rs=nothing
end if
'------------------------------------------------------------------------------------------------------------------------------------
Response.Write "<tr><td width=100% ><hr size=1></td></tr><tr><td height=8></td></tr>"
if about<>""  then
Response.Write "<tr><td width=100% height=18><B>相关信息:</b></td></tr><tr><td height=8></td></tr>"
	sql="select top 5 * from "& db_News_Table &" where (about like '%" & about & "%' or title like '%" & about & "%') and checkked=1 order by newsid desc"
	set rs=conn.execute(sql)

	do while not rs.eof
		title=htmlencode4(trim(rs("title")))
		%>
							<tr> 
								<td height=18> <img src=images/goxp.gif> <a class=middle target=_top href='ReadNews.asp?NewsID=<%=rs("NewsID")%>'><%=Title%></a><font color=#666666>(<%=month(trim(rs("updateTime")))%>月<%=day(trim(rs("updateTime")))%>日)[<font color=#666666><%=rs("click")%></font>]</font></td>
							</tr>
		<% rs.movenext
	loop
	Response.Write "<tr><td width=98% align=right height=18><a class=lift href='Result.asp?keyword=" & about &"&action=title'><img src='images/more.gif' border='0'></a></td></tr>"
	rs.close   
	set rs=nothing  
else
Response.Write "<tr><td width=100% height=18><B><font color=red >相关信息无</font></b></td></tr><tr><td height=8></td></tr>"

end if

set rs1=server.CreateObject("ADODB.RecordSet")
rs1.Source="select top "&reviewnum&" * from "& db_Review_Table &" where NewsID=" & NewsID & " and passed=1 order by reviewid desc"
rs1.Open rs1.Source,conn,1,1
if rs1.EOF then  NoReview=1
	Response.Write "<tr><td width=100% ><hr size=1></td></tr><tr><td height=8></td></tr>"
	Response.Write "<tr><td width=100% ><B>相关评论:</b></td></tr><tr><td height=8></td></tr>"
	%>
							<tr> 
								<td width="100%"> 
	<%
	if NoReview then 						
		Response.Write "<font color=red ><b>相关评论无</b></font>"
	end if
	%>
								</td>
							</tr>
	<%
	if not NoReview then
		while not rs1.EOF
			author=server.HTMLEncode(trim(rs1("author")))
			email=server.HTMLEncode(trim(rs1("email")))
			reviewip=rs1("reviewip")
			content=trim(rs1("content"))
            content=replace(content,"table","ｔａｂｌｅ")
			content=replace(content,"TABLE","ｔａｂｌｅ")
			''content=replace(content,"<","&lt;")
			''content=replace(content,">","&gt;")
			''content=replace(content,chr(13),"<BR>")
			ContentLen=len(Content)
			%>
							<tr> 
								<td>
									<table width="100%" border="0" cellspacing="0" cellpadding="5" style="table-layout:fixed; word-break:break-all">
										<tr bgcolor="#EFEFEF"> 
											<td width="322">发表人：<%=author%></td>
											<td width="270"> 
												<p align="right"> 
			<%if Request.cookies(Forcast_SN)("key")="super" or showip="1" then%>
												IP：<%=reviewip%> 
			<%end if%>
											</td>
										</tr>
										<tr> 
											<td colspan="2">
												<table width="100%" border="0" cellspacing="0" cellpadding="0" style="table-layout:fixed; word-break:break-all">
													<tr> 
														<td>发表人邮件：<a href='mailto:<%=email%>'><%=email%></a></td>
														<td align="right">发表时间：<%=rs1("updatetime")%></td>
													</tr>
												</table>
											</td>
										</tr>
										<tr bgcolor="#FFFFFF">
											<td colspan="2"><%
			If ContentLen=<50  then	
				DisplayContent=Content
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;" & displaycontent
										%></td>
										</tr><%
			else
									%><tr bgcolor="#FFFFFF"> 
											<td colspan="2"><%
				DisplayContent=replace(nohtml(trim(Content)),"&nbsp;","",1,-1,1) '获取表格中留言字段内容并替换格式符
				DisplayContent=replace(DisplayContent,vbcrlf,"",1,-1,1)
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;"& CutStr(displaycontent,50) &"<a href='ReadView.asp?ReviewID=" & rs1("ReviewID") &"&NewsID=" & NewsID &"' target=_blank class=class>详细内容</a>"
			end if
										%></td>
										</tr>
									</table>
								</td>
							</tr><%
			rs1.MoveNext
		wend
	end if

	rs1.Close
	set rs1=nothing	
						%><tr> 
								<td width="98%" height="28"></td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
		</td>
	</tr>

	<%if reviewable="1" then%>  
		<form method="POST" action="AddReview.asp" name=AddReview  onSubmit="return CheckFormAddReview();">
	                <%if Request.cookies(Forcast_SN)("username")<>"" then
		
				%>
				 <tr> 
					<td width="100%" background="images/menu-news.gif" align="center">
						<table border="0" cellspacing="0" width="54%" cellpadding="4">
							<input type=hidden name=NewsID value=<%=NewsID%>>
							<input type=hidden name=url value="http://<%= Request.ServerVariables("SERVER_NAME") %><%=Request.ServerVariables("url")%>?<%=Request.ServerVariables("QUERY_STRING")%>">
							<tr> 
								<td width="26%" align="left"><a name="send"><font color="#FF0000">*</font>用&nbsp;户&nbsp;名：</a></td>
								<td width="70%"><input type="text" name="Author" size="20" value="<%=Request.cookies(Forcast_SN)("UserName")%>" readonly></td>
							</tr>
							<tr> 
								<td width="26%" align="left"><font color="#FF0000">*</font>电子邮件：</td>
								<td width="70%"><input type="text" name="email" size="20" value="<%=Request.cookies(Forcast_SN)("UserEmail")%>" ></td>
							</tr>
							<tr> 
								<td width="26%" align="left" valign="top"><font color="#FF0000">*</font>评论内容：</td>
								<td width="70%" align="left">（100字以内）<% if M_MAIN=1 then %><font color=red >签收文件</font>：
								<img class=None  src="images/watermark.gif" align="absmiddle" border="0" style="cursor:hand;" alt="签收文件" lANGUAGE="javascript" onclick="sign()"></td><% end if%>
							</tr>
							<tr>
								<td width="70%" colspan="2" align="center" valign="top">
									<script language="javascript">
									var bTextMode=false;
									document.write ('<iframe src="guestbox.asp?action=new" id="message" width="350" height="100" frameborder=yes scrolling=no align=left></iframe>')
									frames.message.document.designMode = "On";
									</script>
								</td>
							</tr>
							<tr> 
								<td width="70%" colspan="2" align="center" height="50"> 
									<input type="hidden" name="editor" value="<%=editor%>"> 
									<input name="passed" type="hidden" value="<%if reviewcheck="1" then%>1<%else%>0<%end if%>"> 
									<input type="submit" value="发表评论" name="Submit" OnClick="document.AddReview.content.value = frames.message.document.body.innerHTML;">
									<input name="reset" type=reset value="重新填写"><input type="hidden" name="content" value="">
									<input type="hidden" name="title" value="评论：<%=title%>">
								</td>
							</tr>
						</table>
					</td>
				</tr>
			<%else%>
				<tr> 
					<td width="100%" background="images/menu-news.gif" align="center">
						<table border="0" cellspacing="0" width="54%" cellpadding="4">
							<input type=hidden name=NewsID value=<%=NewsID%>>
							<input type=hidden name=url value="http://<%= Request.ServerVariables("SERVER_NAME") %><%=Request.ServerVariables("url")%>?<%=Request.ServerVariables("QUERY_STRING")%>">
							<tr> 
								<td width="26%" align="right"><a name="send"><font color="#FF0000">*</font>用&nbsp;户&nbsp;名：</a></td>
								<td width="70%"> <input type="text" name="Author" size="20">&nbsp;游客：<INPUT onclick=changedata() type=checkbox value=true  name=none></td>
							</tr>
							<tr> 
								<td width="26%" align="right"><font color="#FF0000">*</font>电子邮件：</td>
								<td width="70%"> <input name="email" type="text" value="">&nbsp;游客：<INPUT onclick=changedata1() type=checkbox value=true  name=none1></td>
							</tr>
							<tr> 
								<td width="26%" align="right" valign="top"><font color="#FF0000">*</font>评论内容：</td>
								<td width="70%" align="left">（100字以内）</td>
							</tr>
							<tr>
								<td width="70%" colspan="2" align="center" valign="top">
									<script language="javascript">
									document.write ('<iframe src="guestbox.asp?action=new" id="message" width="350" height="100" frameborder=yes scrolling=no align=left></iframe>')
									frames.message.document.designMode = "On";
									</script>
								</td>
							</tr>
							<tr> 
								<td width="70%" colspan="2" align="center" height="50"> 
									<input type="hidden" name="editor" value="<%=editor%>"> 
									<input name="passed" type="hidden" value="<%if reviewcheck="1" then%>1<%else%>0<%end if%>"> 
									<input type="submit" value="发表评论" name="Submit" OnClick="document.AddReview.content.value = frames.message.document.body.innerHTML;">
									<input name="reset" type=reset value="重新填写"><input type="hidden" name="content" value="">
									<input type="hidden" name="title" value="评论：<%=title%>">
								</td>
							</tr>
						</table>
					</td>
				</tr>
			<%end if%>
		</form>
	<%end if%>
    
	<tr valign="top">
		<td height="25" background="images/menu-l-guest.gif"></td>
	</tr>
</table>

<table width="760" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF" class="bottom">
	 <tr> 
		<td background="images/menu-guest-l.gif">
			<table width="98%" border="0" cellspacing="0" cellpadding="0" align="center">
				<tr> 
					<td width="20"></td>
					<td  width="255" height="20"> 
						<a href=Review.asp?NewsID=<%=NewsID%> target="_blank" class=bottom><img src="images/icon1.gif"  border="0">发表、查看更多关于该信息的评论</a>
					</td>
					<td  width="214" height="20"> 
						<a href=send.asp?NewsID=<%=NewsID%> target=_blank  class=bottom><img src="images/mail.gif"  border="0">将本信息发给好友</a>
					</td>
					<td  width="168"><a href="javascript:window.print()" class=bottom><img src="images/printer.gif"  border="0">打印本页</a></td>
					<td  width="90"><input type="button" name="close2" value="关闭窗口"  onClick="window.close();return false;"></td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td height="19" background="images/menu-guest-t.gif"></td>
	</tr>
</table>
<!--#include file=bottom.asp -->
</body>
</html>
	<%end if%>
<%end if
conn.close
set conn=nothing
%> 