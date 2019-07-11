<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="inc/config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="ChkUser.asp" -->
<!--#include file="ChkManage.asp" -->
<%
IF request.cookies(Forcast_SN)("ManageKEY")<>"super" or request.cookies(Forcast_SN)("ManagePurview")<>"99999" then
	Show_Err("对不起，您的后台管理权限不够！")
	response.end
else
	response.buffer=true
	Response.Expires=0

	Set rs9 = Server.CreateObject("ADODB.Recordset")
	sql9 ="SELECT * From "& db_System_Table &" Order By id DESC"
	RS9.open sql9,Conn,1,1
	if rs9("system")="1" or request.cookies(Forcast_SN)("ManagePurview")="99999" then
		%>
<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=site.css rel=stylesheet>
<title><%=copyright%><%=version%>&nbsp;<%=ver%> - 系统设置</title>
<style type="text/css">
<!--
.style1 {color: #FF0000}
-->
</style>
<SCRIPT language="JavaScript" type="text/javascript">
// Begin color
	function color(color)
	{
	url = 'color.htm';
	window.open(url,color,"width=430,height=440,status=no,toolbar=no,menubar=no,resizable=no");
	}
// End color-->
</script>
</head>
<body topmargin="0">
<!--#include file="Admin_Top.asp"-->
<div align="center">
<table border="1" cellpadding="3" cellspacing="0" style="border-collapse: collapse" bordercolor="#C0C0C0" width="100%" id="AutoNumber1">
<form method="post" action="SystemSave.asp" name="system">
<tr class="TDtop"> 
	<td colspan="3" height=25 width="100%" > 
		<div align="center">┊ <strong>网站属性设置</strong> ┊</div>
	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="25%" align="right" ><font color=red>*</font>网站名称：</td>
	<td  colspan="2" align="left" >
		<input type="text" name="name" size="50" value="<%=rs9("name")%>" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'"> 
	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="25%" align="right" ><font color=red>*</font>网站地址：</td>
	<td  colspan="2" align="left" width="593">
		http://<input type="text" name="xpurl" size="40" value="<%=rs9("xpurl")%>" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">/ (指系统所在地址，如http://你的域名/news)
	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="25%" align="right" ><font color=red>*</font>网管信箱：</td>
	<td  colspan="2" align="left" >
		<input type="text" name="email" size="50" value="<%=rs9("email")%>" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'"> 
	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="25%" align="right" >版权说明：</td>
	<td  colspan="2" align="left" > <font color="#000000"><b>[<%=rs9("Copyright")%>]</b></font></td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="25%" align="right" >版本信息：</td>
	<td  colspan="2" align="left" > <font color="#000000"><b>[<%=rs9("version")%>&nbsp;<%=rs9("ver")%>] </b></font></td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="25%" height="18" align="right" bgcolor="#FFFFFF">TOP菜单一、二级选择：</td>
	<td colspan="2" align="left">
		<select name="B_BG" size="1" id="gd1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
		<option <%if rs9("B_BG")<>"0" then%>selected<%end if%> value="1">一级</option>
		<option <%if rs9("B_BG")="0" then%>selected<%end if%> value="0">二级</option>
		</select>
	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="25%" height="18" align="right" bgcolor="#FFFFFF">LOGO图标：</td>
	<td colspan="2" align="left">
		<input name="logo" type="text" id="logo" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'" value="<%=rs9("logo")%>" size="50">
	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="25%" height="18" align="right" bgcolor="#FFFFFF">LOGO图标URL：</td>
	<td colspan="2" align="left">
		<input name="logourl" type="text" id="logourl" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'" value="<%=rs9("logourl")%>" size="50">
	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="25%" align="right" bgcolor="#FFFFFF">LOGO类型：</td>
	<td colspan="2" align="left">
		<select name="gd1" size="1" id="gd1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
		<option <%if rs9("gd1")<>"0" then%>selected<%end if%> value="1">photo</option>
		<option <%if rs9("gd1")="0" then%>selected<%end if%> value="0">flash</option>
		</select>
	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="25%" height="18" align="right" bgcolor="#FFFFFF">Banner条：</td>
	<td colspan="2" align="left">
		<input name="banner" type="text" id="banner" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'" value="<%=rs9("banner")%>" size="50">
	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="25%" height="18" align="right" bgcolor="#FFFFFF">Banner条URL：</td>
	<td colspan="2" align="left">
		<input name="bannerurl" type="text" id="bannerurl" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'" value="<%=rs9("bannerurl")%>" size="50">
	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="25%" align="right" bgcolor="#FFFFFF">Banner条类型：</td>
	<td colspan="2" align="left">
		<select name="gd2" size="1" id="gd2" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
		<option <%if rs9("gd2")<>"0" then%>selected<%end if%> value="1">photo</option>
		<option <%if rs9("gd2")="0" then%>selected<%end if%> value="0">flash</option>
		</select>
	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="25%" align="right" >上传文件类型：</td>
	<td  colspan="2" align="left" >
		<input type="text" name="UpFileType" size="50" value="<%=UpFileType%>" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'"><font color="#FF0000">用小写“|”分开</font>
	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="25%" align="right" >上传文件大小：</td>
	<td  colspan="2" align="left" >
		<input type="text" name="MaxFileSize" size="50" value="<%=MaxFileSize%>" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'"><font color="#FF0000">K</font>
	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="25%" align="right" >留言本屏蔽词语：</td>
	<td height="17" colspan="2" align="left" > 
		<input name="byteType" type="text" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'" value="<%=byteType%>" size="50"><font color="#FF0000">用小写“|”分开</font>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td width="25%"  align="right" bgcolor="#FFFFFF">
		自定义TOP菜单：<br><br>
		<font color="#FF0000">HTML格式书写，如不支持FSO，<br>
		编辑config.asp</font><br>
	</td>
	<td colspan="2" align="left">
		<textarea name="basemenu" cols="80" rows="6" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'"><%=basemenu%></textarea><font color="#FF0000"></font>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td width="25%"  align="right" bgcolor="#FFFFFF">
		自定义BOTTOM菜单：<br><br>
		<font color="#FF0000">HTML格式书写，如不支持FSO，<br>
		编辑config.asp</font>
	</td>
	<td colspan="2" align="left">
		<textarea name="BOTTOMmenu" cols="80" rows="6" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'"><%=BOTTOMmenu%></textarea>
	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="25%" align="right" bgcolor="#FFFFFF">浮动广告：</td>
	<td colspan="2" align="left">
		<select name="R_BG" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
		<option <%if rs9("R_BG")<>"0" then%>selected<%end if%> value="1">启用</option>
		<option <%if rs9("R_BG")="0" then%>selected<%end if%> value="0">禁用</option>
		</select>
	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="25%" height="18" align="right" bgcolor="#FFFFFF">广告图片地址：</td>
	<td colspan="2" align="left">
		<input type="text" name="R_TOP" size="50" value="<%=rs9("R_TOP")%>" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="25%" align="right" bgcolor="#FFFFFF">广告链接：</td>
	<td colspan="2" align="left">
		<input type="text" name="L_MAIN" size="50" value="<%=rs9("L_MAIN")%>" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="25%" align="right" bgcolor="#FFFFFF">广告说明：</td>
	<td colspan="2" align="left">
		<input type="text" name="R_MAIN" size="50" value="<%=rs9("R_MAIN")%>" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="25%" align="right" bgcolor="#FFFFFF">新闻页面相关：</td>
	<td colspan="2" align="left">
		新闻页面广告：<select name="M_BG" size="1" id="M_BG" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
		<option <%if rs9("M_BG")<>"0" then%>selected<%end if%> value="1">启用</option>
		<option <%if rs9("M_BG")="0" then%>selected<%end if%> value="0">禁用</option>
		</select>
	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="25%" align="right" bgcolor="#FFFFFF"> <p align="right">跑马灯公告：</p></td>
	<td colspan="2" align="left" >
		<textarea name="gg1" cols="80" rows="4" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'" title=在这里输入本站的公告内容 onMouseOver="window.status='在这里输入本站的公告内容';return true;" onMouseOut="window.status='';return true;" ><%=rs9("gg1")%></textarea> 
	</td>
</tr>
<tr class="TDtop"> 
	<td colspan="3"><div align="center">显示条数设置</div></td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="25%" align="right" >文章显示条数：</td>
	<td colspan="2" align="left">
		<select size="1" name="top_news" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
		<%for i=1 to 30%>
			<option <%if rs9("top_news")=i then%>selected<%end if%> value="<%=i%>"><%=i%></option>
		<%next%>
		</select>
	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="25%" align="right" >大类页面文章显示条数：</td>
	<td  colspan="2" align="left" >
		<select size="1" name="bigclassshownum" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
		<%for i=1 to 30%>
			<option <%if rs9("bigclassshownum")=i then%>selected<%end if%> value="<%=i%>"><%=i%></option>
		<%next%>
		</select>
	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="25%" align="right" >专题显示条数：</td>
	<td  colspan="2" align="left" >
		<select size="1" name="top_sp" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
		<%for i=1 to 30%>
			<option <%if rs9("top_sp")=i then%>selected<%end if%> value="<%=i%>"><%=i%></option>
		<%next%>
		</select>
	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="25%" align="right" >热门文章显示条数：</td>
	<td  colspan="2" align="left" > <select size="1" name="top_txt" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
		<%for i=1 to 30%>
			<option <%if rs9("top_txt")=i then%>selected<%end if%> value="<%=i%>"><%=i%></option>
		<%next%>
		</select>
	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="25%" align="right" >最新图文显示条数：</td>
	<td  colspan="2" align="left" > <select size="1" name="top_img" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
		<%for i=1 to 30%>
			<option <%if rs9("top_img")=i then%>selected<%end if%> value="<%=i%>"><%=i%></option>
		<%next%>
		</select>
	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="25%" align="right" >阅读文章页面评论条数：</td>
	<td  colspan="2" align="left" >
		<select size="1" name="reviewnum" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
		<%for i=1 to 30%>
			<option <%if rs9("reviewnum")=i then%>selected<%end if%> value="<%=i%>"><%=i%></option>
		<%next%>
		</select>
	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="25%" align="right" >图片JS文章调用数：</td>
	<td  colspan="2" align="left" >
		<select size="1" name="picnum" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
		<%for i=1 to 30%>
			<option <%if rs9("picnum")=i then%>selected<%end if%> value="<%=i%>"><%=i%></option>
		<%next%>
		</select>
	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="25%" align="right" >一般JS文章调用数：</td>
	<td  colspan="2" align="left" >
		<select size="1" name="newsnum" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
		<%for i=1 to 30%>
			<option <%if rs9("newsnum")=i then%>selected<%end if%> value="<%=i%>"><%=i%></option>
		<%next%>
		</select>
	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="25%" align="right" >用户排行榜显示数：</td>
	<td  colspan="2" align="left" >
		<select size="1" name="topuser" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
		<%for i=1 to 30%>
			<option <%if rs9("topuser")=i then%>selected<%end if%> value="<%=i%>"><%=i%></option>
		<%next%>
		</select>
	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="25%" align="right" >首页LOGO链接显示数：</td>
	<td  colspan="2" align="left" >
		<select size="1" name="linkshownum" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
		<%for i=1 to 30%>
			<option <%if rs9("linkshownum")=i then%>selected<%end if%> value="<%=i%>"><%=i%></option>
		<%next%>
		</select>
	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td colspan="3"  width="611">
		<div align="center"> 
		<input type="submit" name="Submit" value="提交" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
		<input type="reset" name="Submit2" value="重设" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
		</div>
	</td>
</tr>
</form>
</table>
</div>
</body>
</html>
<!--#include file=Admin_Bottom.asp-->
		<%rs9.close
		set rs9=nothing
		conn.close
		set conn=nothing
	else
		Show_Err("对不起，该功能已经被超级系统管理员关闭，您没有权限操作！")
		response.end
	end if
end if%>