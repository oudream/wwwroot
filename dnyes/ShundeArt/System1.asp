<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
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
	if rs9("powermana")="1" or request.cookies(Forcast_SN)("ManagePurview")="99999" then
	%>
<html>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=site.css rel=stylesheet>
<title><%=copyright%><%=version%>&nbsp;<%=ver%> - 权限设置</title>
</head>
<body topmargin="0">
<!--#include file=Admin_Top.asp-->
<div align="center">
<table border="1" cellpadding="3" cellspacing="0" style="border-collapse: collapse" bordercolor="#C0C0C0" width="100%" id="AutoNumber1">
<form method="post" action="SystemSave1.asp" name="system">
<tr class="TDtop"> 
          <td colspan="3"  width="100%" height=25> 
            <div align="center">┊ <strong>网站属性设置</strong> ┊</div></td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%"  align="right">启用页面信息版权保护：</td>
          <td  colspan="2" align="left"><select name="topbg" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("topbg")<>"0" then%>selected<%end if%> value="1">启用</option>
              <option <%if rs9("topbg")="0" then%>selected<%end if%> value="0">禁用</option>
            </select></td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >启用栏目导航：</td>
          <td  colspan="2" align="left" > <select name="t_bg" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("t_bg")<>"0" then%>selected<%end if%> value="1">启用</option>
              <option <%if rs9("t_bg")="0" then%>selected<%end if%> value="0">禁用</option>
            </select></td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >使用分级浏览机制：</td>
          <td  colspan="2" align="left" > <select name="uselevel" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("uselevel")="1" then%>selected<%end if%> value="1">启用</option>
              <option <%if rs9("uselevel")="0" then%>selected<%end if%> value="0">禁用</option>
            </select></td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >文章默认状态：</td>
          <td  colspan="2" align="left" > <select name="usecheck" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("usecheck")="1" then%>selected<%end if%> value="1">已审核</option>
              <option <%if rs9("usecheck")="0" then%>selected<%end if%> value="0">未审核</option>
            </select>
            点击“<a title=文章默认状态详细说明 onMouseOver="window.status='文章默认状态详细说明';return true;" onMouseOut="window.status='';return true;"  href=javascript:alert("已审核指小类管理员添加文章后，文章不\n必经过系统管理员或文章审核员的审核就\n可显示出来。而未审核指小类管理员添加\n文章后，文章不直接显示出来，必须经过\n系统管理员和文章审核员的审核后才会显\n示。系统管理员和文章审核员不受此限。")>这里</a>”查看详细说明。 
          </td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >录入员可修改已审核文章：</td>
          <td  colspan="2" align="left" > <select name="modnewsable" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("modnewsable")="1" then%>selected<%end if%> value="1">开启</option>
              <option <%if rs9("modnewsable")="0" then%>selected<%end if%> value="0">关闭</option>
            </select>
            点击“<a title=录入员可修改已审核文章详细说明 onMouseOver="window.status='录入员可修改已审核文章详细说明';return true;" onMouseOut="window.status='';return true;" href=javascript:alert("当文章默认状态为未审核时，该选项开启\n表示可以修改，但修改的文章将改为未审\n核状态（出于安全的考虑）；关闭表示不\n可修改；当文章默认状态为已审核时，此\n选项无效。")>这里</a>”查看详细说明。</td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >大类注释状态：</td>
          <td  colspan="2" align="left" > <select name="zs" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("zs")="1" then%>selected<%end if%> value="1">启用</option>
              <option <%if rs9("zs")="0" then%>selected<%end if%> value="0">未用</option>
            </select>
            点击“<a title=大类注释状态详细说明 onMouseOver="window.status='大类注释状态详细说明';return true;" onMouseOut="window.status='';return true;"  href=javascript:alert("大类注释启用后，将会在查看大类文章时\n出现大类的提示信息，未启用则不出现。\n当然，前提是该大类要有提示信息。")>这里</a>”查看详细说明。 
          </td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >小类注释状态：</td>
          <td  colspan="2" align="left" > <select name="zs1" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("zs1")="1" then%>selected<%end if%> value="1">启用</option>
              <option <%if rs9("zs1")="0" then%>selected<%end if%> value="0">未用</option>
            </select>
            点击“<a title=小类注释状态详细说明 onMouseOver="window.status='小类注释状态详细说明';return true;" onMouseOut="window.status='';return true;" href=javascript:alert("小类注释启用后，将会在查看小类文章时\n出现小类的提示信息，未启用则不出现。\n当然，前提是该小类要有提示信息。")>这里</a>”查看详细说明。 
          </td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >专题注释状态：</td>
          <td  colspan="2" align="left" > <select name="zs2" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("zs2")="1" then%>selected<%end if%> value="1">启用</option>
              <option <%if rs9("zs2")="0" then%>selected<%end if%> value="0">未用</option>
            </select>
            点击“<a title=专题注释状态详细说明 onMouseOver="window.status='专题注释状态详细说明';return true;" onMouseOut="window.status='';return true;" href=javascript:alert("专题注释启用后，将会在查看专题文章时\n出现专题的提示信息，未启用则不出现。\n当然，前提是该专题要有提示信息。")>这里</a>”查看详细说明。 
          </td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >是否开启评论功能：</td>
          <td  colspan="2" align="left" > <select name="reviewable" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("reviewable")="1" then%>selected<%end if%> value="1">开启</option>
              <option <%if rs9("reviewable")="0" then%>selected<%end if%> value="0">关闭</option>
            </select></td>
        </tr>
		<tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >是否开启注册用户签收功能：</td>
          <td  colspan="2" align="left" > <select name="M_MAIN" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("M_MAIN")="1" then%>selected<%end if%> value="1">开启</option>
              <option <%if rs9("M_MAIN")="0" then%>selected<%end if%> value="0">关闭</option>
            </select><font color=red> 请确保评论功能已经开启</font></td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >是否显示评论者IP地址：</td>
          <td  colspan="2" align="left" > <select name="showip" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("showip")="1" then%>selected<%end if%> value="1">开启</option>
              <option <%if rs9("showip")="0" then%>selected<%end if%> value="0">关闭</option>
            </select></td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >评论默认状态：</td>
          <td  colspan="2" align="left" > <select name="reviewcheck" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("reviewcheck")="1" then%>selected<%end if%> value="1">已审核</option>
              <option <%if rs9("reviewcheck")="0" then%>selected<%end if%> value="0">未审核</option>
            </select></td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >审核评论菜单中：</td>
          <td  colspan="2" align="left" > <select name="reviewcheckshow" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("reviewcheckshow")="1" then%>selected<%end if%> value="1">显示</option>
              <option <%if rs9("reviewcheckshow")="0" then%>selected<%end if%> value="0">不显示</option>
            </select>
            已审核评论</td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >是否开启友情链接申请功能：</td>
          <td  colspan="2" align="left" > <select name="linkreg" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("linkreg")="1" then%>selected<%end if%> value="1">开启</option>
              <option <%if rs9("linkreg")="0" then%>selected<%end if%> value="0">关闭</option>
            </select></td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >友情链接管理菜单中：</td>
          <td  colspan="2" align="left" > <select name="linkpass" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("linkpass")="1" then%>selected<%end if%> value="1">显示</option>
              <option <%if rs9("linkpass")="0" then%>selected<%end if%> value="0">不显示</option>
            </select>
            已审核网站链接</td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >审核文章菜单中：</td>
          <td  colspan="2" align="left" > <select name="checkshow" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("checkshow")="1" then%>selected<%end if%> value="1">显示</option>
              <option <%if rs9("checkshow")="0" then%>selected<%end if%> value="0">不显示</option>
            </select>
            已审核文章</td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >首页是否显示会员排行榜：</td>
          <td  colspan="2" align="left" > <select name="showuser" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("showuser")="1" then%>selected<%end if%> value="1">显示</option>
              <option <%if rs9("showuser")="0" then%>selected<%end if%> value="0">不显示</option>
            </select></td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >首页是否显示专题文章：</td>
          <td  colspan="2" align="left" > <select name="showspecial" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("showspecial")="1" then%>selected<%end if%> value="1">显示</option>
              <option <%if rs9("showspecial")="0" then%>selected<%end if%> value="0">不显示</option>
            </select></td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >首页是否显示系统数据：</td>
          <td  colspan="2" align="left" > <select name="showdata" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("showdata")="1" then%>selected<%end if%> value="1">显示</option>
              <option <%if rs9("showdata")="0" then%>selected<%end if%> value="0">不显示</option>
            </select></td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >是否显示搜索栏：</td>
          <td  colspan="2" align="left" > <select name="showsearch" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("showsearch")="1" then%>selected<%end if%> value="1">显示</option>
              <option <%if rs9("showsearch")="0" then%>selected<%end if%> value="0">不显示</option>
            </select></td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >搜索程序样式：</td>
          <td  colspan="2" align="left" > <select name="search" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("search")="1" then%>selected<%end if%> value="1">简约型</option>
              <option <%if rs9("search")="0" then%>selected<%end if%> value="0">详细型</option>
            </select></td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >在标题后是否显示作者：</td>
          <td  colspan="2" align="left" > <select name="showauthor" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("showauthor")="1" then%>selected<%end if%> value="1">显示</option>
              <option <%if rs9("showauthor")="0" then%>selected<%end if%> value="0">不显示</option>
            </select></td>
        </tr>
        <tr bgcolor="#FFFFFF">
          <td width="25%"  align="right">是否显示论坛入口及论坛最新：</td>
          <td  colspan="2" align="left"><select name="topfont" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("topfont")="1" then%>selected<%end if%> value="1">显示</option>
              <option <%if rs9("topfont")="0" then%>selected<%end if%> value="0">不显示</option>
            </select></td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >是否显示投票系统：</td>
          <td  colspan="2" align="left" > <select name="showvote" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("showvote")="1" then%>selected<%end if%> value="1">显示</option>
              <option <%if rs9("showvote")="0" then%>selected<%end if%> value="0">不显示</option>
            </select></td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >是否显示文字链接：</td>
          <td  colspan="2" align="left" > <select name="showlink" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("showlink")="1" then%>selected<%end if%> value="1">显示</option>
              <option <%if rs9("showlink")="0" then%>selected<%end if%> value="0">不显示</option>
            </select></td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >是否显示LOGO链接：</td>
          <td  colspan="2" align="left" > <select name="showlinkmap" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("showlinkmap")="1" then%>selected<%end if%> value="1">显示</option>
              <option <%if rs9("showlinkmap")="0" then%>selected<%end if%> value="0">不显示</option>
            </select></td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >是否显示会员登录项目：</td>
          <td  colspan="2" align="left" > <select name="showclub" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("showclub")="1" then%>selected<%end if%> value="1">显示</option>
              <option <%if rs9("showclub")="0" then%>selected<%end if%> value="0">不显示</option>
            </select></td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >是否显示计数器及在线人数：</td>
          <td  colspan="2" align="left" > <select name="showcount" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("showcount")="1" then%>selected<%end if%> value="1">显示</option>
              <option <%if rs9("showcount")="0" then%>selected<%end if%> value="0">不显示</option>
            </select></td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >首页是否显示文章发表时间：</td>
          <td  colspan="2" align="left" > <select name="showtime" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("showtime")="1" then%>selected<%end if%> value="1">显示</option>
              <option <%if rs9("showtime")="0" then%>selected<%end if%> value="0">不显示</option>
            </select></td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >首页文章发表时间显示格式：</td>
          <td  colspan="2" align="left" > <select name="showyear" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("showyear")="1" then%>selected<%end if%> value="1">年月日</option>
              <option <%if rs9("showyear")="0" then%>selected<%end if%> value="0">月日</option>
            </select></td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >首页文章是否显示点击数：</td>
          <td  colspan="2" align="left" > <select name="showclick" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("showclick")="1" then%>selected<%end if%> value="1">显示</option>
              <option <%if rs9("showclick")="0" then%>selected<%end if%> value="0">不显示</option>
            </select></td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >执行操作后等待时间：</td>
          <td  colspan="2" align="left" > <select size="1" name="freetime" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <%for i=0 to 30%>
              <option <%if rs9("freetime")=i then%>selected<%end if%> value="<%=i%>"><%=i%></option>
              <%next%>
            </select>
            秒 </td>
        </tr>
        <tr class="TDtop"> 
          <td colspan="3"><div align="center">管理权限设置</div></td>
        </tr>
        <%if request.cookies(Forcast_SN)("ManageKEY")="super" and request.cookies(Forcast_SN)("ManagePurview")="99999" then%>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >系统管理员系统设置功能：</td>
          <td  colspan="2" align="left" > <select name="system" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("system")="1" then%>selected<%end if%> value="1">启用</option>
              <option <%if rs9("system")="0" then%>selected<%end if%> value="0">屏蔽</option>
            </select>
            点击“<a title=系统管理员系统设置功能详细说明 onMouseOver="window.status='系统管理员系统设置功能详细说明';return true;" onMouseOut="window.status='';return true;" href=javascript:alert("当系统管理员系统设置功能设置为启用时，系统管理员能进行系统设置，但该选项及以下五列不包括在内；当系统管理员系统设置功能设置为屏蔽时，系统管理员不能进行设置。")>这里</a>”查看详细说明。 
          </td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >系统管理员CSS编辑功能：</td>
          <td  colspan="2" align="left" > <select name="css" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("css")="1" then%>selected<%end if%> value="1">启用</option>
              <option <%if rs9("css")="0" then%>selected<%end if%> value="0">屏蔽</option>
            </select> </td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >系统管理员系统初始化功能：</td>
          <td  colspan="2" align="left" > <select name="init" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("init")="1" then%>selected<%end if%> value="1">启用</option>
              <option <%if rs9("init")="0" then%>selected<%end if%> value="0">屏蔽</option>
            </select> </td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >系统管理员管理权限设置：</td>
          <td  colspan="2" align="left" > <select name="powermana" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("powermana")="1" then%>selected<%end if%> value="1">启用</option>
              <option <%if rs9("powermana")="0" then%>selected<%end if%> value="0">屏蔽</option>
            </select> </td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >系统管理员投票管理功能：</td>
          <td  colspan="2" align="left" > <select name="votemana" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("votemana")="1" then%>selected<%end if%> value="1">启用</option>
              <option <%if rs9("votemana")="0" then%>selected<%end if%> value="0">屏蔽</option>
            </select> </td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >系统管理员友情链接管理：</td>
          <td  colspan="2" align="left" > <select name="linkmana" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("linkmana")="1" then%>selected<%end if%> value="1">启用</option>
              <option <%if rs9("linkmana")="0" then%>selected<%end if%> value="0">屏蔽</option>
            </select> </td>
        </tr>
        <%end if%>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >文章审核员管理小类选项：</td>
          <td  colspan="2" align="left" > <select name="shsmallgl" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("shsmallgl")="1" then%>selected<%end if%> value="1">启用</option>
              <option <%if rs9("shsmallgl")="0" then%>selected<%end if%> value="0">禁用</option>
            </select> </td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >文章审核员管理专题选项：</td>
          <td  colspan="2" align="left" > <select name="shspecialgl" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("shspecialgl")="1" then%>selected<%end if%> value="1">启用</option>
              <option <%if rs9("shspecialgl")="0" then%>selected<%end if%> value="0">禁用</option>
            </select> </td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >文章审核员管理评论选项：</td>
          <td  colspan="2" align="left" > <select name="shdelreview" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("shdelreview")="1" then%>selected<%end if%> value="1">启用</option>
              <option <%if rs9("shdelreview")="0" then%>selected<%end if%> value="0">禁用</option>
            </select> </td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >审核员修改他人文章权限：</td>
          <td  colspan="2" align="left" > <select name="checkmod" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("checkmod")="1" then%>selected<%end if%> value="1">可用</option>
              <option <%if rs9("checkmod")="0" then%>selected<%end if%> value="0">禁止</option>
            </select> </td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >审核员删除他人文章权限：</td>
          <td  colspan="2" align="left" > <select name="checkdel" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("checkdel")="1" then%>selected<%end if%> value="1">可用</option>
              <option <%if rs9("checkdel")="0" then%>selected<%end if%> value="0">禁止</option>
            </select> </td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >小类管理员管理小类选项：</td>
          <td  colspan="2" align="left" > <select name="smallgl" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("smallgl")="1" then%>selected<%end if%> value="1">启用</option>
              <option <%if rs9("smallgl")="0" then%>selected<%end if%> value="0">禁用</option>
            </select>
            启用后，录入员只能管理自己添加的小类，其他用户添加的小类无法管理 </td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >小类管理员管理专题选项：</td>
          <td  colspan="2" align="left" > <select name="specialgl" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("specialgl")="1" then%>selected<%end if%> value="1">启用</option>
              <option <%if rs9("specialgl")="0" then%>selected<%end if%> value="0">禁用</option>
            </select>
            启用后，录入员只能管理自己添加的专题，其他用户添加的专题无法管理 </td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >小类管理员管理评论选项：</td>
          <td  colspan="2" align="left" > <select name="delreview" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("delreview")="1" then%>selected<%end if%> value="1">启用</option>
              <option <%if rs9("delreview")="0" then%>selected<%end if%> value="0">禁用</option>
            </select> </td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >小类管理员修改密码选项：</td>
          <td  colspan="2" align="left" > <select name="inputmodpwd" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("inputmodpwd")="1" then%>selected<%end if%> value="1">启用</option>
              <option <%if rs9("inputmodpwd")="0" then%>selected<%end if%> value="0">禁用</option>
            </select> </td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >小类管理员管理本小类文章：</td>
          <td  colspan="2" align="left" > <select name="smallmana" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("smallmana")="1" then%>selected<%end if%> value="1">启用</option>
              <option <%if rs9("smallmana")="0" then%>selected<%end if%> value="0">禁用</option>
            </select>
            启用后，小类管理员可管理本小类的所有文章；禁用后，小类管理员仅能管理本人文章 </td>
        </tr>
        <tr class="TDtop"> 
          <td colspan="3"> <div align="center" >用户注册设置</div></td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >是否允许用户自注册：</td>
          <td  colspan="2" align="left" > <select name="reg" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("reg")="1" then%>selected<%end if%> value="1">允许</option>
              <option <%if rs9("reg")="0" then%>selected<%end if%> value="0">禁用</option>
            </select> </td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >注册后用户是否能发表文章：</td>
          <td  colspan="2" align="left" > <select name="fabiao" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("fabiao")="1" then%>selected<%end if%> value="1">立即发表</option>
              <option <%if rs9("fabiao")="0" then%>selected<%end if%> value="0">等待审核</option>
            </select> </td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >注册用户发表文章默认状态：</td>
          <td  colspan="2" align="left" > <select name="fabiaocheck" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("fabiaocheck")="1" then%>selected<%end if%> value="1">已审核</option>
              <option <%if rs9("fabiaocheck")="0" then%>selected<%end if%> value="0">未审核</option>
            </select>
            点击“<a title=注册用户发表文章默认状态详细说明 onMouseOver="window.status='注册用户发表文章默认状态详细说明';return true;" onMouseOut="window.status='';return true;"  href=javascript:alert("已审核指注册用户添加文章后，文章不必\n经过系统管理员或文章审核员的审核就可\n显示出来。而未审核指注册用户添加文章\n后，文章不直接显示出来，必须经过系统\n管理员和文章审核员的审核后才会显示。")>这里</a>”查看详细说明。 
          </td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >注册用户可修改已审核文章：</td>
          <td  colspan="2" align="left" > <select name="fabiaomod" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("fabiaomod")="1" then%>selected<%end if%> value="1">开启</option>
              <option <%if rs9("fabiaomod")="0" then%>selected<%end if%> value="0">关闭</option>
            </select>
            点击“<a title=注册用户可修改已审核文章详细说明 onMouseOver="window.status='注册用户可修改已审核文章详细说明';return true;" onMouseOut="window.status='';return true;" href=javascript:alert("当注册用户发表文章的默认状态为未审核\n时，该选项开启表示可以修改，但修改的\n文章将改为未审核状态（出于安全的考虑\n）；关闭表示不可修改；当注册用户发表\n文章默认状态为已审核时，此选项无效。")>这里</a>”查看详细说明。 
          </td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td colspan="3"  width="611"> <div align="center"> 
              <%if rs9("powermana")="1" and request.cookies(Forcast_SN)("ManagePurview")<>"99999" then%>
              <input type="hidden" name="system" value="<%=rs9("system")%>">
              <input type="hidden" name="css" value="<%=rs9("css")%>">
              <input type="hidden" name="init" value="<%=rs9("init")%>">
              <input type="hidden" name="powermana" value="<%=rs9("powermana")%>">
              <input type="hidden" name="linkmana" value="<%=rs9("linkmana")%>">
              <input type="hidden" name="votemana" value="<%=rs9("votemana")%>">
              <%end if
              rs9.close
              set rs9=nothing
              conn.close
              set conn=nothing
              ConnUser.close
              set ConnUser=nothing%>
              <input type="submit" name="Submit" value="提交" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <input type="reset" name="Submit2" value="重设" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
            </div></td>
        </tr>
</form>
</table>
</div>
</body>
</html>
<!--#include file=Admin_Bottom.asp-->
	<%else
		Show_Err("对不起，该功能已经被超级系统管理员关闭，您没有权限操作！")
		response.end
	end if
end if%>