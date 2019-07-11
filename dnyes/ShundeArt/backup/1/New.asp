<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp" -->
<!--#include file="ChkManage.asp" -->
<%
if not(request.cookies(Forcast_SN)("ManageKEY")="super" or request.cookies(Forcast_SN)("KEY")="super") then
	Show_Err("对不起，您的后台管理权限不够操作！")
	response.end
else
	if init="1" or (request.cookies(Forcast_SN)("ManagePurview")="99999" and request.cookies(Forcast_SN)("purview")="99999") then
	%>
<html><head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=site.css rel=stylesheet>
<title><%=copyright%><%=version%>&nbsp;<%=ver%> - 系统初始化</title>
<script language="JavaScript">
<!--
function CheckAll(form){
	for (var i=0;i<form.elements.length;i++) {
		var e = form.elements[i];
		if (e.name != 'chkall') e.checked = form.chkall.checked; 
	}
}
//-->
</script>
</head>
<body topmargin="0">
<!--#include file=Admin_Top.asp-->
<table border="1" cellpadding="3" cellspacing="0" style="border-collapse: collapse" bordercolor="#C0C0C0" width="100%" id="AutoNumber1">
<form action=New_submit.asp method=post>
<tr>
	<td class="TDtop" width="100%" height=25 align="center" bgcolor="#ffb100" valign="middle">┊ <strong><font color="#FF0000">系统初始化确认</font></strong>┊</td>
</tr>
<tr>
	<td height="30" align="center" valign="middle"></td>
</tr>
<tr>
	<td align="center">
		<strong>警   告！</strong><br><br>
		<strong>初始化一般在系统第一次安装时进行！</strong><br>
		<strong>在系统正常运行期间初始化将可能丢失数据。</strong><br><br>
系统将被初始化，会删除数据库中的所有信息，但仍会保留数据库结构，确认要对系统做<font color=red>完全初始化</font>操作，请按“确定”，否则请按“取消”！
	</td>
</tr>
<tr>
<td height="30"></td>
</tr>
<tr>
	<td width="100%" align=center > 
		<input type=submit value="确定" name="alert_button" class="Bottom1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">&nbsp;&nbsp;&nbsp;&nbsp;
		<input type="button" value="取消" name="cancel" class="Botton" onClick="javascript:history.go(-1)" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
	</td>
</tr>
<tr>
	<td height="25" align="center" valign="middle" bgcolor="#7C7CB5"></td>
</tr>
</form>
</table>
<table border="1" cellpadding="3" cellspacing="0" style="border-collapse: collapse" bordercolor="#C0C0C0" width="100%" id="AutoNumber1">
<form action=New_submit.asp method=post>
<tr class="TDtop">
	<td width="100%" height="25" colspan="4"><p align="center" ><strong>你也可以选择<font color=red>单独初始化</font>系统的部分</strong></td>
</tr>
<tr>
	<td width="25%"><p align="center">
		<input type=checkbox value="<%=db_Dep_Table%>" name="selectdel" class="Bottom1" style="background-color: #ffffff;">部&nbsp;&nbsp;&nbsp;&nbsp;门
	<%rs=conn.execute("Select count(*) from "& db_Dep_Table)
		response.write "&nbsp;&nbsp;<font color=red>["& rs(0) &"]</font>"
		set rs=nothing%>
	</td>
	<td width="25%"><p align="center">
		<input type=checkbox value="<%=db_Type_Table%>" name="selectdel" class="Bottom1" style="background-color: #ffffff;">总&nbsp;&nbsp;&nbsp;&nbsp;栏
	<%rs=conn.execute("Select count(*) from "& db_Type_Table)
		response.write "&nbsp;&nbsp;<font color=red>["& rs(0) &"]</font>"
		set rs=nothing%>
	</td>
	<td width="25%"><p align="center">
		<input type=checkbox value="<%=db_BigClass_Table%>" name="selectdel" class="Bottom1" style="background-color: #ffffff;">大&nbsp;&nbsp;&nbsp;&nbsp;栏
	<%rs=conn.execute("Select count(*) from "& db_BigClass_Table)
		response.write "&nbsp;&nbsp;<font color=red>["& rs(0) &"]</font>"
		set rs=nothing%>
	</td>
	<td width="25%"><p align="center">
		<input type=checkbox value="<%=db_SmallClass_Table%>" name="selectdel" class="Bottom1" style="background-color: #ffffff;">小&nbsp;&nbsp;&nbsp;&nbsp;栏
	<%rs=conn.execute("Select count(*) from "& db_SmallClass_Table)
		response.write "&nbsp;&nbsp;<font color=red>["& rs(0) &"]</font>"
		set rs=nothing%>
	</td>
</tr>
<tr>
	<td width="25%"><p align="center">
		<input type=checkbox value="<%=db_Special_Table%>" name="selectdel" class="Bottom1" style="background-color: #ffffff;">专&nbsp;&nbsp;&nbsp;&nbsp;题
	<%rs=conn.execute("Select count(*) from "& db_Special_Table)
		response.write "&nbsp;&nbsp;<font color=red>["& rs(0) &"]</font>"
		set rs=nothing%>
	</td>
	<td width="25%"><p align="center">
		<input type=checkbox value="<%=db_News_Table%>" name="selectdel" class="Bottom1" style="background-color: #ffffff;">文&nbsp;&nbsp;&nbsp;&nbsp;章
	<%rs=conn.execute("Select count(*) from "& db_News_Table)
		response.write "&nbsp;&nbsp;<font color=red>["& rs(0) &"]</font>"
		set rs=nothing%>
	</td>
	<td width="25%"><p align="center">
		<input type=checkbox value="<%=db_NewsFile_Table%>" name="selectdel" class="Bottom1" style="background-color: #ffffff;">临&nbsp;&nbsp;&nbsp;&nbsp;时
	<%rs=conn.execute("Select count(*) from "& db_NewsFile_Table)
		response.write "&nbsp;&nbsp;<font color=red>["& rs(0) &"]</font>"
		set rs=nothing%>
	</td>
	<td width="25%"><p align="center">
		<input type=checkbox value="<%=db_UploadPic_Table%>" name="selectdel" class="Bottom1" style="background-color: #ffffff;">上传临时
	<%rs=conn.execute("Select count(*) from "& db_UploadPic_Table)
		response.write "&nbsp;&nbsp;<font color=red>["& rs(0) &"]</font>"
		set rs=nothing%>
	</td>
</tr>
<tr>
	<td width="25%"><p align="center">
		<input type=checkbox value="<%=db_Vote_Table%>" name="selectdel" class="Bottom1" style="background-color: #ffffff;">投&nbsp;&nbsp;&nbsp;&nbsp;票
	<%rs=conn.execute("Select count(*) from "& db_Vote_Table)
		response.write "&nbsp;&nbsp;<font color=red>["& rs(0) &"]</font>"
		set rs=nothing%>
	</td>
	<td width="25%"><p align="center">
		<input type=checkbox value="<%=db_Link_Table%>" name="selectdel" class="Bottom1" style="background-color: #ffffff;">友情链接
	<%rs=conn.execute("Select count(*) from "& db_Link_Table)
		response.write "&nbsp;&nbsp;<font color=red>["& rs(0) &"]</font>"
		set rs=nothing%>
	</td>
	<td width="25%"><p align="center">
		<input type=checkbox value="<%=db_Board_Table%>" name="selectdel" class="Bottom1" style="background-color: #ffffff;">公&nbsp;&nbsp;&nbsp;&nbsp;告
	<%rs=conn.execute("Select count(*) from "& db_Board_Table)
		response.write "&nbsp;&nbsp;<font color=red>["& rs(0) &"]</font>"
		set rs=nothing%>
	</td>
	<td width="25%"><p align="center">
		<input type=checkbox value="<%=db_Attach_Table%>" name="selectdel" class="Bottom1" style="background-color: #ffffff;">附&nbsp;&nbsp;&nbsp;&nbsp;件
	<%rs=conn.execute("Select count(*) from "& db_Attach_Table)
		response.write "&nbsp;&nbsp;<font color=red>["& rs(0) &"]</font>"
		set rs=nothing%>
	</td>
</tr>
<tr>
	<td width="25%"><p align="center">
		<input type=checkbox value="<%=db_Review_Table%>0" name="selectdel" class="Bottom1" style="background-color: #ffffff;">留&nbsp;&nbsp;&nbsp;&nbsp;言
	<%rs=conn.execute("Select count(*) from "& db_Review_Table &" where NewsId<1")
		response.write "&nbsp;&nbsp;<font color=red>["& rs(0) &"]</font>"
		set rs=nothing%>
	</td>
	<td width="25%"><p align="center">
		<input type=checkbox value="<%=db_Review_Table%>1" name="selectdel" class="Bottom1" style="background-color: #ffffff;">评&nbsp;&nbsp;&nbsp;&nbsp;论
	<%rs=conn.execute("Select count(*) from "& db_Review_Table &" where NewsId>0")
		response.write "&nbsp;&nbsp;<font color=red>["& rs(0) &"]</font>"
		set rs=nothing%>
	</td>
	<td width="25%"><p align="center">
		<input type=checkbox value="<%=db_User_Table%>" name="selectdel" class="Bottom1" style="background-color: #ffffff;">普通用户
	<%rs=ConnUser.execute("Select count(*) from "& db_User_Table)
		response.write "&nbsp;&nbsp;<font color=red>["& rs(0) &"]</font>"
		set rs=nothing%>
	</td>
	<td width="25%"><p align="center">
		<input type=checkbox value="<%=db_Manage_Table%>" name="selectdel" class="Bottom1" style="background-color: #ffffff;">管理用户
	<%rs=conn.execute("Select count(*) from "& db_Manage_Table)
		response.write "&nbsp;&nbsp;<font color=red>["& rs(0) &"]</font>"
		set rs=nothing%>
	</td>
</tr>
<tr>
<td colspan=4  height=25 align="center">
	<input type=hidden name=action value='del'>
	<input type=checkbox name=chkall value=on onclick="CheckAll(this.form)" style=" background-color: #ffffff;">全选&nbsp;&nbsp;
	<input type=submit name=action1 onclick="{if(confirm('初始化选定的项目吗？')){this.document.check.submit();return true;}return false;}" value="初始化" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
</td>
</tr>
</form>
</table>
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#C0C0C0" width="100%" id="AutoNumber1">
<%
Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from counters"
rs.open sql,conn,1,3
if request.form("Ch_Count")=1 then
	if request.form("total")="" then
		rs("total")=0
	else
		rs("total")=request.form("total")
	end if
	if request.form("today")="" then
		rs("today")=0
	else
		rs("today")=request.form("today")
	end if
	if request.form("yesterday")="" then
		rs("yesterday")=0
	else
		rs("yesterday")=request.form("yesterday")
	end if
	if request.form("months")="" then
		rs("month")=0
	else
		rs("month")=request.form("months")
	end if
	if request.form("bmonth")="" then
		rs("bmonth")=0
	else
		rs("bmonth")=request.form("bmonth")
	end if
	rs.update
end if
%>
<form method="POST" action="New.asp?ID=<&=ID&>" id=form name=form>
<tr class="TDTop">
	<td width="100%" height="25" colspan=5 align="center"><b>统计数据初始化</b></td>
</tr>
<tr> 
	<td width="20%" align="center">访问总数</td>
	<td width="20%" align="center">今日访问</td>
	<td width="20%" align="center">昨日访问</td>
	<td width="20%" align="center">本月访问</td>
	<td width="20%" align="center">上月访问</td>
</tr>
<tr align="center">
	<td height="25"><input type="text" name="total" value="<%=rs("total")%>" size="9"></td>
	<td height="25"><input type="text" name="today" value="<%=rs("today")%>" size="9"></td>
	<td height="25"><input type="text" name="yesterday" value="<%=rs("yesterday")%>" size="9"></td>
	<td height="25"><input type="text" name="months" value="<%=rs("month")%>" size="9"></td>
	<td height="25"><input type="text" name="bmonth" value="<%=rs("bmonth")%>" size="9"></td>
</tr>
<tr align="center"> 
	<td height="25" colspan=5>
		<input type="hidden" name="Ch_Count" value="1">
		<input type="reset" value="重 置" name="cmdcance2l" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">&nbsp;&nbsp;
		<input type="submit" value="修 改" name="cmdok2" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
	</td>
</tr>
</form>
<%             
	set rs=nothing             
	conn.close             
	set conn=nothing             
%>
</table>
</body>
</html>
<!--#include file=Admin_Bottom.asp-->
	<%else
		Show_Err("对不起，该功能已经被超级系统管理员关闭，您没有权限操作！")
		response.end
	end if
end if%>