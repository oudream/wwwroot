<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="inc/config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp" -->
<!--#include file="ChkURL.asp"-->
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
	if rs9("system")="1" or request.cookies(Forcast_SN)("purview")="99999" then
		'这里一定要先关闭数据库,并且保证在压缩时没有使用数据库内容
		rs9.close
		set rs9=nothing
		conn.close
		set conn=nothing
		%>
<HTML>
<HEAD>
<TITLE>SQL Server 数据库备份与恢复</TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=site.css rel=stylesheet>
</HEAD>
<body topmargin="0">
<!--#include file=Admin_Top.asp-->
<table border="1" cellpadding="3" cellspacing="0" style="border-collapse: collapse" bordercolor="#C0C0C0" width="100%" id="AutoNumber1">
<form method="post" name=myform>
<tr class="TDtop"> 
<td colspan="2" height=25 width="100%" > 
	<div align="center">┊ <strong>SQL Server 数据库备份与恢复</strong> ┊</div>
</td>
</tr>
<tr>
<td align="right">
选择操作：
</td>
<td> 
<INPUT TYPE="radio" NAME="act" id="act_backup"  value="backup" <%if request("act")="" or request("act")="backup" then%> checked<%end if%>><label for=act_backup>备份</label>　
<INPUT TYPE="radio" NAME="act" id="act_restore" value="restore" <%if request("act")="restore" then%> checked<%end if%>><label for=act_restore>恢复（确保无其他用户访问网站）</label>
</td>
</tr>
<tr> 
<td align="right" width="40%"> 
SQL服务器：
</td>
<td width="60%"> 
<INPUT TYPE="text" NAME="sqlserver" value="<%if request("sqlserver")<>"" then%><%=request("sqlserver")%><%else%>127.0.0.1<%end if%>">
</td>
</tr>
<tr>
<td align="right"> 
用户名：
</td>
<td> 
<INPUT TYPE="text" NAME="sqlname" value="<%if request("sqlname")<>"" then%><%=request("sqlname")%><%else%>sa<%end if%>">
</td>
</tr>
<tr>
<td align="right"> 
用户密码：
</td>
<td> 
<INPUT TYPE="password" NAME="sqlpassword" value="<%=request("sqlpassword")%>">
</td>
</tr>
<tr>
<td align="right">
登陆超时:
</td>
<td> 
<INPUT TYPE="text" NAME="sqlLoginTimeout" value="<%if request("sqlLoginTimeout")<>"" then%><%=request("sqlLoginTimeout")%><%else%>15<%end if%>">
</td>
</tr>
<tr>
<td align="right">
数据库名：
</td>
<td> 
<INPUT TYPE="text" NAME="databasename" value="<%if request("databasename")<>"" then%><%=request("databasename")%><%else%>Feitium<%end if%>">
</td>
</tr>
<tr>
<td align="right">
文件路径：
</td>
<td> 
<INPUT TYPE="text" NAME="bak_file" value="<%if request("bak_file")<>"" then%><%=request("bak_file")%><%else%>d:\$dbname$<%=date()%>.bak<%end if%>">(备份或恢复的文件路径)
</td>
</tr>
<tr>
<td colspan="2" align="center" height=25 width="100%" > 
		<%
		'SQL Server 数据库的备份与恢复!
		dim sqlserver,sqlname,sqlpassword,sqlLoginTimeout,databasename,bak_file,act
		sqlserver = request("sqlserver")		'sql服务器
		sqlname = request("sqlname")			'用户名
		sqlpassword = request("sqlpassword")		'密码
		sqlLoginTimeout = request("sqlLoginTimeout")	'登陆超时
		databasename = trim(request("databasename"))	'数据库名
		bak_file = trim(request("bak_file"))		'备份路径
		bak_file = replace(bak_file,"$dbname$",databasename)
		act = lcase(request("act"))			'选项
		if databasename = "" then
			response.write "<font color=green>[请输入数据库名，“$dbname$”将被置换为数据库名]</font>"
		else
			if act = "backup" then
				Set srv=Server.CreateObject("SQLDMO.SQLServer")
				srv.LoginTimeout = sqlLoginTimeout
				srv.Connect sqlserver,sqlname, sqlpassword
				Set bak = Server.CreateObject("SQLDMO.Backup")
				bak.Database=databasename
				bak.Devices=Files
				bak.Files=bak_file
				bak.SQLBackup srv
				if err.number>0 then
					response.write err.number&"<font color=red><br>"
					response.write err.description&"</font>"
				else
					Response.write "<font color=green>备份成功!</font>"
				end if
			elseif act = "restore" then
				'恢复时要在没有使用数据库时进行！
				Set srv=Server.CreateObject("SQLDMO.SQLServer")
				srv.LoginTimeout = sqlLoginTimeout
				srv.Connect sqlserver,sqlname, sqlpassword
				Set rest=Server.CreateObject("SQLDMO.Restore")
				rest.Action=1			' full db restore
				rest.Database=databasename
				rest.Devices=Files
				rest.Files=bak_file
				rest.ReplaceDatabase=True 'Force restore over existing database
				if err.number>0 then
					response.write err.number&"<font color=red><br>"
					response.write err.description&"</font>"
				else
					rest.SQLRestore srv
					Response.write "<font color=green>恢复成功!</font>"
				end if
			end if
		end if%>
</td>
</tr>
<tr>
<td colspan="2" align="center" height=25 width="100%" class="TDtop"> 
<input type="submit" value="确定" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
</td>
</tr>
</form>
</table>
</BODY>
</HTML>
<!--#include file=Admin_Bottom.asp-->
	<%end if
END IF%>