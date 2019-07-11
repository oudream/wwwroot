<!--#include file =conn.asp-->
<!-- #include file="inc/const.asp" -->
<title><%=Forum_info(0)%>--管理页面</title>
<!--#include file=inc/forum_css.asp-->
<meta NAME=GENERATOR Content="Microsoft FrontPage 3.0" CHARSET=GB2312>
<BODY <%=Forum_body(11)%>>
<%
dim dbpath,bkfolder,bkdbname,fso,fso1
dim backpath
	if not master or session("flag")="" then
		Errmsg=Errmsg+"<br>"+"<li>本页面为管理员专用，请<a href=admin_index.asp target=_top>登陆</a>后进入。<br><li>您没有管理本页面的权限。"
		call dvbbs_error()
	else
		call main()
		conn.close
		set conn=nothing
	end if

sub main()
if request("action")="Restore" then
	Dbpath=request.form("Dbpath")
	backpath=request.form("backpath")
	if dbpath="" then
	response.write "请输入您要恢复成的数据库全名"	
	else
	Dbpath=server.mappath(Dbpath)
	end if
	backpath=server.mappath(backpath)
	'Response.write Backpath
	Set Fso=server.createobject("scripting.filesystemobject")
	if fso.fileexists(dbpath) then  					
		fso.copyfile Dbpath,Backpath
		response.write "成功恢复数据！"
	else
		response.write "备份目录下并无您的备份文件！"	
	end if
else
%>
<table align=center cellspacing=1 cellpadding=1 class=tableborder1>		  							  				
  				<tr>
  					<th height=25 >
   					&nbsp;&nbsp;<B>恢复论坛数据</B>( 需要FSO支持，FSO相关帮助请看微软网站 )
  					</th>
  				</tr>
  				<form method="post" action="ADMIN_RestoreData.asp?action=Restore">
  				<tr>
  					<td height=100 class=tablebody1>
  						&nbsp;&nbsp;备份数据库路径(相对)：<input type=text size=30 name=DBpath value="DataBackup\dvbbs5_Backup.MDB">&nbsp;&nbsp;<BR>
  						&nbsp;&nbsp;目标数据库路径(相对)：<input type=text size=30 name=backpath value="<%=db%>"><BR>&nbsp;&nbsp;填写您当前使用的数据库路径，如不想覆盖当前文件，可自行命名（注意路径是否正确），然后修改conn.asp文件，如果目标文件名和当前使用数据库名一致的话，不需修改conn.asp文件<BR>
						&nbsp;&nbsp;<input type=submit value="恢复数据"> <br>
  						-----------------------------------------------------------------------------------------<br>
  						&nbsp;&nbsp;在上面填写本程序的数据库路径全名，本程序的默认备份数据库文件为DataBackup\dvbbs_Backup.MDB，请按照您的备份文件自行修改。<br>
  						&nbsp;&nbsp;您可以用这个功能来备份您的法规数据，以保证您的数据安全！<br>
  						&nbsp;&nbsp;注意：所有路径都是相对与程序空间根目录的相对路径</font>
  					</td>
  				</tr>	
  				</form>
  			</table>
<%
end if

end sub
%>
