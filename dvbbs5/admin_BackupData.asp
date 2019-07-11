<!--#include file =conn.asp-->
<!-- #include file="inc/const.asp" -->
<title><%=Forum_info(0)%>--管理页面</title>
<!--#include file=inc/forum_css.asp-->
<meta NAME=GENERATOR Content=""Microsoft FrontPage 3.0"" CHARSET=GB2312>
<BODY  <%=Forum_body(11)%>>
<%
dim dbpath,bkfolder,bkdbname,fso,fso1
	if not master or session("flag")="" then
		Errmsg=Errmsg+"<br>"+"<li>本页面为管理员专用，请<a href=admin_index.asp target=_top>登陆</a>后进入。<br><li>您没有管理本页面的权限。"
		call dvbbs_error()
	else
		call main()
		conn.close
		set conn=nothing
	end if
sub main()
if request("action")="Backup" then
	call backupdata()
else
%>
		 		<table cellspacing=1 cellpadding=1 align=center class=tableborder1>		  							  				
  				<tr>
  					<th height=25 >
  					&nbsp;&nbsp;<B>备份论坛数据</B>( 需要FSO支持，FSO相关帮助请看微软网站 )
  					</th>
  				</tr>
  				<form method="post" action="ADMIN_BackupData.asp?action=Backup">
  				<tr>
  					<td height=100 class=tablebody1>
  						&nbsp;&nbsp;
						当前数据库路径(相对路径)：<input type=text size=15 name=DBpath value="<%=db%>"><BR>&nbsp;&nbsp;
						备份数据库目录(相对路径)：<input type=text size=15 name=bkfolder value=Databackup>&nbsp;如目录不存在，程序将自动创建<BR>&nbsp;&nbsp;
						备份数据库名称(填写名称)：<input type=text size=15 name=bkDBname value=dvbbs5.MDB>&nbsp;如备份目录有该文件，将覆盖，如没有，将自动创建<BR>
						&nbsp;&nbsp;<input type=submit value="确定"><br>
  						-----------------------------------------------------------------------------------------<br>
  						&nbsp;&nbsp;在上面填写本程序的数据库路径全名，本程序的默认数据库文件为Data\dvbbs5.MDB，<B>请一定不能用默认名称命名备份数据库</B><br>
  						&nbsp;&nbsp;您可以用这个功能来备份您的法规数据，以保证您的数据安全！<br>
  						&nbsp;&nbsp;注意：所有路径都是相对与程序空间根目录的相对路径				</font>
  					</td>
  				</tr>	
  				</form>
  			</table>
<%
end if

end sub

sub backupdata()
		Dbpath=request.form("Dbpath")
		Dbpath=server.mappath(Dbpath)
		bkfolder=request.form("bkfolder")
		bkdbname=request.form("bkdbname")
		Set Fso=server.createobject("scripting.filesystemobject")
		if fso.fileexists(dbpath) then
			If CheckDir(bkfolder) = True Then
			fso.copyfile dbpath,bkfolder& "\"& bkdbname
			else
			MakeNewsDir bkfolder
			fso.copyfile dbpath,bkfolder& "\"& bkdbname
			end if
			response.write "备份数据库成功，您备份的数据库路径为" &bkfolder& "\"& bkdbname
		Else
			response.write "找不到您所需要备份的文件。"
		End if
end sub
'------------------检查某一目录是否存在-------------------
Function CheckDir(FolderPath)
	folderpath=Server.MapPath(".")&"\"&folderpath
    Set fso1 = CreateObject("Scripting.FileSystemObject")
    If fso1.FolderExists(FolderPath) then
       '存在
       CheckDir = True
    Else
       '不存在
       CheckDir = False
    End if
    Set fso1 = nothing
End Function
'-------------根据指定名称生成目录---------
Function MakeNewsDir(foldername)
	dim f
    Set fso1 = CreateObject("Scripting.FileSystemObject")
        Set f = fso1.CreateFolder(foldername)
        MakeNewsDir = True
    Set fso1 = nothing
End Function
%>