<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
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

		'**************************************************
		'函数名：IsObjInstalled
		'作  用：检查组件是否已经安装
		'参  数：strClassString ----组件名
		'返回值：True  ----已经安装
		'       False ----没有安装
		'**************************************************
		dim ObjInstalled
		ObjInstalled=IsObjInstalled("Scripting.FileSystemObject")
		
		Function IsObjInstalled(strClassString)
			On Error Resume Next
			IsObjInstalled = False
			Err = 0
			Dim xTestObj
			Set xTestObj = Server.CreateObject(strClassString)
			If 0 = Err Then IsObjInstalled = True
			Set xTestObj = Nothing
			Err = 0
		End Function
		ThisNameFile="Db_Compact.asp"   '定义本文件的文件名
		%>
<html><head> 
<title>ACCESS数据库压缩</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=site.css rel=stylesheet>
</head> 
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!--#include file=Admin_Top.asp-->
<table cellpadding="0" cellspacing="0" border="1" width="100%" style="border-collapse: collapse" bordercolor="#C0C0C0" width="100%" id="AutoNumber1">
<tr class="TDtop"> 
	<td height="30" colspan="2" align="center"><strong>数 据 库 管 理</strong></td>
</tr>
<tr > 
	<td height="25" colspan="2" align="center"><font color="#FF0000">使用本功能需要服务器支持FSO(Scripting.FileSystemObject)!</font></td>
</tr>
<tr> 
	<td width="20%" height="30" align="center"><strong>管理导航：</strong></td>
	<td width="80%" height="30" align="center">
		<a href="<%=ThisNameFile%>?Action=BackupData">备份数据库</a>
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="<%=ThisNameFile%>?Action=RestoreData">恢复数据库</a> 
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="<%=ThisNameFile%>?action=CompressData">压缩数据库</a> 
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="<%=ThisNameFile%>?action=SpaceSize">系统空间占用</a>
	</td>
</tr>
</table>
<%
Head()
dim action
action=trim(request("action"))

dim dbpath,bkfolder,bkdbname,fso,fso1
select case action
case "CompressData"		'压缩数据
		dim tmprs
		dim allarticle
		dim Maxid
		dim topic,username,dateandtime,body
		call CompressData()

case "BackupData"		'备份数据
		if request("act")="Backup" then
		call updata()
		else		
		call BackupData()
		end if

case "RestoreData"		'恢复数据
	dim backpath
		if request("act")="Restore" then
			Dbpath=request.form("Dbpath")
			backpath=request.form("backpath")
			if dbpath="" then
				response.write "<div align=center><font color=#FF0000>请输入您要恢复成的数据库全名</font> </div>"	
			else
				Dbpath=server.mappath(Dbpath)
			end if
			backpath=server.mappath(backpath)
			
			Set Fso=server.createobject("scripting.filesystemobject")
			if fso.fileexists(dbpath) then  					
				fso.copyfile Dbpath,Backpath
				response.write "<div align=center><font color=#FF0000>成功恢复数据！</font> </div>"
			else
				response.write "<div align=center><font color=#FF0000>备份目录下并无您的备份文件！</font> </div>"	
			end if
		else
			call RestoreData()
		end if	
case "SpaceSize"		'系统空间占用
		call SpaceSize()	
end select
%>
<!--#include file=Admin_Bottom.asp-->
<%'====================系统空间占用=======================
sub SpaceSize()
On error resume next
%>
<table cellpadding="0" cellspacing="0" border="1" width="100%" style="border-collapse: collapse" bordercolor="#C0C0C0" width="100%" id="AutoNumber1">
<tr class="TDtop">
	<td height="25" colspan="2" align="center" > <strong>系统空间占用情况</strong></td>
</tr>
<%
	If ObjInstalled=false Then
		Response.Write "<tr align=center><td height=30 colspan=3> <div align=center><b><font color=red>你的服务器不支持 FSO(Scripting.FileSystemObject)! 不能使用本功能</font></b> </div></td></tr>"
	else
		l_path=left(Request.ServerVariables("PATH_INFO"),instrrev(Request.ServerVariables("PATH_INFO"),"/"))
%>
<tr>
	<td width="22%" height="25">&nbsp;【数据库】<img src="images/mdb.gif" height=22>：</td>
	<td width="78%" height="25">&nbsp;<%=ShowFileSize(l_path & db)%></td>
</tr>
<%
		cpath=Server.MapPath(l_path)
		set fsoBrowse=CreateObject("Scripting.FileSystemObject")
		Set theFolder=fsoBrowse.GetFolder(cpath)
		Set theFiles=theFolder.files
		Set theSubFolders=theFolder.SubFolders
		for each SubF in theSubFolders%>
<tr>
	<td>&nbsp;【<%=SubF.name%>】：</td>
	<td>&nbsp;<img src="images/Admin_db_bar.gif" width=<%=drawbar(SubF.name)%> height=10><%showSpaceinfo(SubF.name)%></td>
</tr>
		<%next
		set fsoBrowse=nothing
		set theFolder=nothing
		set theSubFolders=nothing
		%>
<tr>
	<td>&nbsp;系统空间总计：</td>
	<td>&nbsp;<img src="images/Admin_db_bar.gif" width=400 height=10>&nbsp;<%showspecialspaceinfo("All")%></td>
</tr> 
<%end if%>
</table>
<%
end sub
%>
<%'====================恢复数据库=========================
sub RestoreData()
If IsSqlDataBase = 1 Then
	SQLUserReadme()
	Exit Sub
End If
%>
<table cellpadding="0" cellspacing="0" border="1" width="100%" style="border-collapse: collapse" bordercolor="#C0C0C0" width="100%" id="AutoNumber1">
<form method="post" action="<%=ThisNameFile%>?action=RestoreData&act=Restore">		
<tr class="TDtop"> 
	<td height=25 colspan="2" align="center" > &nbsp;&nbsp;<B>恢复数据</B> </td>
</tr>
<tr> 
	<td width="200" height="30" align="right">备份数据库路径（相对）：</td>
	<td height="30">&nbsp;<input type=text size=40 name=DBpath value="DataBackup/News30000.asa"></td>
</tr>
<tr align="center"> 
	<td height="30" colspan="2"><font color="#FF0000">填写备份数据库文件名替换“News30000.asa”，可自行命名（注意路径不能改变），可以参考下面显示的数据库名称</font></td>
</tr>
<tr> 
	<td width="200" height="30" align="right">当前数据库路径（相对）：</td>
	<td>&nbsp;<input type=text size=40 name=backpath value="<%=db%>"></td>
</tr>
<tr align="center"> 
	<td height="30" colspan="2">填写您当前使用的数据库路径，如不想覆盖当前文件，可自行命名（注意路径是否正确）</td>
</tr>				
<tr align="center"> 
	<td height=30 colspan="2">
		<input type=submit value=" &nbsp;恢复数据&nbsp;" <%If ObjInstalled=false Then response.Write "disabled"%> style="cursor: hand;background-color: #cccccc;">
	</td>
</tr>
<%
	If ObjInstalled=false Then
		Response.Write "<tr align=center><td height=30 colspan=2> <div align=center><b><font color=red>你的服务器不支持 FSO(Scripting.FileSystemObject)! 不能使用本功能</font></b> </div></td></tr>"
	end if
%>	
</form>
</table>
<%
end sub

'====================备份数据库=========================
sub BackupData()
If IsSqlDataBase = 1 Then
	SQLUserReadme()
	Exit Sub
End If
%>
<table cellpadding="0" cellspacing="0" border="1" width="100%" style="border-collapse: collapse" bordercolor="#C0C0C0" width="100%" id="AutoNumber1">
<form method="post" action="<%=ThisNameFile%>?action=BackupData&act=Backup">
<tr class="TDtop"> 
	<td height=25  align="center"> <strong>备份数据库</strong></td>
</tr>
<tr> 
	<td height=100 >
&nbsp;&nbsp;当前数据库路径(相对路径)：<input type=text size=40 name=DBpath value="<%=db%>">
&nbsp;相对路径<BR>
&nbsp;&nbsp;备份数据库目录(相对路径)：<input type=text size=40 name=bkfolder value="Databackup"> &nbsp;如目录不存在，程序将自动创建<BR>
&nbsp;&nbsp;备份数据库名称(填写名称)：<input type=text size=40 name=bkDBname value=<%=date%>.asa> &nbsp;如备份目录有该文件，将覆盖，如没有，将自动创建
	</td>
</tr>
<tr align="center"> 
	<td height="30" > 
		<input type=submit value=" &nbsp;开始备份&nbsp;" <%If ObjInstalled=false Then response.Write "disabled"%> style="cursor: hand;background-color: #cccccc;">
	</td>
</tr>
<%
	If ObjInstalled=false Then
		Response.Write "<tr align=center><td height=30 > <div align=center><b><font color=red>你的服务器不支持 FSO(Scripting.FileSystemObject)! 不能使用本功能</font></b> </div></td></tr>"
	end if
%>
</form>
</table>
<%
end sub

sub updata()
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
			response.write "<br><div align=center><font color=#FF0000>备份数据库成功，您备份的数据库路径为" &bkfolder& "\"& bkdbname & "</font> </div>"
		Else
			response.write "<br><div align=center><font color=#FF0000>找不到您所需要备份的文件。</font> </div>"
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
'-------------根据指定名称生成目录-----------------------
Function MakeNewsDir(foldername)
	dim f
    Set fso1 = CreateObject("Scripting.FileSystemObject")
        Set f = fso1.CreateFolder(foldername)
        MakeNewsDir = True
    Set fso1 = nothing
End Function


'====================压缩数据库 =========================
sub CompressData()

If IsSqlDataBase = 1 Then
	SQLUserReadme()
	Exit Sub
End If
%>
<table cellpadding="0" cellspacing="0" border="1" width="100%" style="border-collapse: collapse" bordercolor="#C0C0C0" width="100%" id="AutoNumber1">
<form action="<%=ThisNameFile%>?action=CompressData" method="post">
<tr class="TDtop"> 
	<td height=25 align="center" class="forumrow"><b>数据库在线压缩</b></td>
</tr>
<tr> 
	<td height=25 align="center" class="forumrow"><b>注意：</b> 输入数据库所在相对路径,并且输入数据库名称（正在使用中数据库不能压缩，请选择备份数据库进行压缩操作）</td>
</tr>
<tr> 
	<td height="25" align="center" class="forumrow"><font color="#FF6600"><b>注意：</b></font><font color="#FF0000">压缩前，强烈建议先备份数据库，以免发生意外错误！</font></td>
</tr>
<tr> 
	<td height="25" align="center" class="forumrow">压缩数据库：<input name="dbpath" type="text" value=<%=db%> size="40"></td>
</tr>
<tr> 
	<td height="25" align="center" class="forumrow"><input type="checkbox" name="boolIs97" value="True">如果使用 Access 97 数据库请选择 (默认为 Access 2000 数据库)</td>
</tr>
<tr>
	<td height="30" align="center" class="forumrow"><input type="submit" value=" &nbsp;开始压缩&nbsp;" <%If ObjInstalled=false Then response.Write "disabled"%> style="cursor: hand;background-color: #cccccc;"></td>
</tr>
<%
	If ObjInstalled=false Then
		Response.Write "<tr align=center><td height=30 colspan=3> <div align=center><b><font color=red>你的服务器不支持 FSO(Scripting.FileSystemObject)! 不能使用本功能</font></b> </div></td></tr>"
	end if
%>
</form>
</table>
<%
dim dbpath,boolIs97
dbpath = request("dbpath")
boolIs97 = request("boolIs97")

If dbpath <> "" Then
dbpath = server.mappath(dbpath)
	response.write(CompactDB(dbpath,boolIs97))
End If

end sub

'=====================压缩参数=========================
Function CompactDB(dbPath, boolIs97)
Dim fso, Engine, strDBPath,JET_3X
strDBPath = left(dbPath,instrrev(DBPath,"\"))
Set fso = CreateObject("Scripting.FileSystemObject")

If fso.FileExists(dbPath) Then
	fso.CopyFile dbpath,strDBPath & "temp.mdb"
	Set Engine = CreateObject("JRO.JetEngine")

	If boolIs97 = "True" Then
		Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp.mdb", _
		"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp1.mdb;" _
		& "Jet OLEDB:Engine Type=" & JET_3X
	Else
		Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp.mdb", _
		"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp1.mdb"
	End If

fso.CopyFile strDBPath & "temp1.mdb",dbpath
fso.DeleteFile(strDBPath & "temp.mdb")
fso.DeleteFile(strDBPath & "temp1.mdb")
Set fso = nothing
Set Engine = nothing

	CompactDB = "<br><div align=center><font color=#FF0000>你的数据库: " & dbpath & "  已经压缩成功!</font> </div>" & vbCrLf

Else
	CompactDB = "<br><div align=center><font color=#FF0000>数据库名称或路径不正确. 请重试!</font> </div>" & vbCrLf
End If

End Function


'=====================系统空间参数=========================
	Sub ShowSpaceInfo(drvpath)
 		dim fso,d,size,showsize,bai
 		set fso=server.createobject("scripting.filesystemobject")
 		drvpathroot=server.mappath(".")
 		set d=fso.getfolder(drvpathroot)
 		totalsize=d.size

 		set fso=server.createobject("scripting.filesystemobject") 		
 		drvpath=server.mappath(drvpath) 		 		
 		set d=fso.getfolder(drvpath) 		
 		size=d.size
 		bai=Cint(size/totalsize*10000)/100&"%&nbsp;&nbsp;" 
 		showsize=size & "字节" 
 		if size>1024 then
 		   size=(Size/1024)
 		   showsize=formatnumber(size,2) & "KB"
 		end if
 		if size>1024 then
 		   size=(size/1024)
 		   showsize=formatnumber(size,2) & "MB"		
 		end if
 		if size>1024 then
 		   size=(size/1024)
 		   showsize=formatnumber(size,2) & "GB"	   
 		end if   
 		response.write "<font face=verdana>"& bai & showsize & "</font>"
 	End Sub	
 	
 	Sub Showspecialspaceinfo(method)
 		dim fso,d,fc,f1,size,showsize,drvpath 		
 		set fso=server.createobject("scripting.filesystemobject")
 		drvpath=server.mappath(".")
 		set d=fso.getfolder(drvpath) 		
 		
 		if method="All" then 		
 			size=d.size
 		elseif method="Program" then
 			set fc=d.Files
 			for each f1 in fc
 				size=size+f1.size
 			next	
 		end if	
 		
 		showsize=size & "Byte" 
 		if size>1024 then
 		   size=(Size/1024)
 		   showsize=size & "KB"
 		end if
 		if size>1024 then
 		   size=(size/1024)
 		   showsize=formatnumber(size,2) & "MB"		
 		end if
 		if size>1024 then
 		   size=(size/1024)
 		   showsize=formatnumber(size,2) & "GB"	   
 		end if   
 		response.write "<font face=verdana>" & showsize & "</font>"
 	end sub 	 	 	

 	Function ShowFileSize(DrvPath_FileName)
 		dim fs,thisfile,whichfile,Size
		whichfile=Server.MapPath(DrvPath_FileName)
 		set fs= CreateObject("Scripting.FileSystemObject")
 		Set thisfile = fs.GetFile(whichfile)
 		Size=thisfile.size

 		if size>1024 then
 		   size=(Size/1024)
 		   showsize=size & "KB"
 		end if
 		if size>1024 then
 		   size=(size/1024)
 		   showsize=formatnumber(size,2) & "MB"		
 		end if
 		if size>1024 then
 		   size=(size/1024)
 		   showsize=formatnumber(size,2) & "GB"	   
 		end if   

 		ShowFileSize=showsize
 	End Function 	

 	Function Drawbar(drvpath)
 		dim fso,drvpathroot,d,size,totalsize,barsize
 		set fso=server.createobject("scripting.filesystemobject")
 		drvpathroot=server.mappath(".")
 		set d=fso.getfolder(drvpathroot)
 		totalsize=d.size
 		
 		drvpath=server.mappath(drvpath) 		
 		set d=fso.getfolder(drvpath)
 		size=d.size
 		
 		barsize=cint((size/totalsize)*400)
 		Drawbar=barsize
 	End Function 	
 	
 	Function Drawspecialbar()
 		dim fso,drvpathroot,d,fc,f1,size,totalsize,barsize
 		set fso=server.createobject("scripting.filesystemobject")
 		drvpathroot=server.mappath(".")
 		set d=fso.getfolder(drvpathroot)
 		totalsize=d.size
 		
 		set fc=d.files
 		for each f1 in fc
 			size=size+f1.size
 		next	
 		
 		barsize=cint((size/totalsize)*400)
 		Drawspecialbar=barsize
 	End Function 	
	end if
end if%>