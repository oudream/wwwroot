<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp" -->
<!--#include file="ChkURL.asp"-->
<!--#include file="ChkManage.asp" -->
<%
IF request.cookies(Forcast_SN)("ManageKEY")<>"super" or request.cookies(Forcast_SN)("ManagePurview")<>"99999" then
	Show_Err("�Բ������ĺ�̨����Ȩ�޲�����")
	response.end
else
	response.buffer=true
	Response.Expires=0
	
	Set rs9 = Server.CreateObject("ADODB.Recordset")
	sql9 ="SELECT * From "& db_System_Table &" Order By id DESC"
	RS9.open sql9,Conn,1,1
	if rs9("system")="1" or request.cookies(Forcast_SN)("purview")="99999" then
		'����һ��Ҫ�ȹر����ݿ�,���ұ�֤��ѹ��ʱû��ʹ�����ݿ�����
		rs9.close
		set rs9=nothing
		conn.close
		set conn=nothing

		'**************************************************
		'��������IsObjInstalled
		'��  �ã��������Ƿ��Ѿ���װ
		'��  ����strClassString ----�����
		'����ֵ��True  ----�Ѿ���װ
		'       False ----û�а�װ
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
		ThisNameFile="Db_Compact.asp"   '���屾�ļ����ļ���
		%>
<html><head> 
<title>ACCESS���ݿ�ѹ��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=site.css rel=stylesheet>
</head> 
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!--#include file=Admin_Top.asp-->
<table cellpadding="0" cellspacing="0" border="1" width="100%" style="border-collapse: collapse" bordercolor="#C0C0C0" width="100%" id="AutoNumber1">
<tr class="TDtop"> 
	<td height="30" colspan="2" align="center"><strong>�� �� �� �� ��</strong></td>
</tr>
<tr > 
	<td height="25" colspan="2" align="center"><font color="#FF0000">ʹ�ñ�������Ҫ������֧��FSO(Scripting.FileSystemObject)!</font></td>
</tr>
<tr> 
	<td width="20%" height="30" align="center"><strong>��������</strong></td>
	<td width="80%" height="30" align="center">
		<a href="<%=ThisNameFile%>?Action=BackupData">�������ݿ�</a>
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="<%=ThisNameFile%>?Action=RestoreData">�ָ����ݿ�</a> 
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="<%=ThisNameFile%>?action=CompressData">ѹ�����ݿ�</a> 
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="<%=ThisNameFile%>?action=SpaceSize">ϵͳ�ռ�ռ��</a>
	</td>
</tr>
</table>
<%
Head()
dim action
action=trim(request("action"))

dim dbpath,bkfolder,bkdbname,fso,fso1
select case action
case "CompressData"		'ѹ������
		dim tmprs
		dim allarticle
		dim Maxid
		dim topic,username,dateandtime,body
		call CompressData()

case "BackupData"		'��������
		if request("act")="Backup" then
		call updata()
		else		
		call BackupData()
		end if

case "RestoreData"		'�ָ�����
	dim backpath
		if request("act")="Restore" then
			Dbpath=request.form("Dbpath")
			backpath=request.form("backpath")
			if dbpath="" then
				response.write "<div align=center><font color=#FF0000>��������Ҫ�ָ��ɵ����ݿ�ȫ��</font> </div>"	
			else
				Dbpath=server.mappath(Dbpath)
			end if
			backpath=server.mappath(backpath)
			
			Set Fso=server.createobject("scripting.filesystemobject")
			if fso.fileexists(dbpath) then  					
				fso.copyfile Dbpath,Backpath
				response.write "<div align=center><font color=#FF0000>�ɹ��ָ����ݣ�</font> </div>"
			else
				response.write "<div align=center><font color=#FF0000>����Ŀ¼�²������ı����ļ���</font> </div>"	
			end if
		else
			call RestoreData()
		end if	
case "SpaceSize"		'ϵͳ�ռ�ռ��
		call SpaceSize()	
end select
%>
<!--#include file=Admin_Bottom.asp-->
<%'====================ϵͳ�ռ�ռ��=======================
sub SpaceSize()
On error resume next
%>
<table cellpadding="0" cellspacing="0" border="1" width="100%" style="border-collapse: collapse" bordercolor="#C0C0C0" width="100%" id="AutoNumber1">
<tr class="TDtop">
	<td height="25" colspan="2" align="center" > <strong>ϵͳ�ռ�ռ�����</strong></td>
</tr>
<%
	If ObjInstalled=false Then
		Response.Write "<tr align=center><td height=30 colspan=3> <div align=center><b><font color=red>��ķ�������֧�� FSO(Scripting.FileSystemObject)! ����ʹ�ñ�����</font></b> </div></td></tr>"
	else
		l_path=left(Request.ServerVariables("PATH_INFO"),instrrev(Request.ServerVariables("PATH_INFO"),"/"))
%>
<tr>
	<td width="22%" height="25">&nbsp;�����ݿ⡿<img src="images/mdb.gif" height=22>��</td>
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
	<td>&nbsp;��<%=SubF.name%>����</td>
	<td>&nbsp;<img src="images/Admin_db_bar.gif" width=<%=drawbar(SubF.name)%> height=10><%showSpaceinfo(SubF.name)%></td>
</tr>
		<%next
		set fsoBrowse=nothing
		set theFolder=nothing
		set theSubFolders=nothing
		%>
<tr>
	<td>&nbsp;ϵͳ�ռ��ܼƣ�</td>
	<td>&nbsp;<img src="images/Admin_db_bar.gif" width=400 height=10>&nbsp;<%showspecialspaceinfo("All")%></td>
</tr> 
<%end if%>
</table>
<%
end sub
%>
<%'====================�ָ����ݿ�=========================
sub RestoreData()
If IsSqlDataBase = 1 Then
	SQLUserReadme()
	Exit Sub
End If
%>
<table cellpadding="0" cellspacing="0" border="1" width="100%" style="border-collapse: collapse" bordercolor="#C0C0C0" width="100%" id="AutoNumber1">
<form method="post" action="<%=ThisNameFile%>?action=RestoreData&act=Restore">		
<tr class="TDtop"> 
	<td height=25 colspan="2" align="center" > &nbsp;&nbsp;<B>�ָ�����</B> </td>
</tr>
<tr> 
	<td width="200" height="30" align="right">�������ݿ�·������ԣ���</td>
	<td height="30">&nbsp;<input type=text size=40 name=DBpath value="DataBackup/News30000.asa"></td>
</tr>
<tr align="center"> 
	<td height="30" colspan="2"><font color="#FF0000">��д�������ݿ��ļ����滻��News30000.asa����������������ע��·�����ܸı䣩�����Բο�������ʾ�����ݿ�����</font></td>
</tr>
<tr> 
	<td width="200" height="30" align="right">��ǰ���ݿ�·������ԣ���</td>
	<td>&nbsp;<input type=text size=40 name=backpath value="<%=db%>"></td>
</tr>
<tr align="center"> 
	<td height="30" colspan="2">��д����ǰʹ�õ����ݿ�·�����粻�븲�ǵ�ǰ�ļ���������������ע��·���Ƿ���ȷ��</td>
</tr>				
<tr align="center"> 
	<td height=30 colspan="2">
		<input type=submit value=" &nbsp;�ָ�����&nbsp;" <%If ObjInstalled=false Then response.Write "disabled"%> style="cursor: hand;background-color: #cccccc;">
	</td>
</tr>
<%
	If ObjInstalled=false Then
		Response.Write "<tr align=center><td height=30 colspan=2> <div align=center><b><font color=red>��ķ�������֧�� FSO(Scripting.FileSystemObject)! ����ʹ�ñ�����</font></b> </div></td></tr>"
	end if
%>	
</form>
</table>
<%
end sub

'====================�������ݿ�=========================
sub BackupData()
If IsSqlDataBase = 1 Then
	SQLUserReadme()
	Exit Sub
End If
%>
<table cellpadding="0" cellspacing="0" border="1" width="100%" style="border-collapse: collapse" bordercolor="#C0C0C0" width="100%" id="AutoNumber1">
<form method="post" action="<%=ThisNameFile%>?action=BackupData&act=Backup">
<tr class="TDtop"> 
	<td height=25  align="center"> <strong>�������ݿ�</strong></td>
</tr>
<tr> 
	<td height=100 >
&nbsp;&nbsp;��ǰ���ݿ�·��(���·��)��<input type=text size=40 name=DBpath value="<%=db%>">
&nbsp;���·��<BR>
&nbsp;&nbsp;�������ݿ�Ŀ¼(���·��)��<input type=text size=40 name=bkfolder value="Databackup"> &nbsp;��Ŀ¼�����ڣ������Զ�����<BR>
&nbsp;&nbsp;�������ݿ�����(��д����)��<input type=text size=40 name=bkDBname value=<%=date%>.asa> &nbsp;�籸��Ŀ¼�и��ļ��������ǣ���û�У����Զ�����
	</td>
</tr>
<tr align="center"> 
	<td height="30" > 
		<input type=submit value=" &nbsp;��ʼ����&nbsp;" <%If ObjInstalled=false Then response.Write "disabled"%> style="cursor: hand;background-color: #cccccc;">
	</td>
</tr>
<%
	If ObjInstalled=false Then
		Response.Write "<tr align=center><td height=30 > <div align=center><b><font color=red>��ķ�������֧�� FSO(Scripting.FileSystemObject)! ����ʹ�ñ�����</font></b> </div></td></tr>"
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
			response.write "<br><div align=center><font color=#FF0000>�������ݿ�ɹ��������ݵ����ݿ�·��Ϊ" &bkfolder& "\"& bkdbname & "</font> </div>"
		Else
			response.write "<br><div align=center><font color=#FF0000>�Ҳ���������Ҫ���ݵ��ļ���</font> </div>"
		End if
end sub
'------------------���ĳһĿ¼�Ƿ����-------------------
Function CheckDir(FolderPath)
	folderpath=Server.MapPath(".")&"\"&folderpath
    Set fso1 = CreateObject("Scripting.FileSystemObject")
    If fso1.FolderExists(FolderPath) then
       '����
       CheckDir = True
    Else
       '������
       CheckDir = False
    End if
    Set fso1 = nothing
End Function
'-------------����ָ����������Ŀ¼-----------------------
Function MakeNewsDir(foldername)
	dim f
    Set fso1 = CreateObject("Scripting.FileSystemObject")
        Set f = fso1.CreateFolder(foldername)
        MakeNewsDir = True
    Set fso1 = nothing
End Function


'====================ѹ�����ݿ� =========================
sub CompressData()

If IsSqlDataBase = 1 Then
	SQLUserReadme()
	Exit Sub
End If
%>
<table cellpadding="0" cellspacing="0" border="1" width="100%" style="border-collapse: collapse" bordercolor="#C0C0C0" width="100%" id="AutoNumber1">
<form action="<%=ThisNameFile%>?action=CompressData" method="post">
<tr class="TDtop"> 
	<td height=25 align="center" class="forumrow"><b>���ݿ�����ѹ��</b></td>
</tr>
<tr> 
	<td height=25 align="center" class="forumrow"><b>ע�⣺</b> �������ݿ��������·��,�����������ݿ����ƣ�����ʹ�������ݿⲻ��ѹ������ѡ�񱸷����ݿ����ѹ��������</td>
</tr>
<tr> 
	<td height="25" align="center" class="forumrow"><font color="#FF6600"><b>ע�⣺</b></font><font color="#FF0000">ѹ��ǰ��ǿ�ҽ����ȱ������ݿ⣬���ⷢ���������</font></td>
</tr>
<tr> 
	<td height="25" align="center" class="forumrow">ѹ�����ݿ⣺<input name="dbpath" type="text" value=<%=db%> size="40"></td>
</tr>
<tr> 
	<td height="25" align="center" class="forumrow"><input type="checkbox" name="boolIs97" value="True">���ʹ�� Access 97 ���ݿ���ѡ�� (Ĭ��Ϊ Access 2000 ���ݿ�)</td>
</tr>
<tr>
	<td height="30" align="center" class="forumrow"><input type="submit" value=" &nbsp;��ʼѹ��&nbsp;" <%If ObjInstalled=false Then response.Write "disabled"%> style="cursor: hand;background-color: #cccccc;"></td>
</tr>
<%
	If ObjInstalled=false Then
		Response.Write "<tr align=center><td height=30 colspan=3> <div align=center><b><font color=red>��ķ�������֧�� FSO(Scripting.FileSystemObject)! ����ʹ�ñ�����</font></b> </div></td></tr>"
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

'=====================ѹ������=========================
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

	CompactDB = "<br><div align=center><font color=#FF0000>������ݿ�: " & dbpath & "  �Ѿ�ѹ���ɹ�!</font> </div>" & vbCrLf

Else
	CompactDB = "<br><div align=center><font color=#FF0000>���ݿ����ƻ�·������ȷ. ������!</font> </div>" & vbCrLf
End If

End Function


'=====================ϵͳ�ռ����=========================
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
 		showsize=size & "�ֽ�" 
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