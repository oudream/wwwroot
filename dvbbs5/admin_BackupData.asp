<!--#include file =conn.asp-->
<!-- #include file="inc/const.asp" -->
<title><%=Forum_info(0)%>--����ҳ��</title>
<!--#include file=inc/forum_css.asp-->
<meta NAME=GENERATOR Content=""Microsoft FrontPage 3.0"" CHARSET=GB2312>
<BODY  <%=Forum_body(11)%>>
<%
dim dbpath,bkfolder,bkdbname,fso,fso1
	if not master or session("flag")="" then
		Errmsg=Errmsg+"<br>"+"<li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��½</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
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
  					&nbsp;&nbsp;<B>������̳����</B>( ��ҪFSO֧�֣�FSO��ذ����뿴΢����վ )
  					</th>
  				</tr>
  				<form method="post" action="ADMIN_BackupData.asp?action=Backup">
  				<tr>
  					<td height=100 class=tablebody1>
  						&nbsp;&nbsp;
						��ǰ���ݿ�·��(���·��)��<input type=text size=15 name=DBpath value="<%=db%>"><BR>&nbsp;&nbsp;
						�������ݿ�Ŀ¼(���·��)��<input type=text size=15 name=bkfolder value=Databackup>&nbsp;��Ŀ¼�����ڣ������Զ�����<BR>&nbsp;&nbsp;
						�������ݿ�����(��д����)��<input type=text size=15 name=bkDBname value=dvbbs5.MDB>&nbsp;�籸��Ŀ¼�и��ļ��������ǣ���û�У����Զ�����<BR>
						&nbsp;&nbsp;<input type=submit value="ȷ��"><br>
  						-----------------------------------------------------------------------------------------<br>
  						&nbsp;&nbsp;��������д����������ݿ�·��ȫ�����������Ĭ�����ݿ��ļ�ΪData\dvbbs5.MDB��<B>��һ��������Ĭ�����������������ݿ�</B><br>
  						&nbsp;&nbsp;������������������������ķ������ݣ��Ա�֤�������ݰ�ȫ��<br>
  						&nbsp;&nbsp;ע�⣺����·��������������ռ��Ŀ¼�����·��				</font>
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
			response.write "�������ݿ�ɹ��������ݵ����ݿ�·��Ϊ" &bkfolder& "\"& bkdbname
		Else
			response.write "�Ҳ���������Ҫ���ݵ��ļ���"
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
'-------------����ָ����������Ŀ¼---------
Function MakeNewsDir(foldername)
	dim f
    Set fso1 = CreateObject("Scripting.FileSystemObject")
        Set f = fso1.CreateFolder(foldername)
        MakeNewsDir = True
    Set fso1 = nothing
End Function
%>