<!--#include file =conn.asp-->
<!-- #include file="inc/const.asp" -->
<title><%=Forum_info(0)%>--����ҳ��</title>
<!--#include file=inc/forum_css.asp-->
<meta NAME=GENERATOR Content="Microsoft FrontPage 3.0" CHARSET=GB2312>
<BODY <%=Forum_body(11)%>>
<%
dim dbpath,bkfolder,bkdbname,fso,fso1
dim backpath
	if not master or session("flag")="" then
		Errmsg=Errmsg+"<br>"+"<li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��½</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
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
	response.write "��������Ҫ�ָ��ɵ����ݿ�ȫ��"	
	else
	Dbpath=server.mappath(Dbpath)
	end if
	backpath=server.mappath(backpath)
	'Response.write Backpath
	Set Fso=server.createobject("scripting.filesystemobject")
	if fso.fileexists(dbpath) then  					
		fso.copyfile Dbpath,Backpath
		response.write "�ɹ��ָ����ݣ�"
	else
		response.write "����Ŀ¼�²������ı����ļ���"	
	end if
else
%>
<table align=center cellspacing=1 cellpadding=1 class=tableborder1>		  							  				
  				<tr>
  					<th height=25 >
   					&nbsp;&nbsp;<B>�ָ���̳����</B>( ��ҪFSO֧�֣�FSO��ذ����뿴΢����վ )
  					</th>
  				</tr>
  				<form method="post" action="ADMIN_RestoreData.asp?action=Restore">
  				<tr>
  					<td height=100 class=tablebody1>
  						&nbsp;&nbsp;�������ݿ�·��(���)��<input type=text size=30 name=DBpath value="DataBackup\dvbbs5_Backup.MDB">&nbsp;&nbsp;<BR>
  						&nbsp;&nbsp;Ŀ�����ݿ�·��(���)��<input type=text size=30 name=backpath value="<%=db%>"><BR>&nbsp;&nbsp;��д����ǰʹ�õ����ݿ�·�����粻�븲�ǵ�ǰ�ļ���������������ע��·���Ƿ���ȷ����Ȼ���޸�conn.asp�ļ������Ŀ���ļ����͵�ǰʹ�����ݿ���һ�µĻ��������޸�conn.asp�ļ�<BR>
						&nbsp;&nbsp;<input type=submit value="�ָ�����"> <br>
  						-----------------------------------------------------------------------------------------<br>
  						&nbsp;&nbsp;��������д����������ݿ�·��ȫ�����������Ĭ�ϱ������ݿ��ļ�ΪDataBackup\dvbbs_Backup.MDB���밴�����ı����ļ������޸ġ�<br>
  						&nbsp;&nbsp;������������������������ķ������ݣ��Ա�֤�������ݰ�ȫ��<br>
  						&nbsp;&nbsp;ע�⣺����·��������������ռ��Ŀ¼�����·��</font>
  					</td>
  				</tr>	
  				</form>
  			</table>
<%
end if

end sub
%>
