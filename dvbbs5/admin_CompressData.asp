<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<title><%=Forum_info(0)%>--����ҳ��</title>
<!--#include file=inc/forum_css.asp-->

<BODY <%=Forum_body(11)%>>
<%
	if not master or session("flag")="" then
		Errmsg=Errmsg+"<br>"+"<li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��½</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
		call dvbbs_error()
	else
		dim tmprs
		dim allarticle
		dim Maxid
		dim topic,username,dateandtime,body
		call main()
		set rs=nothing
		conn.close
		set conn=nothing
	end if

	sub main()
%>
<table class=tableborder1 cellspacing=1 cellpadding=1 align=center>
<form action=Admin_CompressData.asp>
<tr>
<td class=tablebody1 height=25><b>ע�⣺</b><br>�������ݿ��������·��,�����������ݿ����ƣ�����ʹ�������ݿⲻ��ѹ������ѡ�񱸷����ݿ����ѹ�������� </td>
</tr>
<tr>
<td class=tablebody1>ѹ�����ݿ⣺<input type="text" name="dbpath" value=Data\Dvbbs5.MDB>&nbsp;
<input type="submit" value="��ʼѹ��"></td>
</tr>
<tr>
<td class=tablebody1><input type="checkbox" name="boolIs97" value="True">���ʹ�� Access 97 ���ݿ���ѡ��
(Ĭ��Ϊ Access 2000 ���ݿ�)<br><br></td>
</tr>
<form>
</table>
<%
Dim dbpath,boolIs97
dbpath = request("dbpath")
boolIs97 = request("boolIs97")

If dbpath <> "" Then
dbpath = server.mappath(dbpath)
	response.write(CompactDB(dbpath,boolIs97))
End If
	end sub
%>
<%
Const JET_3X = 4

Function CompactDB(dbPath, boolIs97)
Dim fso, Engine, strDBPath
strDBPath = left(dbPath,instrrev(DBPath,"\"))
Set fso = CreateObject("Scripting.FileSystemObject")

If fso.FileExists(dbPath) Then
Set Engine = CreateObject("JRO.JetEngine")

	If boolIs97 = "True" Then
		Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbpath, _
		"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp.mdb;" _
		& "Jet OLEDB:Engine Type=" & JET_3X
	Else
		Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbpath, _
		"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp.mdb"
	End If

fso.CopyFile strDBPath & "temp.mdb",dbpath
fso.DeleteFile(strDBPath & "temp.mdb")
Set fso = nothing
Set Engine = nothing

	CompactDB = "������ݿ�, " & dbpath & ", �Ѿ�ѹ���ɹ�!" & vbCrLf

Else
	CompactDB = "���ݿ����ƻ�·������ȷ. ������!" & vbCrLf
End If

End Function
%>
