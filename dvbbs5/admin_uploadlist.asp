<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!--#include file=inc/char.asp-->

<title><%=Forum_info(0)%>--����ҳ��</title>
<%=Eval(""""&gettemplate("Forum_css",4)&"""")%>

<BODY <%=Forum_body(11)%>>
<%
	dim path
	if not master or instr(session("flag"),"55")=0 then
		Errmsg=Errmsg+"<br>"+"<li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��½</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
		call dvbbs_error()
	else
		call main()
	end if

	sub main()
%>
<table cellpadding=0 cellspacing=0 border=0 width=<%=Forum_body(12)%> bgcolor=<%=Forum_body(0)%> align=center>
  <tr>
    <td>
      <table cellpadding=3 cellspacing=1 border=0 width=100%>
        <tr bgcolor='<%=Forum_body(2)%>'>
        <td><font color="<%=Forum_body(6)%>">��ӭ<b><%=membername%></b>�������ҳ��</font>
        </td>
        </tr>
            <tr bgcolor=<%=Forum_body(4)%>>
              <td width="100%" valign=top><p>
              <table width="100%" border="0" cellspacing="3" cellpadding="0">
                <tr> 
                  <td bgcolor=<%=Forum_body(3)%>> <font color="<%=Forum_body(7)%>">
                    <p><b>����ϴ��ļ�</b>�������ܱ��������֧��FSOȨ�޷���ʹ�ã�FSOʹ�ð��������΢����վ���ļ�Ŀ¼Ϊuploadimages���������������֧��FSO���ֶ�����</p></font>
                  </td>
                </tr>
����������������������<tr> 
<form method="POST" action="?action=pathname">
                  <td >Ҫ�鿴���ļ��У�
<input type="text" name="path" size="20"><input type="submit" value="�ύ" name="B1"  style="background-color: #FFCC99; border-style: solid; border-color: #FEF4D3">
<input type="reset" value="ȫ����д" name="B2"  style="background-color: #FFCC99; border-style: solid; border-color: #FEF4D3">
       ��<font color=red>����д��ȷ���ļ�������·��</font>��</font></td></form>
                </tr>
<%
if request("path")<>"" then
path=request("path")
else 
path="uploadimages"
end if
%>
               <tr> 
                  <td bgcolor=<%=Forum_body(2)%>>
                   <font color="<%=Forum_body(6)%>"> <b>��ǰ�����<%=path%>���ļ�Ŀ¼�������ļ�</b>��
                  </td>
                </tr>
                <tr> 
                  <td>	<font color="<%=Forum_body(7)%>">
<%
'Created by shatan Jan 8
dim objFSO
dim uploadfolder
dim uploadfiles
dim upname
dim upfilename
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
if request("filename")<>"" then
objFSO.DeleteFile(Server.MapPath(path&"\"&request("filename")))
end if
Set uploadFolder=objFSO.GetFolder(Server.MapPath(path & "\"))
Set uploadFiles=uploadFolder.Files
For Each Upname In uploadFiles
upfilename=path& "/" &upname.name
response.write "<a href="""&upfilename&""">"&upfilename&"</a> <a href='?path="&path&"&filename="&upname.name&"'>ɾ��</a><br>"
next
set uploadFolder=nothing
set uploadFiles=nothing
%></font>
		</td>
            </tr>
        </table>
        </td>
    </tr>
</table>
<%
	end sub
%>