<!-- #include file="conn.asp" -->
<!-- #include file="inc/const.asp" -->


<title><%=Forum_info(0)%>--����ҳ��</title>
<%=Eval(""""&gettemplate("Forum_css",4)&"""")%>

<BODY <%=Forum_body(11)%>>
<%
	if not master or instr(session("flag"),"28")=0 then
		Errmsg=Errmsg+"<br>"+"<li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��½</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
		call dvbbs_error()
	else
		call main()
	end if

	sub main()
%>
<table cellpadding=0 cellspacing=0 border=0 width=<%=Forum_body(12)%> bgcolor=<%=Forum_body(0)%> align=center>
  <tr>
    <td height="30"> 
      <table cellpadding=3 cellspacing=1 border=0 width=100%>
        <tr bgcolor='<%=Forum_body(2)%>'>
        <td><font color="<%=Forum_body(6)%>">��ӭ<b><%=membername%></b>�������ҳ��</font>
        </td>
        </tr>
            <tr bgcolor=<%=Forum_body(4)%>>              
          <td width="100%" valign=top><font color="<%=Forum_body(7)%>">
<%
dim objFSO
dim fdata
dim objCountFile
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
if request("save")="" then
		Set objCountFile = objFSO.OpenTextFile(Server.MapPath("forum_css.asp"),1,True)
		If Not objCountFile.AtEndOfStream Then fdata = objCountFile.ReadAll
	else
		fdata=request("fdata")
		Set objCountFile=objFSO.CreateTextFile(Server.MapPath("forum_css.asp"),True)
		objCountFile.Write fdata
	end if
	objCountFile.Close
	Set objCountFile=Nothing
	Set objFSO = Nothing
%> 
<form method=post>
            <table width="100%" border="0" cellspacing="3" cellpadding="0">
              <tr> 
                <td width="3%" height="23">&nbsp;</td>
                <td width="97%" height="23"><font color="<%=Forum_body(7)%>">��̳ģ��༭��<br>
                  &nbsp;&nbsp;&nbsp;&nbsp;ע�⣺�ļ������������㰲װĿ¼�µ�<font color="<%=Forum_body(8)%>">FORUM.CSS</font>�������Ŀռ䲻֧��<font color="<%=Forum_body(8)%>">FSO</font>����ֱ�ӱ༭���ļ����������CSS�����Բ��˽⣬�벻Ҫ����༭�������д���<%%>�Ĳ��ֲ�Ҫ���б༭</font></td>
              </tr>
              <tr> 
                <td width="3%">&nbsp; </td>
                <td width="97%"> 
                  <textarea name="fdata" cols="80" rows="20"><%=fdata%></textarea>
                </td>
              </tr>
              <tr> 
                <td width="3%">&nbsp;</td>
                <td width="97%"><font color="<%=Forum_body(7)%>">
                  <input type="reset" name="Reset" value="�����޸�">
                  <input type="submit" name="save" value="�ύ�޸�"> ��ǰ�ļ�·����<font color=<%=Forum_body(8)%>><b><%=Server.MapPath("forum.css")%></b></font></font>
                </td>
              </tr>
            </table></form>
            <p>&nbsp;</p></font>
            </td>
            </tr>
        </table>
        </td>
    </tr>
</table>
<%

end sub%>