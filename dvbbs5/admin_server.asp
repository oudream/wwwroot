<!-- #include file="conn.asp" -->
<!-- #include file="inc/const.asp" -->


<title><%=Forum_info(0)%>--����ҳ��</title>
<%=Eval(""""&gettemplate("Forum_css",4)&"""")%>
<meta NAME=GENERATOR Content=""Microsoft FrontPage 3.0"" CHARSET=GB2312>
<BODY <%=Forum_body(11)%>>
<%
	if not master or instr(session("flag"),"210")=0 then
		Errmsg=Errmsg+"<br>"+"<li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��½</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
		call dvbbs_error()
	else

	Dim theInstalledObjects(17)
    theInstalledObjects(0) = "MSWC.AdRotator"
    theInstalledObjects(1) = "MSWC.BrowserType"
    theInstalledObjects(2) = "MSWC.NextLink"
    theInstalledObjects(3) = "MSWC.Tools"
    theInstalledObjects(4) = "MSWC.Status"
    theInstalledObjects(5) = "MSWC.Counters"
    theInstalledObjects(6) = "IISSample.ContentRotator"
    theInstalledObjects(7) = "IISSample.PageCounter"
    theInstalledObjects(8) = "MSWC.PermissionChecker"
    theInstalledObjects(9) = "Scripting.FileSystemObject"
    theInstalledObjects(10) = "adodb.connection"
    
    theInstalledObjects(11) = "SoftArtisans.FileUp"
    theInstalledObjects(12) = "SoftArtisans.FileManager"
    theInstalledObjects(13) = "JMail.SMTPMail"
    theInstalledObjects(14) = "CDONTS.NewMail"
    theInstalledObjects(15) = "Persits.MailSender"
    theInstalledObjects(16) = "LyfUpload.UploadFile"
    theInstalledObjects(17) = "Persits.Upload.1"

		call main()
	end if

	sub main()
%>
<table cellpadding=0 cellspacing=0 border=0 width=<%=Forum_body(12)%> bgcolor=<%=Forum_body(0)%> align=center>
  <tr>
    <td>
      <table cellpadding=3 cellspacing=1 border=0 width=100%>
        <tr bgcolor='<%=Forum_body(2)%>'>
        <td align=center colspan="2"><font color="<%=Forum_body(6)%>">��ӭ<b><%=membername%></b>�������ҳ��</font>
        </td>
        </tr>
            <tr bgcolor=<%=Forum_body(4)%>>
              <td width="100%" valign=top><font color="<%=Forum_body(7)%>"> <%call servervar()%></font></td>
            </tr>
        </table>
        </td>
    </tr>
</table>
<%
end sub

sub servervar()
%>
              <table width="100%" border="0" cellspacing="3" cellpadding="0">
                <tr> 
                  <td colspan="2" bgcolor=<%=Forum_body(3)%> height=20> 
                    <font color="<%=Forum_body(7)%>"><b>�������йصı���</b></font>
                  </td>
                </tr>
                <tr> 
                  <td width="30%" valign=top><font color="<%=Forum_body(7)%>">��ʾ�ͻ�����������HTTP����</font></td>
                  <td width="70%"><font color="<%=Forum_body(7)%>"><%=request.ServerVariables("All_Http")%></font></td>
                </tr>
                <tr> 
                  <td width="30%" valign=top><font color="<%=Forum_body(7)%>">��ȡISAPIDLL��metabase·��</font></td>
                  <td width="70%"><font color="<%=Forum_body(7)%>"><%=request.ServerVariables("APPL_MD_PATH")%></font></td>
                </tr>
                <tr> 
                  <td width="30%" valign=top><font color="<%=Forum_body(7)%>">��ʾվ������·��</font></td>
                  <td width="70%"><font color="<%=Forum_body(7)%>"><%=request.ServerVariables("APPL_PHYSICAL_PATH")%></font></td>
                </tr>
                <tr> 
                  <td width="30%" valign=top><font color="<%=Forum_body(7)%>">·����Ϣ</font></td>
                  <td width="70%"><font color="<%=Forum_body(7)%>"><%=request.ServerVariables("PATH_INFO")%></font></td>
                </tr>
                <tr> 
                  <td width="30%" valign=top><font color="<%=Forum_body(7)%>">��ʾ�������IP��ַ</font></td>
                  <td width="70%"><font color="<%=Forum_body(7)%>"><%=request.ServerVariables("REMOTE_ADDR")%></font></td>
                </tr>
                <tr> 
                  <td width="30%" valign=top><font color="<%=Forum_body(7)%>">������IP��ַ</font></td>
                  <td width="70%"><font color="<%=Forum_body(7)%>"><%=Request.ServerVariables("LOCAL_ADDR")%></font></td>
                </tr>
                <tr> 
                  <td width="30%" valign=top><font color="<%=Forum_body(7)%>">��ʾִ��SCRIPT������·��</font></td>
                  <td width="70%"><font color="<%=Forum_body(7)%>"><%=request.ServerVariables("SCRIPT_NAME")%></font></td>
                </tr>
                <tr> 
                  <td width="30%" valign=top><font color="<%=Forum_body(7)%>">���ط���������������DNS��������IP��ַ</font></td>
                  <td width="70%"><font color="<%=Forum_body(7)%>"><%=request.ServerVariables("SERVER_NAME")%></font></td>
                </tr>
                <tr> 
                  <td width="30%" valign=top><font color="<%=Forum_body(7)%>">���ط�������������Ķ˿�</font></td>
                  <td width="70%"><font color="<%=Forum_body(7)%>"><%=request.ServerVariables("SERVER_PORT")%></font></td>
                </tr>
                <tr> 
                  <td width="30%" valign=top><font color="<%=Forum_body(7)%>">Э������ƺͰ汾</font></td>
                  <td width="70%"><font color="<%=Forum_body(7)%>"><%=request.ServerVariables("SERVER_PROTOCOL")%></font></td>
                </tr>
                <tr> 
                  <td width="30%" valign=top><font color="<%=Forum_body(7)%>">�����������ƺͰ汾</font></td>
                  <td width="70%"><font color="<%=Forum_body(7)%>"><%=request.ServerVariables("SERVER_SOFTWARE")%></font></td>
                </tr>
                <tr> 
                  <td width="30%" valign=top><font color="<%=Forum_body(7)%>">����������ϵͳ</font></td>
                  <td width="70%"><font color="<%=Forum_body(7)%>"><%=Request.ServerVariables("OS")%></font></td>
                </tr>
                <tr> 
                  <td width="30%" valign=top><font color="<%=Forum_body(7)%>">�ű���ʱʱ��</font></td>
                  <td width="70%"><font color="<%=Forum_body(7)%>"><%=Server.ScriptTimeout%> </font>��</td>
                </tr>
                <tr> 
                  <td width="30%" valign=top><font color="<%=Forum_body(7)%>">������CPU����</font></td>
                  <td width="70%"><font color="<%=Forum_body(7)%>"><%=Request.ServerVariables("NUMBER_OF_PROCESSORS")%> ��</font></td>
                </tr>
                <tr> 
                  <td width="30%" valign=top><font color="<%=Forum_body(7)%>">��������������</font></td>
                  <td width="70%"><font color="<%=Forum_body(7)%>"><%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></font></td>
                </tr>
                <tr> 
                  <td colspan="2" bgcolor=<%=Forum_body(3)%> height=20> <font color="<%=Forum_body(7)%>">
                    <b>���֧�����</b></font>
                  </td>
                </tr>
                <tr> 
                  <td colspan="2" height=20> <font color="<%=Forum_body(7)%>">
<%
Dim strClass
    strClass = Trim(Request.Form("classname"))
    If "" <> strClass then
    Response.Write "<br>��ָ��������ļ������"
      If Not IsObjInstalled(strClass) then 
        Response.Write "<br><font color=red>���ź����÷�������֧��" & strclass & "�����</font>"
      Else
        Response.Write "<br><font color=green>��ϲ���÷�����֧��" & strclass & "�����</font>"
      End If
      Response.Write "<br>"
    end if
%></font>
                  </td>
                </tr>
                <tr> 
                  <td colspan="2" height=20> <font color="<%=Forum_body(7)%>">
                    <b>����IIS�Դ����</b></font>
                  </td>
                </tr>
                <tr> 
                  <td colspan="2" height=20> <font color="<%=Forum_body(7)%>">
	<table border=0 width="95%" cellspacing=1 cellpadding=0>
	<tr height=22 align=center>
	<td width="70%"><font color="<%=Forum_body(7)%>">�� �� �� ��</font></td><td width="15%"><font color="<%=Forum_body(7)%>">֧ ��</font></td><td width="15%"><font color="<%=Forum_body(7)%>">��֧��</font></td>
	</tr>
    <%
    dim i
    For i=0 to 10
      Response.Write "<TR align=center height=22><TD align=left>&nbsp;" & theInstalledObjects(i) & "<font color=#888888>&nbsp;"
	  select case i
		case 9
		Response.Write "(FSO �ı��ļ���д)"
		case 10
		Response.Write "(ACCESS ���ݿ�)"
	  end select
	  Response.Write "</font></td>"
      If Not IsObjInstalled(theInstalledObjects(i)) Then 
        Response.Write "<td></td><td><font color=red><b>��</b></font></td>"
      Else
        Response.Write "<td><b>��</b></td><td></td>"
      End If
      Response.Write "</TR>" & vbCrLf
    Next
    %>
	</table>
                  </td>
                </tr>
                <tr> 
                  <td colspan="2" height=20> 
                    <font color="<%=Forum_body(7)%>"><b>���������������</font></b>
                  </td>
                </tr>
                <tr> 
                  <td colspan="2" height=20> 
	<table border=0 width="95%" cellspacing=1 cellpadding=0>
	<tr height=22 align=center>
	<td width="70%"><font color="<%=Forum_body(7)%>">�� �� �� ��</font></td><td width="15%"><font color="<%=Forum_body(7)%>">֧ ��</font></td><td width="15%"><font color="<%=Forum_body(7)%>">��֧��</font></td>
	</tr>
    <%

    For i=11 to UBound(theInstalledObjects)
      Response.Write "<TR align=center height=18><TD align=left>&nbsp;" & theInstalledObjects(i) & "<font color=#888888>&nbsp;"
	  select case i
		case 11
		Response.Write "(SA-FileUp �ļ��ϴ�)"
		case 12
		Response.Write "(SA-FM �ļ�����)"
		case 13
		Response.Write "(JMail �ʼ�����)"
		case 14
		Response.Write "(CDONTS �ʼ����� SMTP Service)"
		case 15
		Response.Write "(ASPEmail �ʼ�����)"
		case 16
		Response.Write "(LyfUpload �ļ��ϴ�)"
		case 17
		Response.Write "(ASPUpload �ļ��ϴ�)"
		
	  end select
	  Response.Write "</font></td>"
      If Not IsObjInstalled(theInstalledObjects(i)) Then 
        Response.Write "<td></td><td><font color=red><b>��</b></font></td>"
      Else
        Response.Write "<td><b>��</b></td><td></td>"
      End If

      Response.Write "</TR>" & vbCrLf
    Next
    %>
	</table>
                  </td>
                </tr>
                <tr> 
                  <td colspan="2" height=20> <font color="<%=Forum_body(7)%>">
                    <b>�����������֧�������⣺</b>��������������������Ҫ���������ProgId��ClassId��</font>
                  </td>
                </tr>
                <tr> 
                  <td colspan="2" height=20> 
	<table border=0 width="95%" cellspacing=1 cellpadding=0>
<FORM action="admin_server" method=post id=form1 name=form1>
	  <tr>
		<td align=center height=30>
<input class=input type=text value="" name="classname" size=40>
<INPUT type=submit value="ȷ��" id=submit1 name=submit1>
<INPUT type=reset value="����" id=reset1 name=reset1> 
</td>
	  </tr>
</FORM>
	</table>
                  </td>
                </tr>
              </table>
<%
end sub
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
%>