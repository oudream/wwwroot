<!-- #include file="conn.asp" -->
<!-- #include file="inc/const.asp" -->


<title><%=Forum_info(0)%>--管理页面</title>
<%=Eval(""""&gettemplate("Forum_css",4)&"""")%>
<meta NAME=GENERATOR Content=""Microsoft FrontPage 3.0"" CHARSET=GB2312>
<BODY <%=Forum_body(11)%>>
<%
	if not master or instr(session("flag"),"210")=0 then
		Errmsg=Errmsg+"<br>"+"<li>本页面为管理员专用，请<a href=admin_index.asp target=_top>登陆</a>后进入。<br><li>您没有管理本页面的权限。"
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
        <td align=center colspan="2"><font color="<%=Forum_body(6)%>">欢迎<b><%=membername%></b>进入管理页面</font>
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
                    <font color="<%=Forum_body(7)%>"><b>服务器有关的变量</b></font>
                  </td>
                </tr>
                <tr> 
                  <td width="30%" valign=top><font color="<%=Forum_body(7)%>">显示客户发出的所有HTTP标题</font></td>
                  <td width="70%"><font color="<%=Forum_body(7)%>"><%=request.ServerVariables("All_Http")%></font></td>
                </tr>
                <tr> 
                  <td width="30%" valign=top><font color="<%=Forum_body(7)%>">检取ISAPIDLL的metabase路径</font></td>
                  <td width="70%"><font color="<%=Forum_body(7)%>"><%=request.ServerVariables("APPL_MD_PATH")%></font></td>
                </tr>
                <tr> 
                  <td width="30%" valign=top><font color="<%=Forum_body(7)%>">显示站点物理路径</font></td>
                  <td width="70%"><font color="<%=Forum_body(7)%>"><%=request.ServerVariables("APPL_PHYSICAL_PATH")%></font></td>
                </tr>
                <tr> 
                  <td width="30%" valign=top><font color="<%=Forum_body(7)%>">路径信息</font></td>
                  <td width="70%"><font color="<%=Forum_body(7)%>"><%=request.ServerVariables("PATH_INFO")%></font></td>
                </tr>
                <tr> 
                  <td width="30%" valign=top><font color="<%=Forum_body(7)%>">显示请求机器IP地址</font></td>
                  <td width="70%"><font color="<%=Forum_body(7)%>"><%=request.ServerVariables("REMOTE_ADDR")%></font></td>
                </tr>
                <tr> 
                  <td width="30%" valign=top><font color="<%=Forum_body(7)%>">服务器IP地址</font></td>
                  <td width="70%"><font color="<%=Forum_body(7)%>"><%=Request.ServerVariables("LOCAL_ADDR")%></font></td>
                </tr>
                <tr> 
                  <td width="30%" valign=top><font color="<%=Forum_body(7)%>">显示执行SCRIPT的虚拟路径</font></td>
                  <td width="70%"><font color="<%=Forum_body(7)%>"><%=request.ServerVariables("SCRIPT_NAME")%></font></td>
                </tr>
                <tr> 
                  <td width="30%" valign=top><font color="<%=Forum_body(7)%>">返回服务器的主机名，DNS别名，或IP地址</font></td>
                  <td width="70%"><font color="<%=Forum_body(7)%>"><%=request.ServerVariables("SERVER_NAME")%></font></td>
                </tr>
                <tr> 
                  <td width="30%" valign=top><font color="<%=Forum_body(7)%>">返回服务器处理请求的端口</font></td>
                  <td width="70%"><font color="<%=Forum_body(7)%>"><%=request.ServerVariables("SERVER_PORT")%></font></td>
                </tr>
                <tr> 
                  <td width="30%" valign=top><font color="<%=Forum_body(7)%>">协议的名称和版本</font></td>
                  <td width="70%"><font color="<%=Forum_body(7)%>"><%=request.ServerVariables("SERVER_PROTOCOL")%></font></td>
                </tr>
                <tr> 
                  <td width="30%" valign=top><font color="<%=Forum_body(7)%>">服务器的名称和版本</font></td>
                  <td width="70%"><font color="<%=Forum_body(7)%>"><%=request.ServerVariables("SERVER_SOFTWARE")%></font></td>
                </tr>
                <tr> 
                  <td width="30%" valign=top><font color="<%=Forum_body(7)%>">服务器操作系统</font></td>
                  <td width="70%"><font color="<%=Forum_body(7)%>"><%=Request.ServerVariables("OS")%></font></td>
                </tr>
                <tr> 
                  <td width="30%" valign=top><font color="<%=Forum_body(7)%>">脚本超时时间</font></td>
                  <td width="70%"><font color="<%=Forum_body(7)%>"><%=Server.ScriptTimeout%> </font>秒</td>
                </tr>
                <tr> 
                  <td width="30%" valign=top><font color="<%=Forum_body(7)%>">服务器CPU数量</font></td>
                  <td width="70%"><font color="<%=Forum_body(7)%>"><%=Request.ServerVariables("NUMBER_OF_PROCESSORS")%> 个</font></td>
                </tr>
                <tr> 
                  <td width="30%" valign=top><font color="<%=Forum_body(7)%>">服务器解译引擎</font></td>
                  <td width="70%"><font color="<%=Forum_body(7)%>"><%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></font></td>
                </tr>
                <tr> 
                  <td colspan="2" bgcolor=<%=Forum_body(3)%> height=20> <font color="<%=Forum_body(7)%>">
                    <b>组件支持情况</b></font>
                  </td>
                </tr>
                <tr> 
                  <td colspan="2" height=20> <font color="<%=Forum_body(7)%>">
<%
Dim strClass
    strClass = Trim(Request.Form("classname"))
    If "" <> strClass then
    Response.Write "<br>您指定的组件的检查结果："
      If Not IsObjInstalled(strClass) then 
        Response.Write "<br><font color=red>很遗憾，该服务器不支持" & strclass & "组件！</font>"
      Else
        Response.Write "<br><font color=green>恭喜！该服务器支持" & strclass & "组件。</font>"
      End If
      Response.Write "<br>"
    end if
%></font>
                  </td>
                </tr>
                <tr> 
                  <td colspan="2" height=20> <font color="<%=Forum_body(7)%>">
                    <b>－－IIS自带组件</b></font>
                  </td>
                </tr>
                <tr> 
                  <td colspan="2" height=20> <font color="<%=Forum_body(7)%>">
	<table border=0 width="95%" cellspacing=1 cellpadding=0>
	<tr height=22 align=center>
	<td width="70%"><font color="<%=Forum_body(7)%>">组 件 名 称</font></td><td width="15%"><font color="<%=Forum_body(7)%>">支 持</font></td><td width="15%"><font color="<%=Forum_body(7)%>">不支持</font></td>
	</tr>
    <%
    dim i
    For i=0 to 10
      Response.Write "<TR align=center height=22><TD align=left>&nbsp;" & theInstalledObjects(i) & "<font color=#888888>&nbsp;"
	  select case i
		case 9
		Response.Write "(FSO 文本文件读写)"
		case 10
		Response.Write "(ACCESS 数据库)"
	  end select
	  Response.Write "</font></td>"
      If Not IsObjInstalled(theInstalledObjects(i)) Then 
        Response.Write "<td></td><td><font color=red><b>×</b></font></td>"
      Else
        Response.Write "<td><b>√</b></td><td></td>"
      End If
      Response.Write "</TR>" & vbCrLf
    Next
    %>
	</table>
                  </td>
                </tr>
                <tr> 
                  <td colspan="2" height=20> 
                    <font color="<%=Forum_body(7)%>"><b>－－其他常见组件</font></b>
                  </td>
                </tr>
                <tr> 
                  <td colspan="2" height=20> 
	<table border=0 width="95%" cellspacing=1 cellpadding=0>
	<tr height=22 align=center>
	<td width="70%"><font color="<%=Forum_body(7)%>">组 件 名 称</font></td><td width="15%"><font color="<%=Forum_body(7)%>">支 持</font></td><td width="15%"><font color="<%=Forum_body(7)%>">不支持</font></td>
	</tr>
    <%

    For i=11 to UBound(theInstalledObjects)
      Response.Write "<TR align=center height=18><TD align=left>&nbsp;" & theInstalledObjects(i) & "<font color=#888888>&nbsp;"
	  select case i
		case 11
		Response.Write "(SA-FileUp 文件上传)"
		case 12
		Response.Write "(SA-FM 文件管理)"
		case 13
		Response.Write "(JMail 邮件发送)"
		case 14
		Response.Write "(CDONTS 邮件发送 SMTP Service)"
		case 15
		Response.Write "(ASPEmail 邮件发送)"
		case 16
		Response.Write "(LyfUpload 文件上传)"
		case 17
		Response.Write "(ASPUpload 文件上传)"
		
	  end select
	  Response.Write "</font></td>"
      If Not IsObjInstalled(theInstalledObjects(i)) Then 
        Response.Write "<td></td><td><font color=red><b>×</b></font></td>"
      Else
        Response.Write "<td><b>√</b></td><td></td>"
      End If

      Response.Write "</TR>" & vbCrLf
    Next
    %>
	</table>
                  </td>
                </tr>
                <tr> 
                  <td colspan="2" height=20> <font color="<%=Forum_body(7)%>">
                    <b>－－其他组件支持情况检测：</b>在下面的输入框中输入你要检测的组件的ProgId或ClassId。</font>
                  </td>
                </tr>
                <tr> 
                  <td colspan="2" height=20> 
	<table border=0 width="95%" cellspacing=1 cellpadding=0>
<FORM action="admin_server" method=post id=form1 name=form1>
	  <tr>
		<td align=center height=30>
<input class=input type=text value="" name="classname" size=40>
<INPUT type=submit value="确定" id=submit1 name=submit1>
<INPUT type=reset value="重填" id=reset1 name=reset1> 
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