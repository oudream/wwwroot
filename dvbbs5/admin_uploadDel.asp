<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->


<title><%=Forum_info(0)%>--管理页面</title>
<%=Eval(""""&gettemplate("Forum_css",4)&"""")%>

<BODY <%=Forum_body(11)%>>
<%
	Server.ScriptTimeout=999999
	if instr(session("flag"),"55")=0 then
		Errmsg=Errmsg+"<br>"+"<li>本页面为管理员专用，请<a href=admin_index.asp target=_top>登陆</a>后进入。<br><li>您没有管理本页面的权限。"
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
        <td><font color="<%=Forum_body(6)%>">欢迎<b><%=membername%></b>进入管理页面</font>
        </td>
        </tr>
            <tr bgcolor=<%=Forum_body(4)%>>
              <td width="100%" valign=top><p>
              <table width="100%" border="0" cellspacing="3" cellpadding="0">
                <tr> 
                  <td bgcolor=<%=Forum_body(3)%>> <font color="<%=Forum_body(7)%>">
                    <p><b>清理上传文件</b>：本功能必须服务器支持FSO权限方能使用，FSO使用帮助请浏览微软网站。文件目录为uploadimages，如果您服务器不支持FSO请手动管理。</p></font>
                  </td>
                </tr>
                <tr> 
                  <td>	<font color="<%=Forum_body(7)%>">
<% if request.querystring("del")<>"yes" then
	response.write "<a href=?del=yes>点击开始清理</a>（根据论坛帖数可能需要相当长的时间）"
else
	del()
end if %>
</font>
		</td>
            </tr>
        </table>
        </td>
    </tr>
</table>
<%
end sub


sub del()
	dim objFSO
	dim uploadFolder
	dim uploadFiles
	dim upfilename
	dim upname
	i=0
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	Set uploadFolder=objFSO.GetFolder(Server.MapPath("uploadimages\"))
	Set uploadFiles=uploadFolder.Files
	For Each Upname In uploadFiles
		upfilename="uploadimages/"&upname.name
		sql="select top 1 announceid from bbs1 where body like '%"&upfilename&"%' "
		'response.write sql
		'response.end
		set rs=conn.execute(sql)
		if rs.eof then
			objFSO.DeleteFile(Server.MapPath(upfilename))
			response.write "已删除文件："&upfilename&"<br>"
			i=i+1
		end if
		rs.close
		set rs=nothing
	next
	set uploadFolder=nothing
	set uploadFiles=nothing
	response.write "<br>清理完成，共清理上传文件 "&i&" 个"
end sub
%>

