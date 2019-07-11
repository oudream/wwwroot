<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!--#include file=inc/char.asp-->

<title><%=Forum_info(0)%>--管理页面</title>
<%=Eval(""""&gettemplate("Forum_css",4)&"""")%>

<BODY <%=Forum_body(11)%>>
<%
	dim path
	if not master or instr(session("flag"),"55")=0 then
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
                    <p><b>浏览上传文件</b>：本功能必须服务器支持FSO权限方能使用，FSO使用帮助请浏览微软网站。文件目录为uploadimages，如果您服务器不支持FSO请手动管理。</p></font>
                  </td>
                </tr>
　　　　　　　　　　　<tr> 
<form method="POST" action="?action=pathname">
                  <td >要查看的文件夹：
<input type="text" name="path" size="20"><input type="submit" value="提交" name="B1"  style="background-color: #FFCC99; border-style: solid; border-color: #FEF4D3">
<input type="reset" value="全部重写" name="B2"  style="background-color: #FFCC99; border-style: solid; border-color: #FEF4D3">
       （<font color=red>请填写正确的文件夹名或路径</font>）</font></td></form>
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
                   <font color="<%=Forum_body(6)%>"> <b>当前浏览　<%=path%>　文件目录的所有文件</b>：
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
response.write "<a href="""&upfilename&""">"&upfilename&"</a> <a href='?path="&path&"&filename="&upname.name&"'>删除</a><br>"
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