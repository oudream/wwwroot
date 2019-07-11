<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Config.asp"-->
<%
session("Admin") = ""
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>:::丁书记上传程序 ::: 演示版本</title>
<link href="upstyle.css" rel="stylesheet" type="text/css">
<script src="Images/up.Js"></script>
</head>

<body background="../images/mainbg.gif" leftmargin="0" topmargin="0">
<table width="778" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#003366">
  <tr> 
    <td height="75">&nbsp;</td>
  </tr>
  <tr> 
    <td height="22" bgcolor="#FF9900" class="text">当前位置：:::丁书记上传程序Ver1.2 ::: 演示版本 
      &gt;&gt;&gt; 首页</td>
  </tr>
  <tr>
    <td height="2" bgcolor="#cccccc"></td>
  </tr>
</table>
<table width="778" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="778" height="200">
        <param name="movie" value="FLASH_AD.swf">
        <param name="quality" value="high">
        <embed src="FLASH_AD.swf" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" width="778" height="200"></embed></object></td>
  </tr>
  <tr>
    <td height="1" bgcolor="#CCCCCC"></td>
  </tr>
</table>
<table width="778" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#0153A3">
  <tr> 
    <td height="25"> <%
	set frst = Server.CreateObject("adodb.recordset")
	sql = "select * from info order by uploadtime desc"
	frst.open sql,myconn,1,1
	fcount = frst.recordcount
	if fcount > 0 then 	
		''显示参数
		dim tbbgcolor
		''分页参数
		dim maxperpage,pages,page
		maxperpage = 5
		frst.pagesize = maxperpage
		pages = frst.pagecount
		''页面参数设置
		page = Request.QueryString("page")
		if not isnumeric(page) then page = 1 else page = cint(page)
		if page < 1 then page = 1
		if page > pages then page = pages
		frst.absolutepage = page
		''显示内容
'Set Isize=Server.CreateObject("WinImg.ImgSize")
		for i = 1 to maxperpage
			if frst.eof then exit for
			if i mod 2 = 0 then tbbgcolor = "" else tbbgcolor = "#0066cc"
			fid = frst("id").Value
			ftitle = frst("fileTitle").Value
			fdesc = frst("fileDesc").Value
			ftype = frst("fileType").Value
			fpath = frst("filePath").Value
			fsize = frst("filesize").Value
			fhits = frst("hits").Value
			fuploadtime = frst("uploadTime").Value
'FileNameStr=Server.Mappath(fpath)
'Isize.GetImgSize Cstr(FileNameStr)
	%>
      <table width="100%" border="0" align="center" cellpadding="3" cellspacing="0">
        <tr class="text"> 
          <td width="150"><div align="right">文件名称：</div></td>
          <td><a href="<%=fpath%>" target="_NEWwIN"><%=ftitle%>( 文件大小：<%=fsize%> bytes)</a> </td>
          <td align="right"></td>
        </tr>
        <tr class="text"> 
          <td width="150"><div align="right">文件类型：</div></td>
          <td colspan="2"><%=ftype%></td>
        </tr>
        <tr class="text"> 
          <td width="150"><div align="right">上传日期：</div></td>
          <td colspan="2"><%=fuploadtime%></td>
        </tr>
        <tr class="text"> 
          <td width="150"><div align="right">说明：</div></td>
          <td colspan="2"><%=fdesc%></td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td height="1"></td>
          <td height="1" colspan="2"></td>
        </tr>
      </table>
      <%
		  	frst.movenext
		next
	  else
	  %> 
      <table width="100%" border="0" align="center" cellpadding="3" cellspacing="0">
        <tr class="text"> 
          <td>还没有上传的内容！</td>
        </tr>
      </table>
      <%
	  end if
	  frst.close
	  set frst = nothing
	  myconn.close
	  set myconn = nothing
	  %>
	 <table width="100%" border="0" align="center" cellpadding="3" cellspacing="0">
        <tr class="text"> 
          <td align="center">
<%
		  If Page > 1 Then Response.Write ("<a href='?page=1'>首页</a><a href='?page="& Page - 1 &"'>上一页</a>")
		  If page < pages Then Response.Write ("&nbsp;<a href='?page="& Page + 1 &"'>下一页</a>&nbsp;<a href='?page="& Pages &"'>末页</a>")
		  %></td>
        </tr>
      </table></td>
  </tr>
</table>
<table width="778" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#003366">
  <tr> 
    <td height="25">
        <table width="100%" border="0" align="center" cellpadding="3" cellspacing="0" bgcolor="#006699"><form action="SaveUpload.asp" method="post" enctype="multipart/form-data" name="form1">
          <tr bgcolor="#006699" class="text"> 
            <td width="33%" height="25"> <div align="left"><strong>上传内容</strong>
		</div></td>
            <td><%
		Response.Write "  允许上传的文件类型:<br> "
		Set Fs = Server.CreateObject("Scripting.FileSystemObject")
		For Each str In OKAr
		If Fs.FileExists(Server.MapPath("Images\"& Str &".gif")) Then
		Response.Write "<img src='Images/" & str & ".gif' alt='" & str & "文件'> "
		Else
		Response.Write "<img src='Images/X.gif' alt='" & str & "文件'> "
		End If
		Next
		Set Fs = Nothing
		Response.Write "<br>  允许上传的文件最大:"&Oksize / 1024&"KB"
		%></td>
          </tr>
          <tr class="text"> 
            <td width="33%"><div align="right"><strong>文件名称：</strong></div></td>
            <td><input name="filetitle" type="text" class="TextBoxT" size="25"> 
              <select name="filetype" class="TextBoxT" id="filetype">
                <option value="素材图片">素材图片</option>
                <option value="常用工具">常用工具</option>
                <option value="程序源码">程序源码</option>
              </select> </td>
          </tr>
          <tr class="text"> 
            <td valign="top"><div align="right"><strong>上传的文件：</strong></div></td>
            <td><input name="filedata" type="file" class="TextBoxT" id="filedata" size="27"></td>
          </tr>
          <tr class="text"> 
            <td valign="top"> <div align="right"><strong>文件说明：</strong><br>
              </div></td>
            <td><textarea name="filedesc" cols="36" rows="4" class="TextBoxT" id="filedesc"></textarea></td>
          </tr>
          <tr> 
            <td>&nbsp;</td>
            <td><input type="submit" name="Submit" value="上传内容"></td>
          </tr></form>
        </table>
      </td>
  </tr>
</table>
<table width="778" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#0153A3">
  <tr> 
    <td height="25">&nbsp;</td>
  </tr>
</table>
</body>
</html>
