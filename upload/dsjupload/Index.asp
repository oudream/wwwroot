<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Config.asp"-->
<%
session("Admin") = ""
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>:::������ϴ����� ::: ��ʾ�汾</title>
<link href="upstyle.css" rel="stylesheet" type="text/css">
<script src="Images/up.Js"></script>
</head>

<body background="../images/mainbg.gif" leftmargin="0" topmargin="0">
<table width="778" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#003366">
  <tr> 
    <td height="75">&nbsp;</td>
  </tr>
  <tr> 
    <td height="22" bgcolor="#FF9900" class="text">��ǰλ�ã�:::������ϴ�����Ver1.2 ::: ��ʾ�汾 
      &gt;&gt;&gt; ��ҳ</td>
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
		''��ʾ����
		dim tbbgcolor
		''��ҳ����
		dim maxperpage,pages,page
		maxperpage = 5
		frst.pagesize = maxperpage
		pages = frst.pagecount
		''ҳ���������
		page = Request.QueryString("page")
		if not isnumeric(page) then page = 1 else page = cint(page)
		if page < 1 then page = 1
		if page > pages then page = pages
		frst.absolutepage = page
		''��ʾ����
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
          <td width="150"><div align="right">�ļ����ƣ�</div></td>
          <td><a href="<%=fpath%>" target="_NEWwIN"><%=ftitle%>( �ļ���С��<%=fsize%> bytes)</a> </td>
          <td align="right"></td>
        </tr>
        <tr class="text"> 
          <td width="150"><div align="right">�ļ����ͣ�</div></td>
          <td colspan="2"><%=ftype%></td>
        </tr>
        <tr class="text"> 
          <td width="150"><div align="right">�ϴ����ڣ�</div></td>
          <td colspan="2"><%=fuploadtime%></td>
        </tr>
        <tr class="text"> 
          <td width="150"><div align="right">˵����</div></td>
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
          <td>��û���ϴ������ݣ�</td>
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
		  If Page > 1 Then Response.Write ("<a href='?page=1'>��ҳ</a><a href='?page="& Page - 1 &"'>��һҳ</a>")
		  If page < pages Then Response.Write ("&nbsp;<a href='?page="& Page + 1 &"'>��һҳ</a>&nbsp;<a href='?page="& Pages &"'>ĩҳ</a>")
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
            <td width="33%" height="25"> <div align="left"><strong>�ϴ�����</strong>
		</div></td>
            <td><%
		Response.Write "  �����ϴ����ļ�����:<br> "
		Set Fs = Server.CreateObject("Scripting.FileSystemObject")
		For Each str In OKAr
		If Fs.FileExists(Server.MapPath("Images\"& Str &".gif")) Then
		Response.Write "<img src='Images/" & str & ".gif' alt='" & str & "�ļ�'> "
		Else
		Response.Write "<img src='Images/X.gif' alt='" & str & "�ļ�'> "
		End If
		Next
		Set Fs = Nothing
		Response.Write "<br>  �����ϴ����ļ����:"&Oksize / 1024&"KB"
		%></td>
          </tr>
          <tr class="text"> 
            <td width="33%"><div align="right"><strong>�ļ����ƣ�</strong></div></td>
            <td><input name="filetitle" type="text" class="TextBoxT" size="25"> 
              <select name="filetype" class="TextBoxT" id="filetype">
                <option value="�ز�ͼƬ">�ز�ͼƬ</option>
                <option value="���ù���">���ù���</option>
                <option value="����Դ��">����Դ��</option>
              </select> </td>
          </tr>
          <tr class="text"> 
            <td valign="top"><div align="right"><strong>�ϴ����ļ���</strong></div></td>
            <td><input name="filedata" type="file" class="TextBoxT" id="filedata" size="27"></td>
          </tr>
          <tr class="text"> 
            <td valign="top"> <div align="right"><strong>�ļ�˵����</strong><br>
              </div></td>
            <td><textarea name="filedesc" cols="36" rows="4" class="TextBoxT" id="filedesc"></textarea></td>
          </tr>
          <tr> 
            <td>&nbsp;</td>
            <td><input type="submit" name="Submit" value="�ϴ�����"></td>
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
