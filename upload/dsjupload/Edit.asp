<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Config.asp"-->
<%
CheckLogin
Dim ID,Msg
	ID = Request.QueryString("ID")
	If Request.QueryString("Action") = "Save" Then SaveData ID
	Sub SaveData(ID)
	If ID < 1 Then 
	Response.Write("��������")
	Response.End()
	End If
	myConn.execute("update info set FILETITLE='"&Request.Form("fname")&"',FILEDESC='"&Request.Form("fdesc")&"',FILETYPE='"&Request.Form("ftype")&"',FILEPATH='"&Request.Form("fpath")&"',FILESIZE='"&Request.Form("fsize")&"' where ID="&ID)
	Msg = "�ɹ��޸����ļ�������Ϣ"
	End Sub

If msg <> "" Then
Response.Write("<meta http-equiv=refresh content='3;URL=Admin_List.asp'>"&Msg&"<br>��ҳ����3���ڷ���<BR>�����������û�з�Ӧ����<a href=Admin_List.asp>����˴�����</a>")
Response.End()
End If
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
      &gt;&gt;&gt; ��̨����&nbsp;��&nbsp;<a href="./">�����ϴ���ҳ</a></td>
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
<form name="Edit" action="Edit.asp?Action=Save&ID=<%=ID%>" method="post">
  <tr> 
    <td height="25"> <%
	set frst = Server.CreateObject("adodb.recordset")
	sql = "select * from info Where Id="&ID
	frst.open sql,myconn,1,1
	If not frst.Eof then
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
            <td><input type="text" name="fname" class="TextBoxT" value="<%=ftitle%>"> 
            </td>
            <td width="30%" rowspan="3">&nbsp; </td>
          </tr>
          <tr class="text"> 
            <td width="150"><div align="right">�ļ����ͣ�</div></td>
            <td><select name="ftype" class="TextBoxT" id="filetype">
                <option value="�ز�ͼƬ"<%if ftype="�ز�ͼƬ" then%> selected<% end if %>>�ز�ͼƬ</option>
                <option value="���ù���"<%if ftype="���ù���" then%> selected<% end if %>>���ù���</option>
                <option value="����Դ��"<%if ftype="����Դ��" then%> selected<% end if %>>����Դ��</option>
              </select></td>
          </tr>
          <tr class="text"> 
            <td width="150"><div align="right">�ļ�·����</div></td>
            <td><input name="fpath" type="text" class="TextBoxT" value="<%=fpath%>" size="50"> 
              <%
			Set Fs = Server.CreateObject("Scripting.FileSystemObject")
			If Fs.FileExists(server.mappath(fPath)) Then
			Response.Write("<img src=Images/isexists.gif")
			End If
			%>
            </td>
          </tr>
          <tr class="text"> 
            <td width="150"><div align="right">˵����</div></td>
            <td colspan="2"><input type="text" name="fdesc" class="TextBoxT" value="<%=fdesc%>"></td>
          </tr>
          <tr class="text"> 
            <td height="1" align="right" bgcolor="#0153A3">�ļ���С��</td>
            <td height="1" colspan="2" bgcolor="#0153A3"> <input type="text" name="fsize" class="TextBoxT" value="<%=fsize%>">
              bytes</td>
          </tr>
          <tr class="text"> 
            <td height="1" align="right" bgcolor="#0153A3">&nbsp;</td>
            <td height="1" colspan="2" bgcolor="#0153A3"> <input type="submit" name="Submit" value="�޸�">&nbsp;<input type="button" name="Submit2" value="����" onclick='javascript:window.location="Admin_List.asp"'></td>
          </tr>
          <tr class="text"> 
            <td height="1" align="right" bgcolor="#0153A3">&nbsp;</td>
            <td height="1" colspan="2" bgcolor="#0153A3"></td>
          </tr>
        </table>
      <%
	  else
	  %> 
      <table width="100%" border="0" align="center" cellpadding="3" cellspacing="0">
        <tr class="text"> 
          <td>û�ж�Ӧ�����ݣ�</td>
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
            <td align="right" class="heading">&nbsp;&nbsp;&nbsp;&nbsp;</td>
        </tr>
      </table></td>
  </tr></form>
</table>

<table width="778" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#0153A3">
  <tr> 
    <td height="25">&nbsp;</td>
  </tr>
</table>
</body>
</html>