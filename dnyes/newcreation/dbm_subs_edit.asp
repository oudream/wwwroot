<script language="JavaScript">

function checkform()
{
	if (form_administrator.filetitle.value.length == 0) {
		alert("��������Ʒ���� ");
		form_administrator.filetitle.focus();
		return false;
		}
	if (form_administrator.filedata.value.length == 0) {
		alert("�����Ҫ�ϴ����ļ�");
		form_administrator.filedata.focus();
		return false;
		}
	return true;
}
</script>
<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">

<title></title>
</head>
<body topmargin="0">
<!--#include file="dbm_adminconn.asp" -->
<!--#include file="Config.asp"-->


<%
subs_id=request("subs_id")
sql="select * from dbm_subs where subs_id=" & subs_id 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1
%>
<table width="75%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
<form action="SaveAsUpload.asp" method="post" enctype="multipart/form-data"  name="form_administrator" id="form_administrator" onSubmit="return checkform();">
  <table width="100%" border="1" align="center" cellpadding="3" cellspacing="0" bordercolor="#CCCCCC" bgcolor="#ffffff">
      <TR align="center"> 
        <TD height="30" colspan="2"><font color="#FF0000" size="+1"><strong>��Ʒ�޸�</strong></font></TD>
      </TR>
    <tr bgcolor="#006699" class="text"> 
      <td height="25" colspan="2" align="center" bgcolor="#FFFFFF"> 
        <%
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
      <td width="46%"><div align="right"><strong>��Ʒ���ƣ�</strong></div></td>
      <td><input name="filetitle" type="text" class="TextBoxT" size="50" value="<%=rs("filetitle")%>"> <input name="filetype" type="hidden" id="filetype" value="��ƷͼƬ">  * 
      </td>
    </tr>
    <tr class="text"> 
      <td valign="top"><div align="right"><strong>�ϴ����ļ���</strong></div></td>
      <td><input name="filedata" type="file" class="TextBoxT" id="filedata" size="50"> * </td>
    </tr>
    <tr class="text"> 
      <td valign="top"> <div align="right"><strong>��Ʒ˵����</strong><br>
        </div></td>
      <td><textarea name="filedesc" cols="60" rows="20" class="TextBoxT" id="filedesc"><%=rs("filedesc")%></textarea></td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td><input type="submit" name="Submit" value="�ϴ����ύ����"><input name="subs_id" type="hidden" id="subs_id" value=<%=subs_id%>></td>
    </tr>
  </table>
        </form>
<%
rs.close
set rs=nothing
%>
<br>
<br>
</body>
</html>
