<script language="JavaScript">

function checkform()
{
	if (form_administrator.sname.value.length == 0) {
		alert("��������Ʒ���� ");
		form_administrator.sname.focus();
		return false;
		}
	if (form_administrator.memo.value.length == 0) {
		alert("����д˵��");
		form_administrator.memo.focus();
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
<form action="dbm_upASfile.asp" method="post" enctype="multipart/form-data"  name="form_administrator" id="form_administrator" onSubmit="return checkform();">
  <table width="100%" border="1" align="center" cellpadding="3" cellspacing="0" bordercolor="#CCCCCC" bgcolor="#ffffff">
    <TR align="center"> 
      <TD height="30" colspan="2"><font color="#FF0000" size="+1"><strong>��Ʒ�޸�</strong></font></TD>
    </TR>
    <tr bgcolor="#006699" class="text"> 
      <td height="25" colspan="2" align="center" bgcolor="#FFFFFF"> <%
		Response.Write "  �����뻻���ļ�����:<br> "
		Set Fs = Server.CreateObject("Scripting.FileSystemObject")
		For Each str In OKAr
		If Fs.FileExists(Server.MapPath("Images\"& Str &".gif")) Then
		Response.Write "<img src='Images/" & str & ".gif' alt='" & str & "�ļ�'> "
		Else
		Response.Write "<img src='Images/X.gif' alt='" & str & "�ļ�'> "
		End If
		Next
		Set Fs = Nothing
		Response.Write "<br>  �����뻻���ļ����:"&Oksize / 1024&"KB"
		%></td>
    </tr>
    <tr class="text"> 
      <td width="46%"><div align="right"><strong>��Ʒ���ƣ�</strong></div></td>
      <td><input name="sname" id="sname" type="text" class="TextBoxT" size="50" value="<%=rs("sname")%>">
        * </td>
    </tr>
    <tr align="center" class="text"> 
      <td colspan="2"><font color="#FF0000" size="+1"><strong>ѡ���Ϊ��ʱ��ʾ���޸�</strong></font></td>
    </tr>
    <tr class="text"> 
      <td valign="top"><div align="right"><strong>�뻻���ļ� -- 1��</strong></div></td>
      <td><input name="filedata1" type="file" class="TextBoxT" id="filedata1" size="50">*</td>
    </tr>
    <tr class="text"> 
      <td valign="top"><div align="right"><strong>�뻻���ļ� -- 2��</strong></div></td>
      <td><input name="filedata2" type="file" class="TextBoxT" id="filedata2" size="50">*</td>
    </tr>
    <tr class="text"> 
      <td valign="top"><div align="right"><strong>�뻻���ļ� -- 3��</strong></div></td>
      <td><input name="filedata3" type="file" class="TextBoxT" id="filedata3" size="50">*</td>
    </tr>
    <tr class="text"> 
      <td valign="top"><div align="right"><strong>�뻻���ļ� -- 4��</strong></div></td>
      <td><input name="filedata4" type="file" class="TextBoxT" id="filedata4" size="50">*</td>
    </tr>
    <tr class="text"> 
      <td valign="top"><div align="right"><strong>�뻻���ļ� -- 5��</strong></div></td>
      <td><input name="filedata5" type="file" class="TextBoxT" id="filedata5" size="50">*</td>
    </tr>
    <tr class="text"> 
      <td valign="top"><div align="right"><strong>�뻻���ļ� -- 6��</strong></div></td>
      <td><input name="filedata6" type="file" class="TextBoxT" id="filedata6" size="50"> 
      </td>
    </tr>
    <tr class="text"> 
      <td valign="top"><div align="right"><strong>�뻻���ļ� -- 7��</strong></div></td>
      <td><input name="filedata7" type="file" class="TextBoxT" id="filedata7" size="50"> 
      </td>
    </tr>
    <tr class="text"> 
      <td valign="top"><div align="right"><strong>�뻻���ļ� -- 8��</strong></div></td>
      <td><input name="filedata8" type="file" class="TextBoxT" id="filedata8" size="50"> 
      </td>
    </tr>
    <tr class="text"> 
      <td valign="top"> <div align="right"><strong>��Ʒ˵����</strong><br>
        </div></td>
      <td><textarea name="memo" id="memo" cols="60" rows="20" class="TextBoxT"><%=rs("memo")%></textarea></td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td> <input type="submit" name="Submit" value="�ϴ����ύ�༭����"> <input name="subs_id" type="hidden" id="subs_id" value=<%=subs_id%>> 
        &nbsp;&nbsp;&nbsp; <INPUT type="button" id="goback" name="goback" value="����" onClick="history.go( -1 );"> 
      </td>
    </tr>
  </table>
        <input type="hidden" name="filepath" value="photo">
        <input type="hidden"  name="act" value="upload">
        </form>
<%
rs.close
set rs=nothing
%>
<br>
<br>
</body>
</html>
