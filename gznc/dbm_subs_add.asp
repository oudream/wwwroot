<script language="JavaScript">

function checkform()
{
	if (form_administrator.sname.value.length == 0) {
		alert("��������Ʒ���� ");
		form_administrator.sname.focus();
		return false;
		}
	if (form_administrator.filedata1.value.length == 0) {
		alert("�����Ҫ�ϴ����ļ� -- 1");
		form_administrator.filedata1.focus();
		return false;
		}
	if (form_administrator.memo.value.length == 0) {
		alert("������˵��");
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
<LINK href="dbm_css.css" type=text/css rel=stylesheet>
</head>
<body topmargin="0">
<!--#include file="dbm_adminconn.asp" -->
<!--#include file="Config.asp"-->

<%
if request("option")="���" then

sname=trim(request("sname"))
sname=replace(sname,"'","''")
sfile=trim(request("sfile"))
sfile=replace(sfile,"'","''")
brief=trim(request("brief"))
brief=replace(brief,"'","''")

usql="select * from dbm_subs where sname='" & sname & "'"
set urs=server.createobject("ADODB.Recordset")
urs.open usql,conn,1,1
if not ( urs.eof or urs.bof ) then
response.Write("<SCRIPT LANGUAGE=JavaScript> alert('��������Ʒ��ͬ�����ڣ�����������������');</SCRIPT>")
urs.close
set urs=nothing
else
urs.close
set urs=nothing

if brief="" then brief=" "

asfile=split(sfile,".",-1,1)
if isarray(asfile) and ubound(asfile)>0 then
sfilelastname=asfile(1)
else
%>
<SCRIPT LANGUAGE=JavaScript>
alert(' �����Ҫ�ϴ����ļ�. ');
window.history.go( -1 );
</SCRIPT>
<%
response.End()
end if

sfilename = Server.MapPath(".")&"\subsimg\"&sname&"."&sfilelastname

sql="insert into dbm_subs (sname,sfile,brief,inserttime) values ('"&sname&"','"&sfile&"','"&brief&"','"&now()&"')"
conn.Execute(sql)
'If not Err Then
'    Response.Redirect("dbm_folder_option.asp?Action=UpFile&Action2=Post&LocalFile="&sfile&"&ToPath="&sfilename&"&zurl=dbm_subs_add.asp")
'End If		
response.Write("<SCRIPT LANGUAGE=JavaScript> alert('��ӳɹ�');</SCRIPT>")
end if
end if
%>
<!----------------------------------------------------------------------------------------------------------------------
<!--                                                   ERROR MSGS
<!-------------------------------------------------------------------------------------------------------------------- -->
<table width="75%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
<form action="dbm_upfile.asp" method="post" enctype="multipart/form-data"  name="form_administrator" id="form_administrator" onSubmit="return checkform();">
  <table width="100%" border="1" align="center" cellpadding="3" cellspacing="0" bordercolor="#CCCCCC" bgcolor="#ffffff">
      <TR align="center"> 
        <TD height="30" colspan="2"><font color="#FF0000" size="+1"><strong>��Ʒ���</strong></font></TD>
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
		Response.Write "<br>  �����ϴ����ļ����:"&Oksize / 1024&"KB ���밴˳��ѡ��Ҫ�ϴ���ͼƬ��"
		%></td>
    </tr>
    <tr class="text"> 
      <td width="46%"><div align="right"><strong>��Ʒ���ƣ�</strong></div></td>
      <td><input name="sname" type="text" class="sname" size="50"> * 
      </td>
    </tr>
    <tr class="text"> 
      <td valign="top"><div align="right"><strong>�ϴ����ļ� -- 1��</strong></div></td>
      <td><input name="filedata1" type="file" class="TextBoxT" id="filedata1" size="50"><font color="#FF0000">*</font></td>
    </tr>
    <tr class="text"> 
      <td valign="top"><div align="right"><strong>�ϴ����ļ� -- 2��</strong></div></td>
      <td><input name="filedata2" type="file" class="TextBoxT" id="filedata2" size="50">
      </td>
    </tr>
    <tr class="text"> 
      <td valign="top"><div align="right"><strong>�ϴ����ļ� -- 3��</strong></div></td>
      <td><input name="filedata3" type="file" class="TextBoxT" id="filedata3" size="50">
      </td>
    </tr>
    <tr class="text"> 
      <td valign="top"><div align="right"><strong>�ϴ����ļ� -- 4��</strong></div></td>
      <td><input name="filedata4" type="file" class="TextBoxT" id="filedata4" size="50">
      </td>
    </tr>
    <tr class="text"> 
      <td valign="top"><div align="right"><strong>�ϴ����ļ� -- 5��</strong></div></td>
      <td><input name="filedata5" type="file" class="TextBoxT" id="filedata5" size="50">
      </td>
    </tr>
    <tr class="text"> 
      <td valign="top"><div align="right"><strong>�ϴ����ļ� -- 6��</strong></div></td>
      <td><input name="filedata6" type="file" class="TextBoxT" id="filedata6" size="50">
      </td>
    </tr>
    <tr class="text"> 
      <td valign="top"><div align="right"><strong>�ϴ����ļ� -- 7��</strong></div></td>
      <td><input name="filedata7" type="file" class="TextBoxT" id="filedata7" size="50">
      </td>
    </tr>
    <tr class="text"> 
      <td valign="top"><div align="right"><strong>�ϴ����ļ� -- 8��</strong></div></td>
      <td><input name="filedata8" type="file" class="TextBoxT" id="filedata8" size="50">
      </td>
    </tr>
    <tr class="text"> 
      <td valign="middle"> <div align="right"><strong>��Ʒ˵����</strong><br>
        </div></td>
      <td valign="middle"><textarea name="memo" cols="60" rows="20" class="TextBoxT" id="memo"></textarea>
        *</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td><input type="submit" name="Submit" value="�ϴ����ύ����"></td>
    </tr>
  </table>
        <input type="hidden" name="filepath" value="photo">
        <input type="hidden"  name="act" value="upload">
        </form>
<br><br>
</body>
</html>
