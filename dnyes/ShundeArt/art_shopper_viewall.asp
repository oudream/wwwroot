<script language="JavaScript">

function checkform()
{
	if (form_shopper.allname.value.length == 0) {
		alert("��������Ϊ��!����������!");
		form_shopper.allname.focus();
		return false;
		}
	
	if(form_shopper.email.value.length > 1)
	{
		if(! isMail(form_shopper.email.value)) {
		alert('��������Ч��E-mail��ַ!');
		form_shopper.email.focus();
		return false;
		}
	}
	return true;
}

function isEnglish(name)
{
	for(i = 0; i < name.length; i++) {
		if(name.charCodeAt(i) > 128)
			return false;
	}
	return true;
}

function isMail(name)
{
	if(!isEnglish(name))
		return false;
	i = name.indexOf("@");
	j = name.lastIndexOf("@");
	if(i == -1)
		return false;
	if(i != j)
		return false;
	if(i == name.length)
		return false;
	return true;
}
</script>

<!--#include file="art_adminconn.asp" -->

<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--media=print ������Կ����ڴ�ӡʱ��Ч--> 
<style media=print> 
.Noprint{display:none;} 
.PageNext{page-break-after: always;} 
</style> 

<style> 
.tdp 
{ 
border-bottom: 1 solid #000000; 
border-left: 1 solid #000000; 
border-right: 0 solid #ffffff; 
border-top: 0 solid #ffffff; 
} 
.tabp 
{ 
border-color: #000000 #000000 #000000 #000000; 
border-style: solid; 
border-top-width: 2px; 
border-right-width: 2px; 
border-bottom-width: 1px; 
border-left-width: 1px; 
} 
.NOPRINT { 
font-family: "����"; 
font-size: 9pt; 
} 

</style>
<title></title></head>
<body topmargin="0">
<table width="550" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td>
	  <center class="Noprint" >
    <OBJECT id=WebBrowser classid=CLSID:8856F961-340A-11D0-A96B-00C04FD705A2 height=0 width=0>
    </OBJECT>
    <input type=button value=��ӡ onclick=document.all.WebBrowser.ExecWB(6,1)>&nbsp;
    <input type=button value=ֱ�Ӵ�ӡ onclick=document.all.WebBrowser.ExecWB(6,6)>&nbsp;
    <input type=button value=ҳ������ onclick=document.all.WebBrowser.ExecWB(8,1)>&nbsp;
    <input name="button" type=button onClick=document.all.WebBrowser.ExecWB(7,1) value=��ӡԤ��>
      </center> 
</td>
    </tr>
  </table>
<%
sql="select * from art_shopper where art_shopper_id="&request("art_shopper_id") 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1 

if rs.eof or rs.bof then 
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' �������� ! ');</SCRIPT>")
response.End()
end if
%>
<TABLE width=600 
                  border=0 align="center" cellPadding=5 cellSpacing=1 bgcolor="#000000" class="tabp">
<FORM action="art_shopper_edit.asp" method="post" name="form_shopper" id="form_shopper" onSubmit="return checkform();">
    <TBODY>
      <TR bgcolor="#FFFFFF" class="tdp"> 
        <TD width="20%" align="right" class="tdp">������</TD>
        <TD class="tdp"><%=rs("allname")%></TD>
      </TR>
      <TR bgcolor="#FFFFFF" class="tdp"> 
        <TD align="right" class="tdp">�Ա�</TD>
        <TD class="tdp"><%=rs("sex")%></TD>
      </TR>
      <TR bgcolor="#FFFFFF" class="tdp"> 
        <TD align="right" class="tdp">���壺</TD>
        <TD class="tdp"><%=rs("nation")%></TD>
      </TR>
      <TR bgcolor="#FFFFFF" class="tdp"> 
        <TD align="right" class="tdp">ԭ�壺</TD>
        <TD class="tdp"><%=rs("original")%></TD>
      </TR>
      <TR bgcolor="#FFFFFF" class="tdp"> 
        <TD align="right" class="tdp">�������£�</TD>
        <TD class="tdp"><%=rs("birth")%></TD>
      </TR>
      <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							              email
<!------------------------------------------------------------------------------------------------------------------------------------ -->
      <TR bgcolor="#FFFFFF" class="tdp"> 
        <TD align="right" class="tdp">������λ��</TD>
        <TD class="tdp"><%=rs("workplace")%></TD>
      </TR>
      <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							              email
<!------------------------------------------------------------------------------------------------------------------------------------ -->
      <TR bgcolor="#FFFFFF" class="tdp"> 
        <TD align="right" class="tdp">ְ��</TD>
        <TD class="tdp"><%=rs("duty")%></TD>
      </TR>
      <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							              email
<!------------------------------------------------------------------------------------------------------------------------------------ -->
      <TR bgcolor="#FFFFFF" class="tdp"> 
        <TD align="right" class="tdp">��Ȥ���ã�</TD>
        <TD class="tdp"><%=rs("hobby")%></TD>
      </TR>
      <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							              USERID
<!----------------------------------------------------------------------------------------------------------------------------------- -->
      <TR bgcolor="#FFFFFF" class="tdp"> 
        <TD align="right" class="tdp">סַ��</TD>
        <TD class="tdp"><%=rs("addr")%></TD>
      </TR>
      <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							         PASSWORD
<!------------------------------------------------------------------------------------------------------------------------------------ -->
      <TR bgcolor="#FFFFFF" class="tdp"> 
        <TD align="right" class="tdp">�ʱࣺ</TD>
        <TD class="tdp"><%=rs("zip")%></TD>
      </TR>
      <TR bgcolor="#FFFFFF" class="tdp"> 
        <TD align="right" class="tdp">�绰��</TD>
        <TD class="tdp"><%=rs("tel")%></TD>
      </TR>
      <TR bgcolor="#FFFFFF" class="tdp"> 
        <TD align="right" class="tdp">E-mail��</TD>
        <TD class="tdp"><%=rs("email")%></TD>
      </TR>
      <TR bgcolor="#FFFFFF" class="tdp"> 
        <TD align="right" class="tdp"> ��ע��</TD>
        <TD height="300" class="tdp"><%=rs("memo")%></TD>
      </TR>
    </TBODY>
	</form>
  </TABLE>
</body>
</html>