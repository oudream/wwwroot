<script language="JavaScript">

function checkform()
{
	if (form_administrator.allname.value.length == 0) {
		alert("请输入全名");
		form_administrator.allname.focus();
		return false;
		}
	if(form_administrator.email.value.length > 1){
	if(! isMail(form_administrator.email.value)) {
		alert('请输入正确的邮件地址');
		form_administrator.email.focus();
		return false;
		}}
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
	if(! isEnglish(name))
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
<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--media=print 这个属性可以在打印时有效--> 
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
font-family: "宋体"; 
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
    <input type=button value=打印 onclick=document.all.WebBrowser.ExecWB(6,1)>&nbsp;
    <input type=button value=直接打印 onclick=document.all.WebBrowser.ExecWB(6,6)>&nbsp;
    <input type=button value=页面设置 onclick=document.all.WebBrowser.ExecWB(8,1)>&nbsp;
    <input name="button" type=button onClick=document.all.WebBrowser.ExecWB(7,1) value=打印预览>
      </center> 
</td>
    </tr>
  </table>
<!--#include file="art_adminconn.asp" -->





<%
sql="select * from art_business where art_business_id="&request("art_business_id") 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1 

if rs.eof or rs.bof then 
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' None data ! ');</SCRIPT>")
response.End()
end if
%>
  
<TABLE width=600 height=307 
                  border=0 align="center" cellPadding=5 cellSpacing=1 bgcolor="#000000" class="tabp">
  <TBODY>
    <TR bgcolor="#FFFFFF" class="tdp"> 
      <TD width="20%" align="right" class="tdp"> 全称：</TD>
      <TD class="tdp"> <%=rs("allname")%></TD>
    </TR>
    <TR bgcolor="#FFFFFF" class="tdp"> 
      <TD align="right" class="tdp"> 地址：</TD>
      <TD class="tdp"> <%=rs("addr")%> </TD>
    </TR>
    <TR bgcolor="#FFFFFF" class="tdp"> 
      <TD align="right" class="tdp"> 电话：</TD>
      <TD class="tdp"> <%=rs("tel")%> </TD>
    </TR>
    <TR bgcolor="#FFFFFF" class="tdp"> 
      <TD align="right" class="tdp"> 传真：</TD>
      <TD class="tdp"> <%=rs("fax")%> </TD>
    </TR>
    <TR bgcolor="#FFFFFF" class="tdp"> 
      <TD align="right" class="tdp"> 邮编：</TD>
      <TD class="tdp"> <%=rs("zip")%> </TD>
    </TR>
    <TR bgcolor="#FFFFFF" class="tdp"> 
      <TD align="right" class="tdp"> 邮箱：</TD>
      <TD class="tdp"> <%=rs("email")%> </TD>
    </TR>
    <TR bgcolor="#FFFFFF" class="tdp"> 
      <TD align="right" class="tdp"> 网址：</TD>
      <TD class="tdp"> <%=rs("web_addr")%> </TD>
    </TR>
    <TR bgcolor="#FFFFFF" class="tdp"> 
      <TD align="right" class="tdp"> 联系人</TD>
      <TD class="tdp"> <%=rs("legal_representative")%> </TD>
    </TR>
    <TR bgcolor="#FFFFFF" class="tdp"> 
      <TD align="right" class="tdp"> 简介：</TD>
      <TD height="300" valign="top" class="tdp"> <%=rs("brief")%> </TD>
    </TR>
    <TR bgcolor="#FFFFFF" class="tdp"> 
      <TD align="right" class="tdp"> 备注：</TD>
      <TD height="300" valign="top" class="tdp"> <%=rs("memo")%> </TD>
    </TR>
  </TBODY>
</TABLE>
<br>
<br>
</body>
</html>
