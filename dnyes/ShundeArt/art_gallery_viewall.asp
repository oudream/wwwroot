
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
sql="select * from art_gallery where art_gallery_id="&request("art_gallery_id") 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1 

if rs.eof or rs.bof then 
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' None data ! ');</SCRIPT>")
response.End()
end if
%>
<TABLE width=600 
                  border=0 align="center" cellPadding=5 cellSpacing=1 bgcolor="#000000" class="tabp">
  <TBODY>
    <TR bgcolor="#FFFFFF" class="tdp"> 
      <TD colspan="2" align="right" class="tdp" width="36%">���ƣ�</TD>
      <TD class="tdp"><%=rs("gname")%></TD>
    </TR>
    <TR bgcolor="#FFFFFF" class="tdp"> 
      <TD colspan="2" align="right" class="tdp">��ַ��</TD>
      <TD class="tdp"><%=rs("addr")%></TD>
    </TR>
    <TR bgcolor="#FFFFFF" class="tdp"> 
      <TD colspan="2" align="right" class="tdp">�ʱࣺ</TD>
      <TD class="tdp"><%=rs("zip")%></TD>
    </TR>
    <TR bgcolor="#FFFFFF" class="tdp"> 
      <TD colspan="2" align="right" class="tdp"> �绰��</TD>
      <TD class="tdp"><%=rs("tel")%></TD>
    </TR>
    <TR bgcolor="#FFFFFF" class="tdp"> 
      <TD colspan="2" align="right" class="tdp">���棺</TD>
      <TD class="tdp"><%=rs("fax")%></TD>
    </TR>
    <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							              email
<!------------------------------------------------------------------------------------------------------------------------------------ -->
    <TR bgcolor="#FFFFFF" class="tdp"> 
      <TD colspan="2" align="right" class="tdp"> �������䣺</TD>
      <TD class="tdp"><%=rs("email")%></TD>
    </TR>
    <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							              email
<!------------------------------------------------------------------------------------------------------------------------------------ -->
    <TR bgcolor="#FFFFFF" class="tdp"> 
      <TD colspan="2" align="right" class="tdp"> ��ַ��</TD>
      <TD class="tdp"><%=rs("website_addr")%></TD>
    </TR>
    <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							              email
<!------------------------------------------------------------------------------------------------------------------------------------ -->
    <TR bgcolor="#FFFFFF" class="tdp"> 
      <TD width="10%" rowspan="3" align="right" class="tdp"> �����ˣ�</TD>
      <TD width="26%" align="right" class="tdp">������</TD>
      <TD class="tdp"><%=rs("principal")%></TD>
    </TR>
    <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							              USERID
<!------------------------------------------------------------------------------------------------------------------------------------ -->
    <TR> 
      <TD align="right" bgcolor="#FFFFFF" class="tdp">�ֻ����룺</TD>
      <TD bgcolor="#FFFFFF" class="tdp"><%=rs("mp_principal")%></TD>
    </TR>
    <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							         PASSWORD
<!------------------------------------------------------------------------------------------------------------------------------------ -->
    <TR> 
      <TD align="right" bgcolor="#FFFFFF" class="tdp">�칫�绰��</TD>
      <TD bgcolor="#FFFFFF" class="tdp"><%=rs("wp_principal")%></TD>
    </TR>
    <TR bgcolor="#FFFFFF" class="tdp"> 
      <TD colspan="2" align="right" class="tdp"> ��飺</TD>
      <TD height="300" valign="top" class="tdp"><%=rs("briefintro")%></TD>
    </TR>
    <TR bgcolor="#FFFFFF" class="tdp">
      <TD colspan="2" align="right" class="tdp"> ��ע��</TD>
      <TD height="300" valign="top" class="tdp"><%=rs("memo")%></TD>
    </TR>
  </TBODY>
</TABLE>
</body>
</html>