<!--#include file="mymefaq5411jkjkh/favorible/showme.asp" -->
<!--#include file="css.asp"-->
<%
id=request("id")
if id="" then
response.redirect "viewmyorders.asp"
response.end
end if
%>
<%
sql="select * from orders where id="&id
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
if rs("ifdeal")=true then
response.Write("<table width=100% height=100% border=0 align=center cellpadding=0 cellspacing=0><tr><td align=center>�ö������ڴ����У�����ɾ��,�����κ���������<a href="&"emailtome.asp"&"><font color=red>������ϵ</font></a><br><br></td></tr></table>")
%>
<script language="JavaScript">
setTimeout("history.go( -1 );return true;",2000);
</script>
<%
response.end()
end if
%>
<body leftmargin="0" topmargin="0">
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td align="center" valign="middle">
<table width="498" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#CCCCCC">
  <tr bgcolor="#f5f5f5"> 
          <td height="40" colspan="2">���Ƿ�ȷ��Ҫɾ���˶�����</td>
  </tr>
  <tr> 
    <td width="25%" height="25" bgcolor="#f5f5f5">ID</td>
    <td bgcolor="#FFFFFF"><%=rs("id")%></td>
  </tr>
  <tr> 
    <td height="25" bgcolor="#f5f5f5">�����ţ�</td>
    <td bgcolor="#FFFFFF"><%=rs("inBillNo")%></td>
  </tr>
  <tr> 
    <td height="25" bgcolor="#f5f5f5">��Ʒ���ƣ�</td>
    <td bgcolor="#FFFFFF"><%=rs("xiangxi")%></td>
  </tr>
  <tr> 
    <td height="25" bgcolor="#f5f5f5">�����̣�</td>
    <td bgcolor="#FFFFFF"><%=rs("xxgys")%></td>
  </tr>
  <tr> 
    <td height="25" bgcolor="#f5f5f5">����ʱ�䣺</td>
    <td bgcolor="#FFFFFF"><%=rs("ordertime")%></td>
  </tr>
  <tr> 
    <td height="25" bgcolor="#f5f5f5">֧����ʽ��</td>
    <td bgcolor="#FFFFFF"><%=rs("paymenttype")%></td>
  </tr>
  <tr> 
    <td height="25" bgcolor="#f5f5f5">�˶���״̬��</td>
    <td bgcolor="#FFFFFF">
<%if rs("ifcomp")=true then
response.write "�Ѿ���ͨ"
else
if rs("ifpay")=true then
response.write "�����ѵ�����δ��ͨ��"
else
response.write "��δ����"
end if
end if
%> </td>
  </tr>
  <tr> 
    <td height="25" bgcolor="#f5f5f5"><font color="red">�����ܽ�</font></td>
    <td bgcolor="#FFFFFF"><%=rs("summoney")%> RMB </td>
  </tr>
  <tr> 
          <td height="25" bgcolor="#f5f5f5">�ر�˵����</td>
    <td bgcolor="#FFFFFF">
<%
ordernote = replace(rs("ordernote"),chr(13),"<br>")
%>	
<%=ordernote%>						  
	</td>
  </tr>
<%if rs("iffalse")=true then%>
  <tr> 
    <td height="25" bgcolor="#f5f5f5"><font color="red">�˶���������Ϣ��</font></td>
    <td bgcolor="#FFFFFF"> 
<%
errmsg=replace(rs("errmsg"),chr(13),"<br>")
errmsg=replace(errmsg,chr(32),"&nbsp;")
response.write errmsg
%>
 </td>
  </tr>
<%end if%>
  <tr align="center"> 
    <td height="45" colspan="2" bgcolor="#FFFFFF">
              <p><a href="ordercancel.asp?id=<%=request("id")%>">ȷ��ɾ��</a>&nbsp;&nbsp;&nbsp;&nbsp;/&nbsp;&nbsp;&nbsp;&nbsp;<a href="viewmyorders.asp">����</a></p>
      </td>
  </tr>
</table>
      <%
rs.close
set rs=nothing
%>
      <table width="75%" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td height="80">&nbsp;</td>
        </tr>
      </table> </td>
  </tr>
</table>

