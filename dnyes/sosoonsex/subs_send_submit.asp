<script language="JavaScript">

function addnum(value_1,value_2,selObj){ 
selObj.value=value_1*value_2;
}

function addnum_all(a,b,c,d,e,f,g,h,i,j,k,l,m,n){ 
form_administrator.counter_all.value=0;
if ( a != 0 ) 
form_administrator.counter_all.value=parseFloat(form_administrator.counter_all.value)+parseFloat(form_administrator.counter_1.value);
if ( b != 0 ) 
form_administrator.counter_all.value=parseFloat(form_administrator.counter_all.value)+parseFloat(form_administrator.counter_2.value);
if ( c != 0 ) 
form_administrator.counter_all.value=parseFloat(form_administrator.counter_all.value)+parseFloat(form_administrator.counter_3.value);
if ( d != 0 ) 
form_administrator.counter_all.value=parseFloat(form_administrator.counter_all.value)+parseFloat(form_administrator.counter_4.value);
if ( e != 0 ) 
form_administrator.counter_all.value=parseFloat(form_administrator.counter_all.value)+parseFloat(form_administrator.counter_5.value);
if ( f != 0 ) 
form_administrator.counter_all.value=parseFloat(form_administrator.counter_all.value)+parseFloat(form_administrator.counter_6.value);
if ( g != 0 ) 
form_administrator.counter_all.value=parseFloat(form_administrator.counter_all.value)+parseFloat(form_administrator.counter_7.value);
if ( h != 0 ) 
form_administrator.counter_all.value=parseFloat(form_administrator.counter_all.value)+parseFloat(form_administrator.counter_8.value);
if ( i != 0 ) 
form_administrator.counter_all.value=parseFloat(form_administrator.counter_all.value)+parseFloat(form_administrator.counter_9.value);
if ( j != 0 ) 
form_administrator.counter_all.value=parseFloat(form_administrator.counter_all.value)+parseFloat(form_administrator.counter_10.value);
if ( k != 0 ) 
form_administrator.counter_all.value=parseFloat(form_administrator.counter_all.value)+parseFloat(form_administrator.counter_11.value);
if ( l != 0 ) 
form_administrator.counter_all.value=parseFloat(form_administrator.counter_all.value)+parseFloat(form_administrator.counter_12.value);
if ( m != 0 ) 
form_administrator.counter_all.value=parseFloat(form_administrator.counter_all.value)+parseFloat(form_administrator.counter_13.value);
if ( n != 0 ) 
form_administrator.counter_all.value=parseFloat(form_administrator.counter_all.value)+parseFloat(form_administrator.counter_14.value);
}

function checkform(a,b,c,d,e,f,g,h,i,j,k,l,m,n)
{
	if( a != 0 ) {
	if(! TFNumber(form_administrator.mount_1.value)) {
		alert("��������ȷ������.");
		form_administrator.mount_1.focus();
		return false;
		}
	if(! isNumber(form_administrator.price_1.value)) {
		alert("��������ȷ�ĵ���.");
		form_administrator.price_1.focus();
		return false;
		}
	if(! isNumber(form_administrator.counter_1.value)) {
		alert("��������ȷ�Ľ��.");
		form_administrator.counter_1.focus();
		return false;
		}
		}
		

	if( b != 0 ) {
	if(! TFNumber(form_administrator.mount_2.value)) {
		alert("��������ȷ������.");
		form_administrator.mount_2.focus();
		return false;
		}
	if(! isNumber(form_administrator.price_2.value)) {
		alert("��������ȷ�ĵ���.");
		form_administrator.price_2.focus();
		return false;
		}
	if(! isNumber(form_administrator.counter_2.value)) {
		alert("��������ȷ�Ľ��.");
		form_administrator.counter_2.focus();
		return false;
		}
		}
		

	if( c != 0 ) {
	if(! TFNumber(form_administrator.mount_3.value)) {
		alert("��������ȷ������.");
		form_administrator.mount_3.focus();
		return false;
		}
	if(! isNumber(form_administrator.price_3.value)) {
		alert("��������ȷ�ĵ���.");
		form_administrator.price_3.focus();
		return false;
		}
	if(! isNumber(form_administrator.counter_3.value)) {
		alert("��������ȷ�Ľ��.");
		form_administrator.counter_3.focus();
		return false;
		}
		}
		

	if( d != 0 ) {
	if(! TFNumber(form_administrator.mount_4.value)) {
		alert("��������ȷ������.");
		form_administrator.mount_4.focus();
		return false;
		}
	if(! isNumber(form_administrator.price_4.value)) {
		alert("��������ȷ�ĵ���.");
		form_administrator.price_4.focus();
		return false;
		}
	if(! isNumber(form_administrator.counter_4.value)) {
		alert("��������ȷ�Ľ��.");
		form_administrator.counter_4.focus();
		return false;
		}
		}
		

	if( e != 0 ) {
	if(! TFNumber(form_administrator.mount_5.value)) {
		alert("��������ȷ������.");
		form_administrator.mount_5.focus();
		return false;
		}
	if(! isNumber(form_administrator.price_5.value)) {
		alert("��������ȷ�ĵ���.");
		form_administrator.price_5.focus();
		return false;
		}
	if(! isNumber(form_administrator.counter_5.value)) {
		alert("��������ȷ�Ľ��.");
		form_administrator.counter_5.focus();
		return false;
		}
		}
		

	if( f != 0 ) {
	if(! TFNumber(form_administrator.mount_6.value)) {
		alert("��������ȷ������.");
		form_administrator.mount_6.focus();
		return false;
		}
	if(! isNumber(form_administrator.price_6.value)) {
		alert("��������ȷ�ĵ���.");
		form_administrator.price_6.focus();
		return false;
		}
	if(! isNumber(form_administrator.counter_6.value)) {
		alert("��������ȷ�Ľ��.");
		form_administrator.counter_6.focus();
		return false;
		}
		}
		

	if( g != 0 ) {
	if(! TFNumber(form_administrator.mount_7.value)) {
		alert("��������ȷ������.");
		form_administrator.mount_7.focus();
		return false;
		}
	if(! isNumber(form_administrator.price_7.value)) {
		alert("��������ȷ�ĵ���.");
		form_administrator.price_7.focus();
		return false;
		}
	if(! isNumber(form_administrator.counter_7.value)) {
		alert("��������ȷ�Ľ��.");
		form_administrator.counter_7.focus();
		return false;
		}
		}
		

	if( h != 0 ) {
	if(! TFNumber(form_administrator.mount_8.value)) {
		alert("��������ȷ������.");
		form_administrator.mount_8.focus();
		return false;
		}
	if(! isNumber(form_administrator.price_8.value)) {
		alert("��������ȷ�ĵ���.");
		form_administrator.price_8.focus();
		return false;
		}
	if(! isNumber(form_administrator.counter_8.value)) {
		alert("��������ȷ�Ľ��.");
		form_administrator.counter_8.focus();
		return false;
		}
		}
		
	if( i != 0 ) {
	if(! TFNumber(form_administrator.mount_9.value)) {
		alert("��������ȷ������.");
		form_administrator.mount_9.focus();
		return false;
		}
	if(! isNumber(form_administrator.price_9.value)) {
		alert("��������ȷ�ĵ���.");
		form_administrator.price_9.focus();
		return false;
		}
	if(! isNumber(form_administrator.counter_9.value)) {
		alert("��������ȷ�Ľ��.");
		form_administrator.counter_9.focus();
		return false;
		}
		}
		
	if( j != 0 ) {
	if(! TFNumber(form_administrator.mount_10.value)) {
		alert("��������ȷ������.");
		form_administrator.mount_10.focus();
		return false;
		}
	if(! isNumber(form_administrator.price_10.value)) {
		alert("��������ȷ�ĵ���.");
		form_administrator.price_10.focus();
		return false;
		}
	if(! isNumber(form_administrator.counter_10.value)) {
		alert("��������ȷ�Ľ��.");
		form_administrator.counter_10.focus();
		return false;
		}
		}
		
	if( k != 0 ) {
	if(! TFNumber(form_administrator.mount_11.value)) {
		alert("��������ȷ������.");
		form_administrator.mount_11.focus();
		return false;
		}
	if(! isNumber(form_administrator.price_11.value)) {
		alert("��������ȷ�ĵ���.");
		form_administrator.price_11.focus();
		return false;
		}
	if(! isNumber(form_administrator.counter_11.value)) {
		alert("��������ȷ�Ľ��.");
		form_administrator.counter_11.focus();
		return false;
		}
		}
		
	if( l != 0 ) {
	if(! TFNumber(form_administrator.mount_12.value)) {
		alert("��������ȷ������.");
		form_administrator.mount_12.focus();
		return false;
		}
	if(! isNumber(form_administrator.price_12.value)) {
		alert("��������ȷ�ĵ���.");
		form_administrator.price_12.focus();
		return false;
		}
	if(! isNumber(form_administrator.counter_12.value)) {
		alert("��������ȷ�Ľ��.");
		form_administrator.counter_12.focus();
		return false;
		}
		}
		
	if( m != 0 ) {
	if(! TFNumber(form_administrator.mount_13.value)) {
		alert("��������ȷ������.");
		form_administrator.mount_13.focus();
		return false;
		}
	if(! isNumber(form_administrator.price_13.value)) {
		alert("��������ȷ�ĵ���.");
		form_administrator.price_13.focus();
		return false;
		}
	if(! isNumber(form_administrator.counter_13.value)) {
		alert("��������ȷ�Ľ��.");
		form_administrator.counter_13.focus();
		return false;
		}
		}
		
	if( n != 0 ) {
	if(! TFNumber(form_administrator.mount_14.value)) {
		alert("��������ȷ������.");
		form_administrator.mount_14.focus();
		return false;
		}
	if(! isNumber(form_administrator.price_14.value)) {
		alert("��������ȷ�ĵ���.");
		form_administrator.price_14.focus();
		return false;
		}
	if(! isNumber(form_administrator.counter_14.value)) {
		alert("��������ȷ�Ľ��.");
		form_administrator.counter_14.focus();
		return false;
		}
		}
		
	if(! isNumber(form_administrator.counter_all.value)) {
		alert("��������ȷ�Ľ��.");
		form_administrator.counter_all.focus();
		return false;
		}
	if (form_administrator.gid.value.length == 0) {
		alert("����д�ͻ� ");
		form_administrator.gid.focus();
		return false;
		}
	if (form_administrator.pyear.value.length == 0) {
		alert("��������� ");
		form_administrator.pyear.focus();
		return false;
		}
	if(! TFNumber(form_administrator.pyear.value)) {
		alert("����������д���");
		form_administrator.pyear.focus();
		return false;
		}
	if (form_administrator.pyear.value > 2005 || form_administrator.pyear.value < 2003) {
		alert("��������ȷ�����, ������ 2004 �� 2005 . ");
		form_administrator.pyear.focus();
		return false;
		}
		
	if (form_administrator.pmonth.value.length == 0) {
		alert("�������·� ");
		form_administrator.pmonth.focus();
		return false;
		}
	if(! TFNumber(form_administrator.pmonth.value)) {
		alert("����������д�·�");
		form_administrator.pmonth.focus();
		return false;
		}
	if (form_administrator.pmonth.value > 12 || form_administrator.pmonth.value < 1) {
		alert("��������ȷ���·�, ������ 1 �� 12 . ");
		form_administrator.pmonth.focus();
		return false;
		}

	if (form_administrator.pday.value.length == 0) {
		alert("���������� ");
		form_administrator.pday.focus();
		return false;
		}
	if(! TFNumber(form_administrator.pday.value)) {
		alert("����������д����");
		form_administrator.pday.focus();
		return false;
		}
	if (form_administrator.pday.value > 31 || form_administrator.pday.value < 1) {
		alert("��������ȷ������, ������ 1 �� 31 . ");
		form_administrator.pday.focus();
		return false;
		}


	if (form_administrator.operationer_id.value == 0) {
		alert("��ѡ�񱾹�˾������ ");
		form_administrator.operationer_id.focus();
		return false;
		}
	return true;
}

function isNumber(name)
{
	if (name.length == 0) 
		return false;
	for(i = 0; i < name.length; i++) {
		if (!((name.charAt(i) >= "0" && name.charAt(i) <= "9") || name.charAt(i) == "."))
			return false;
	}
	return true;
}

function TFNumber(name)
{
	if (name.length == 0) 
		return false;
	for(i = 0; i < name.length; i++) {
		if (name.charAt(i) < "0" || name.charAt(i) > "9")
			return false;
	}
	return true;
}

</script>
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="css.css" type=text/css rel=stylesheet>
</head>
<body text="#000000" leftmargin="20" topmargin="8" marginwidth="0" marginheight="0">
<!--#include file="adminconn.asp" -->
<!----------------------------------------------------------------------------------------------------------------------
<!--                                                   ERROR MSGS
<!-------------------------------------------------------------------------------------------------------------------- -->
<FORM action="subs_send_submit_ok.asp" method="post" name="form_administrator" id="form_administrator" onSubmit="return checkform(<%=session("subs_id_1")%>,<%=session("subs_id_2")%>,<%=session("subs_id_3")%>,<%=session("subs_id_4")%>,<%=session("subs_id_5")%>,<%=session("subs_id_6")%>,<%=session("subs_id_7")%>,<%=session("subs_id_8")%>,<%=session("subs_id_9")%>,<%=session("subs_id_10")%>,<%=session("subs_id_11")%>,<%=session("subs_id_12")%>,<%=session("subs_id_13")%>,<%=session("subs_id_14")%>);">
  <TABLE cellSpacing=0 cellPadding=0 width=600 border=0>
    <TBODY>
      <TR align="center">
        <TD height="60" colSpan=2><br><font size="6" face="����">�� �� �� �� �� �� ��</font> 
        </TD>
      </TR>
      <TR align="center"> 
        <TD height="50" colSpan=2><table width="100%" border="1" cellspacing="0" cellpadding="0">
            <tr> 
              <td width="80">����</td>
              <td width="150" height="35">��Ʒ���Ƽ��ͺ�/���</td>
              <td width="30" height="35" align="center">��λ</td>
              <td width="80" height="35" align="center">����</td>
              <td width="80" height="35" align="center">����</td>
              <td width="80" height="35" align="center">���</td>
              <td width="130" height="35" align="center">��ע</td>
            </tr>
            <%
Set rs=Server.CreateObject("ADODB.RecordSet")
Set brs=Server.CreateObject("ADODB.RecordSet")
Set prs=Server.CreateObject("ADODB.RecordSet")
if session("subs_id_1")<>0 then
sql="select * from subs where subs_id=" & session("subs_id_1")
rs.open sql,conn,1,1
if not(rs.eof or rs.bof) then 
bsql="select * from brand where brand_id=" & rs("brand_id")
brs.open bsql,conn,1,1
%>
            <tr> 
              <td width="80">CPU</td>
              <td width="150" height="30"> <%
if not(brs.eof or brs.bof) then response.Write(brs("bname"))
response.Write(" -- " & rs("code"))			 
brs.close
%> </td>
              <td height="30" align="center"> <%
response.Write(rs("suit"))			 
%> </td>
              <td height="30" align="center"><input name="mount_1" id="mount_1" type="text" size="8" maxlength="8" value=1></td>
              <td height="30" align="center"><input name="price_1" id="price_1" type="text" size="8" maxlength="8" value=<%=rs("pricein")%>></td>
              <td height="30" align="center"><input name="counter_1" id="counter_1" type="text" size="8" maxlength="8" onFocus="addnum(form_administrator.mount_1.value,form_administrator.price_1.value,this);"></td>
              <td height="30" align="center"><input name="memo_1" id="memo_1" type="text" size="15" maxlength="100"></td>
            </tr>
            <%
end if
rs.close
end if
%>
            <%
if session("subs_id_2")<>0 then
sql="select * from subs where subs_id=" & session("subs_id_2")
rs.open sql,conn,1,1
if not(rs.eof or rs.bof) then 
bsql="select * from brand where brand_id=" & rs("brand_id")
brs.open bsql,conn,1,1
%>
            <tr> 
              <td width="80">����</td>
              <td width="150" height="30"> <%
if not(brs.eof or brs.bof) then response.Write(brs("bname"))
response.Write(" -- " & rs("code"))			 
brs.close
%> </td>
              <td height="30" align="center"> <%
response.Write(rs("suit"))			 
%> </td>
              <td height="30" align="center"><input name="mount_2" id="mount_2" type="text" size="8" maxlength="8" value=1></td>
              <td height="30" align="center"><input name="price_2" id="price_2" type="text" size="8" maxlength="8" value=<%=rs("pricein")%>></td>
              <td height="30" align="center"><input name="counter_2" id="counter_2" type="text" size="8" maxlength="8" onFocus="addnum(form_administrator.mount_2.value,form_administrator.price_2.value,this);"></td>
              <td height="30" align="center"><input name="memo_2" id="memo_2" type="text" size="15" maxlength="100"></td>
            </tr>
            <%
end if
rs.close
end if
%>
            <%
if session("subs_id_3")<>0 then
sql="select * from subs where subs_id=" & session("subs_id_3")
rs.open sql,conn,1,1
if not(rs.eof or rs.bof) then 
bsql="select * from brand where brand_id=" & rs("brand_id")
brs.open bsql,conn,1,1
%>
            <tr> 
              <td width="80">�ڴ�</td>
              <td width="150" height="30"> <%
if not(brs.eof or brs.bof) then response.Write(brs("bname"))
response.Write(" -- " & rs("code"))			 
brs.close
%> </td>
              <td height="30" align="center"> <%
response.Write(rs("suit"))			 
%> </td>
              <td height="30" align="center"><input name="mount_3" id="mount_3" type="text" size="8" maxlength="8" value=1></td>
              <td height="30" align="center"><input name="price_3" id="price_3" type="text" size="8" maxlength="8" value=<%=rs("pricein")%>></td>
              <td height="30" align="center"><input name="counter_3" id="counter_3" type="text" size="8" maxlength="8" onFocus="addnum(form_administrator.mount_3.value,form_administrator.price_3.value,this);"></td>
              <td height="30" align="center"><input name="memo_3" id="memo_3" type="text" size="15" maxlength="100"></td>
            </tr>
            <%
end if
rs.close
end if
%>
            <%
if session("subs_id_4")<>0 then
sql="select * from subs where subs_id=" & session("subs_id_4")
rs.open sql,conn,1,1
if not(rs.eof or rs.bof) then 
bsql="select * from brand where brand_id=" & rs("brand_id")
brs.open bsql,conn,1,1
%>
            <tr> 
              <td width="80">Ӳ��</td>
              <td width="150" height="30"> <%
if not(brs.eof or brs.bof) then response.Write(brs("bname"))
response.Write(" -- " & rs("code"))			 
brs.close
%> </td>
              <td height="30" align="center"> <%
response.Write(rs("suit"))			 
%> </td>
              <td height="30" align="center"><input name="mount_4" id="mount_4" type="text" size="8" maxlength="8" value=1></td>
              <td height="30" align="center"><input name="price_4" id="price_4" type="text" size="8" maxlength="8" value=<%=rs("pricein")%>></td>
              <td height="30" align="center"><input name="counter_4" id="counter_4" type="text" size="8" maxlength="8" onFocus="addnum(form_administrator.mount_4.value,form_administrator.price_4.value,this);"></td>
              <td height="30" align="center"><input name="memo_4" id="memo_4" type="text" size="15" maxlength="100"></td>
            </tr>
            <%
end if
rs.close
end if
%>
            <%
if session("subs_id_5")<>0 then
sql="select * from subs where subs_id=" & session("subs_id_5")
rs.open sql,conn,1,1
if not(rs.eof or rs.bof) then 
bsql="select * from brand where brand_id=" & rs("brand_id")
brs.open bsql,conn,1,1
%>
            <tr> 
              <td width="80">��ʾ��</td>
              <td width="150" height="30"> <%
if not(brs.eof or brs.bof) then response.Write(brs("bname"))
response.Write(" -- " & rs("code"))			 
brs.close
%> </td>
              <td height="30" align="center"> <%
response.Write(rs("suit"))			 
%> </td>
              <td height="30" align="center"><input name="mount_5" id="mount_5" type="text" size="8" maxlength="8" value=1></td>
              <td height="30" align="center"><input name="price_5" id="price_5" type="text" size="8" maxlength="8" value=<%=rs("pricein")%>></td>
              <td height="30" align="center"><input name="counter_5" id="counter_5" type="text" size="8" maxlength="8" onFocus="addnum(form_administrator.mount_5.value,form_administrator.price_5.value,this);"></td>
              <td height="30" align="center"><input name="memo_5" id="memo_5" type="text" size="15" maxlength="100"></td>
            </tr>
            <%
end if
rs.close
end if
%>
            <%
if session("subs_id_6")<>0 then
sql="select * from subs where subs_id=" & session("subs_id_6")
rs.open sql,conn,1,1
if not(rs.eof or rs.bof) then 
bsql="select * from brand where brand_id=" & rs("brand_id")
brs.open bsql,conn,1,1
%>
            <tr> 
              <td width="80">��ʾ��</td>
              <td width="150" height="30"> <%
if not(brs.eof or brs.bof) then response.Write(brs("bname"))
response.Write(" -- " & rs("code"))			 
brs.close
%> </td>
              <td height="30" align="center"> <%
response.Write(rs("suit"))			 
%> </td>
              <td height="30" align="center"><input name="mount_6" id="mount_6" type="text" size="8" maxlength="8" value=1></td>
              <td height="30" align="center"><input name="price_6" id="price_6" type="text" size="8" maxlength="8" value=<%=rs("pricein")%>></td>
              <td height="30" align="center"><input name="counter_6" id="counter_6" type="text" size="8" maxlength="8" onFocus="addnum(form_administrator.mount_6.value,form_administrator.price_6.value,this);"></td>
              <td height="30" align="center"><input name="memo_6" id="memo_6" type="text" size="15" maxlength="100"></td>
            </tr>
            <%
end if
rs.close
end if
%>
            <%
if session("subs_id_7")<>0 then
sql="select * from subs where subs_id=" & session("subs_id_7")
rs.open sql,conn,1,1
if not(rs.eof or rs.bof) then 
bsql="select * from brand where brand_id=" & rs("brand_id")
brs.open bsql,conn,1,1
%>
            <tr> 
              <td width="80">����</td>
              <td width="150" height="30"> <%
if not(brs.eof or brs.bof) then response.Write(brs("bname"))
response.Write(" -- " & rs("code"))			 
brs.close
%> </td>
              <td height="30" align="center"> <%
response.Write(rs("suit"))			 
%> </td>
              <td height="30" align="center"><input name="mount_7" id="mount_7" type="text" size="8" maxlength="8" value=1></td>
              <td height="30" align="center"><input name="price_7" id="price_7" type="text" size="8" maxlength="8" value=<%=rs("pricein")%>></td>
              <td height="30" align="center"><input name="counter_7" id="counter_7" type="text" size="8" maxlength="8" onFocus="addnum(form_administrator.mount_7.value,form_administrator.price_7.value,this);"></td>
              <td height="30" align="center"><input name="memo_7" id="memo_7" type="text" size="15" maxlength="100"></td>
            </tr>
            <%
end if
rs.close
end if
%>
            <%
if session("subs_id_8")<>0 then
sql="select * from subs where subs_id=" & session("subs_id_8")
rs.open sql,conn,1,1
if not(rs.eof or rs.bof) then 
bsql="select * from brand where brand_id=" & rs("brand_id")
brs.open bsql,conn,1,1
%>
            <tr> 
              <td width="80">����</td>
              <td width="150" height="30"> <%
if not(brs.eof or brs.bof) then response.Write(brs("bname"))
response.Write(" -- " & rs("code"))			 
brs.close
%> </td>
              <td height="30" align="center"> <%
response.Write(rs("suit"))			 
%> </td>
              <td height="30" align="center"><input name="mount_8" id="mount_8" type="text" size="8" maxlength="8" value=1></td>
              <td height="30" align="center"><input name="price_8" id="price_8" type="text" size="8" maxlength="8" value=<%=rs("pricein")%>></td>
              <td height="30" align="center"><input name="counter_8" id="counter_8" type="text" size="8" maxlength="8" onFocus="addnum(form_administrator.mount_8.value,form_administrator.price_8.value,this);"></td>
              <td height="30" align="center"><input name="memo_8" id="memo_8" type="text" size="15" maxlength="100"></td>
            </tr>
            <%
end if
rs.close
end if
%>
            <%
if session("subs_id_9")<>0 then
sql="select * from subs where subs_id=" & session("subs_id_9")
rs.open sql,conn,1,1
if not(rs.eof or rs.bof) then 
bsql="select * from brand where brand_id=" & rs("brand_id")
brs.open bsql,conn,1,1
%>
            <tr> 
              <td width="80">����</td>
              <td width="150" height="30"> <%
if not(brs.eof or brs.bof) then response.Write(brs("bname"))
response.Write(" -- " & rs("code"))			 
brs.close
%> </td>
              <td height="30" align="center"> <%
response.Write(rs("suit"))			 
%> </td>
              <td height="30" align="center"><input name="mount_9" id="mount_9" type="text" size="8" maxlength="8" value=1></td>
              <td height="30" align="center"><input name="price_9" id="price_9" type="text" size="8" maxlength="8" value=<%=rs("pricein")%>></td>
              <td height="30" align="center"><input name="counter_9" id="counter_9" type="text" size="8" maxlength="8" onFocus="addnum(form_administrator.mount_9.value,form_administrator.price_9.value,this);"></td>
              <td height="30" align="center"><input name="memo_9" id="memo_9" type="text" size="15" maxlength="100"></td>
            </tr>
            <%
end if
rs.close
end if
%>
            <%
if session("subs_id_10")<>0 then
sql="select * from subs where subs_id=" & session("subs_id_10")
rs.open sql,conn,1,1
if not(rs.eof or rs.bof) then 
bsql="select * from brand where brand_id=" & rs("brand_id")
brs.open bsql,conn,1,1
%>
            <tr> 
              <td width="80">���</td>
              <td width="150" height="30"> <%
if not(brs.eof or brs.bof) then response.Write(brs("bname"))
response.Write(" -- " & rs("code"))			 
brs.close
%> </td>
              <td height="30" align="center"> <%
response.Write(rs("suit"))			 
%> </td>
              <td height="30" align="center"><input name="mount_10" id="mount_10" type="text" size="8" maxlength="8" value=1></td>
              <td height="30" align="center"><input name="price_10" id="price_10" type="text" size="8" maxlength="8" value=<%=rs("pricein")%>></td>
              <td height="30" align="center"><input name="counter_10" id="counter_10" type="text" size="8" maxlength="8" onFocus="addnum(form_administrator.mount_10.value,form_administrator.price_10.value,this);"></td>
              <td height="30" align="center"><input name="memo_10" id="memo_10" type="text" size="15" maxlength="100"></td>
            </tr>
            <%
end if
rs.close
end if
%>
            <%
if session("subs_id_11")<>0 then
sql="select * from subs where subs_id=" & session("subs_id_11")
rs.open sql,conn,1,1
if not(rs.eof or rs.bof) then 
bsql="select * from brand where brand_id=" & rs("brand_id")
brs.open bsql,conn,1,1
%>
            <tr> 
              <td width="80">�����Դ</td>
              <td width="150" height="30"> <%
if not(brs.eof or brs.bof) then response.Write(brs("bname"))
response.Write(" -- " & rs("code"))			 
brs.close
%> </td>
              <td height="30" align="center"> <%
response.Write(rs("suit"))			 
%> </td>
              <td height="30" align="center"><input name="mount_11" id="mount_11" type="text" size="8" maxlength="8" value=1></td>
              <td height="30" align="center"><input name="price_11" id="price_11" type="text" size="8" maxlength="8" value=<%=rs("pricein")%>></td>
              <td height="30" align="center"><input name="counter_11" id="counter_11" type="text" size="8" maxlength="8" onFocus="addnum(form_administrator.mount_11.value,form_administrator.price_11.value,this);"></td>
              <td height="30" align="center"><input name="memo_11" id="memo_11" type="text" size="15" maxlength="100"></td>
            </tr>
            <%
end if
rs.close
end if
%>
            <%
if session("subs_id_12")<>0 then
sql="select * from subs where subs_id=" & session("subs_id_12")
rs.open sql,conn,1,1
if not(rs.eof or rs.bof) then 
bsql="select * from brand where brand_id=" & rs("brand_id")
brs.open bsql,conn,1,1
%>
            <tr> 
              <td width="80">����</td>
              <td width="150" height="30"> <%
if not(brs.eof or brs.bof) then response.Write(brs("bname"))
response.Write(" -- " & rs("code"))			 
brs.close
%> </td>
              <td height="30" align="center"> <%
response.Write(rs("suit"))			 
%> </td>
              <td height="30" align="center"><input name="mount_12" id="mount_12" type="text" size="8" maxlength="8" value=1></td>
              <td height="30" align="center"><input name="price_12" id="price_12" type="text" size="8" maxlength="8" value=<%=rs("pricein")%>></td>
              <td height="30" align="center"><input name="counter_12" id="counter_12" type="text" size="8" maxlength="8" onFocus="addnum(form_administrator.mount_12.value,form_administrator.price_12.value,this);"></td>
              <td height="30" align="center"><input name="memo_12" id="memo_12" type="text" size="15" maxlength="100"></td>
            </tr>
            <%
end if
rs.close
end if
%>
            <%
if session("subs_id_13")<>0 then
sql="select * from subs where subs_id=" & session("subs_id_13")
rs.open sql,conn,1,1
if not(rs.eof or rs.bof) then 
bsql="select * from brand where brand_id=" & rs("brand_id")
brs.open bsql,conn,1,1
psql="select * from sort where sort_id=" & rs("sort_id")
prs.Open psql,conn,1,1
%>
            <tr> 
              <td width="80"><%=prs("sname")%></td>
              <td width="150" height="30"> <%
response.Write(brs("bname"))
response.Write(" -- " & rs("code"))			 
%> </td>
              <td height="30" align="center"> <%
response.Write(rs("suit"))			 
%> </td>
              <td height="30" align="center"><input name="mount_13" id="mount_13" type="text" size="8" maxlength="8" value=1></td>
              <td height="30" align="center"><input name="price_13" id="price_13" type="text" size="8" maxlength="8" value=<%=rs("pricein")%>></td>
              <td height="30" align="center"><input name="counter_13" id="counter_13" type="text" size="8" maxlength="8" onFocus="addnum(form_administrator.mount_13.value,form_administrator.price_13.value,this);"></td>
              <td height="30" align="center"><input name="memo_13" id="memo_13" type="text" size="15" maxlength="100"></td>
            </tr>
            <%
prs.close
brs.close
end if
rs.close
end if
%>
            <%
if session("subs_id_14")<>0 then
sql="select * from subs where subs_id=" & session("subs_id_14")
rs.open sql,conn,1,1
if not(rs.eof or rs.bof) then 
bsql="select * from brand where brand_id=" & rs("brand_id")
brs.open bsql,conn,1,1
psql="select * from sort where sort_id=" & rs("sort_id")
prs.Open psql,conn,1,1
%>
            <tr> 
              <td width="80"><%=prs("sname")%></td>
              <td width="150" height="30"> <%
response.Write(brs("bname"))
response.Write(" -- " & rs("code"))			 
%> </td>
              <td height="30" align="center"> <%
response.Write(rs("suit"))			 
%> </td>
              <td height="30" align="center"><input name="mount_14" id="mount_14" type="text" size="8" maxlength="8" value=1></td>
              <td height="30" align="center"><input name="price_14" id="price_14" type="text" size="8" maxlength="8" value=<%=rs("pricein")%>></td>
              <td height="30" align="center"><input name="counter_14" id="counter_14" type="text" size="8" maxlength="8" onFocus="addnum(form_administrator.mount_14.value,form_administrator.price_14.value,this);"></td>
              <td height="30" align="center"><input name="memo_14" id="memo_14" type="text" size="15" maxlength="100"></td>
            </tr>
            <%
prs.close
brs.close
end if
rs.close
end if
%>
            <tr> 
              <td height="35" colspan="4" align="center">&nbsp;</td>
              <td height="35" align="center">�ܽ�� </td>
              <td height="35" align="center"><input name="counter_all" id="counter_all" type="text" size="8" maxlength="9" onFocus="addnum_all(<%=session("subs_id_1")%>,<%=session("subs_id_2")%>,<%=session("subs_id_3")%>,<%=session("subs_id_4")%>,<%=session("subs_id_5")%>,<%=session("subs_id_6")%>,<%=session("subs_id_7")%>,<%=session("subs_id_8")%>,<%=session("subs_id_9")%>,<%=session("subs_id_10")%>,<%=session("subs_id_11")%>,<%=session("subs_id_12")%>,<%=session("subs_id_13")%>,<%=session("subs_id_14")%>);"></td>
              <td width="130" height="35" align="center">&nbsp;</td>
            </tr>
          </table>
          <font color="#000000" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></TD>
      </TR>
      <TR> 
        <TD height="50" colSpan=2><table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr> 
              <td width="14%" height="30"><font size="3">�ͻ�</font></td>
              <td width="43%" height="30"><font size="3"> 
                <input name="gid" type="text" size="30" maxlength="100">
                </font></td>
              <td width="43%" height="30" align="right"> <input name="pyear" id="pyear" type="text" size="5" maxlength="4" value="<%=year(now)%>">
                �� 
                <input name="pmonth" id="pmonth" type="text" size="3" maxlength="2" value="<%=month(now)%>">
                �� 
                <input name="pday" id="pday" type="text" size="3" maxlength="2" value="<%=day(now)%>">
                ��</td>
            </tr>
            <tr> 
              <td width="14%" height="30">�ͻ���ע</td>
              <td height="30" colspan="2"><textarea name="gmemo" cols="70" rows="5"></textarea></td>
            </tr>
          </table></TD>
      </TR>
      <TR> 
        <TD height="50" colSpan=2><table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr> 
              <td>����˾�����˼���ϵ�绰 
                <input name="operationer_id" type="text" id="operationer_id" size="60">
              </td>
            </tr>
          </table></TD>
      </TR>
      <TR> 
        <TD height="50" colSpan=2> <P><STRONG><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000066 size=2> 
            <INPUT type=submit id="option" name="option" value="���ɱ�׼���۵�">
			&nbsp;&nbsp;&nbsp;<input type="reset" name="cancelit" id="cancelit" value="���">
			&nbsp;&nbsp;&nbsp;<input type="button" name="cancelit" id="cancelit" value="�رմ˴���" onClick="javascript:window.close();">
            </FONT></STRONG></P>
          </TD>
      </TR>
    </TBODY>
  </TABLE>
</FORM>
<br>
<br>
</body>
</html>
