<script language="JavaScript">

function checkform()
{
	if(! isNumber(form_administrator.zbig.value)) {
		alert('����������д');
		form_administrator.zbig.focus();
		return false;
		}
	if(! isNumber(form_administrator.zsmall.value)) {
		alert('����������д');
		form_administrator.zsmall.focus();
		return false;
		}
	return true;
}

function isNumber(name)
{
	for(i = 0; i < name.length; i++) {
		if (!((name.charAt(i) >= "0" && name.charAt(i) <= "9") || name.charAt(i) == "."))
			return false;
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
</script>

<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=site.css rel=stylesheet>
<title></title>
</head>
<body topmargin="0">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="20">&nbsp;</td>
  </tr>
</table>

<!--#include file="art_adminconn.asp" -->



<%
if trim(request("option"))="del" then
conn.execute "delete * from art_subs where subs_id=" & request("subs_id")
response.Write("<SCRIPT LANGUAGE=JavaScript> alert('ɾ���ɹ�');</SCRIPT>")
end if
%>



<%
zname=trim(request("zname"))
zname=replace(zname,"'","''")
zstate=trim(request("zstate"))
zstate=replace(zstate,"'","''")
zsproperties=trim(request("zsproperties"))
zsproperties=replace(zsproperties,"'","''")

zbig=request("zbig")
zsmall=request("zsmall")

zothers=trim(request("zothers"))
zothers=replace(zothers,"'","''")
%>








        
  <table width="100%" border="1" align="center" cellpadding="2" cellspacing="0" bordercolor="#CCCCCC">
<form name="form_administrator" id="form_administrator" method="post" action="art_subs_view.asp" onSubmit="return checkform();">
    <tr align="center" valign="middle"> 
      <td width="243" height="30">���� 
        <input name="zname" type="text" id="zname" size="20" maxlength="100" value="<%=zname%>"></td>
      <td width="139" height="30">״̬ 
		  <select name="zstate" id="zstate">
          <%
if zstate="" then
%>
            <option value="�ѳ���">�ѳ���</option>
            <option value="δ����" selected>δ����</option>
            <option value="�ղ�">�ղ�</option>
          <%
end if
%>
          <%
if zstate="�ѳ���" then
%>
            <option value="�ѳ���" selected>�ѳ���</option>
            <option value="δ����">δ����</option>
            <option value="�ղ�">�ղ�</option>
          <%
end if
%>
          <%
if zstate="δ����" then
%>
            <option value="�ѳ���">�ѳ���</option>
            <option value="δ����" selected>δ����</option>
            <option value="�ղ�">�ղ�</option>
          <%
end if
%>
          <%
if zstate="�ղ�" then
%>
            <option value="�ѳ���">�ѳ���</option>
            <option value="δ����">δ����</option>
            <option value="�ղ�" selected>�ղ�</option>
          <%
end if
%>
        </select></td>
      <td width="258"> ������ 
        <input name="zsproperties" type="text" id="zsproperties" size="20" maxlength="100" value="<%=zsproperties%>"></td>
      <td width="112" rowspan="2"><input type="submit" name="Submit" value="��ѯ>>"></td>
    </tr>
    <tr align="center" valign="middle"> 
      <td height="30" colspan="2">��Ʒ�۸���� 
        <input name="zbig" id="zbig" type="text" size="10" maxlength="8" value="<%=zbig%>">
        ��Ʒ�۸�С�� 
        <input name="zsmall" id="zsmall" type="text" size="10" maxlength="8" value="<%=zsmall%>"></td>
      <td>�������� 
        <input name="zothers" type="text" id="zothers" size="20" maxlength="100" value="<%=zothers%>"></td>
    </tr>
      </form>
  </table>








<%
Set rs=Server.CreateObject("ADODB.RecordSet")
Set srs=Server.CreateObject("ADODB.RecordSet") 
%>
<br>
<table width="100%" border="1" align="center" cellpadding="2" cellspacing="0" bordercolor="#CCCCCC">
  <%
sql="select * from art_subs where sname<> ''"

if zname<>"" then sql=sql & " and sname like '%"&zname&"%'"
if zsproperties<>"" then sql=sql & " and sproperties like '%"&zsproperties&"%'"
if zstate<>"" then sql=sql & " and sstate like '%"&zstate&"%'"
if zbig<>"" then sql=sql & " and price >= " & zbig
if zsmall<>"" then sql=sql & " and price <= " & zsmall
if zothers<>"" then sql=sql & " and ( subs_id like '%"&zothers&"%' or simg like '%"&zothers&"%' or brief like '%"&zothers&"%' or memo like '%"&zothers&"%')"

sql = sql &" order by subs_id DESC"

rs.open sql,conn,1,1 

if rs.eof or rs.bof then 
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' �����߻�û���κ���Ʒ ! ');</SCRIPT>")
%>
  <tr> 
    <td colspan="8" align="center" class="header">&nbsp;</td>
  </tr>
  <%
response.End()
end if
if not isempty(request("page")) then   
pagecount=cint(request("page"))   
else   
pagecount=1
end if

rs.pagesize=10
rs.AbsolutePage=pagecount
%>
  <tr bgcolor="#FFFFFF"> 
    <td colspan="8" align="center" class="header"><a href="art_subs_view.asp"><font color="#FF0000" size="+1"><strong>��Ʒ�鿴</strong></font></a> 
    </td>
  </tr>
  <tr align="center"> 
    <td height="30" bgcolor="#FFFFFF">��Ʒ����</td>
    <td width="80" bgcolor="#FFFFFF">����(ȫ��)</td>
    <td width="100" bgcolor="#FFFFFF">������</td>
    <td width="100" bgcolor="#FFFFFF">�г���</td>
    <td width="50" bgcolor="#FFFFFF">״̬</td>
    <td width="150" bgcolor="#FFFFFF">����</td>
  </tr>
  <%
do while not rs.eof 
%>
  <tr align="center"> 
    <td height="30" bgcolor="#FFFFFF"><%=rs("sname")%>&nbsp;</td>
    <td bgcolor="#FFFFFF"> <%
ssql="select * from art_artist where artist_id="&rs("artist_id") 
srs.open ssql,conn,1,1 
if not(srs.eof or srs.bof) then response.Write(srs("allname")) 
srs.close
	%> &nbsp;</td>
    <td bgcolor="#FFFFFF"><%=rs("sproperties")%>&nbsp;</td>
    <td bgcolor="#FFFFFF"><%=rs("price")%>&nbsp;</td>
    <td bgcolor="#FFFFFF"><%=rs("sstate")%>&nbsp;</td>
    <td bgcolor="#FFFFFF"><a href="art_subs_viewall.asp?subs_id=<%=rs("subs_id")%>" target="_blank">��ϸ</a>&nbsp; 
      | &nbsp;<a href="art_subs_edit.asp?subs_id=<%=rs("subs_id")%>">�༭</a>&nbsp; 
      | &nbsp;<a href="art_subs_view.asp?subs_id=<%=rs("subs_id")%>&option=del" onClick="return confirm('ȷ��ɾ��������')">ɾ��</a>&nbsp;</td>
  </tr>
  <%
rs.movenext
i=i+1                                                                     
if i>=rs.pagesize then exit do                                                           
loop
rs.movefirst
%>
  <tr bgcolor="#FFFFFF"> 
    <td height="35" colspan="8">&nbsp;</td>
  </tr>
</table>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="70%" height="30"> <table border="0" align="center" cellpadding="0" cellspacing="0">
        <tr> 
          <%
zj=1
zk=1
for zi=1 to rs.pagecount
if zj<=zk*15 then
%>
          <td><a href="art_subs_view.asp?zname=<%=zname%>&zstate=<%=zstate%>&zsproperties=<%=zsproperties%>&zbig=<%=zbig%>&zsmall=<%=zsmall%>&zothers=<%=zothers%>&page=<%=zj%>">&nbsp;<%=zj%>&nbsp;</a></td>
          <%
else
zk=zk+1
%>
        </tr>
        <tr> 
          <td><a href="art_subs_view.asp?zname=<%=zname%>&zstate=<%=zstate%>&zsproperties=<%=zsproperties%>&zbig=<%=zbig%>&zsmall=<%=zsmall%>&zothers=<%=zothers%>&page=<%=zj%>">&nbsp;<%=zj%>&nbsp;</a></td>
          <%
end if
zj=zj+1
next
%>
        </tr>
      </table></td>
    <td width="33%" align="left"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td align="right"> Record :<font color=red><b><%=rs.recordcount%></b></font> 
            Page <font color=red><b><%=pagecount%></b></font> of <font color=red><b><%=rs.pagecount%></b></font> 
          </td>
        </tr>
      </table></td>
  </tr>
</table>
<%
rs.close
set rs=nothing
set srs=nothing
%>
</body>
</html>
