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
conn.execute "delete * from art_artist where artist_id=" & request("artist_id")
conn.execute "delete * from art_subs where artist_id=" & request("artist_id")
response.Write("<SCRIPT LANGUAGE=JavaScript> alert('ɾ���ɹ�');</SCRIPT>")
end if
%>




<%
zname=trim(request("zname"))
zname=replace(zname,"'","''")
zaddr=trim(request("zaddr"))
zaddr=replace(zaddr,"'","''")
zsex=trim(request("zsex"))
zsex=replace(zsex,"'","''")
zothers=trim(request("zothers"))
zothers=replace(zothers,"'","''")
%>

        
  <table width="100%" border="1" align="center" cellpadding="2" cellspacing="0" bordercolor="#CCCCCC">
<form name="form_administrator" id="form_administrator" method="post" action="art_artist_view.asp" onSubmit="return checkform();">
    <tr align="center" valign="middle"> 
      <td width="227" height="30">���� </td>
      <td width="222" height="30">��ַ </td>
      <td width="168">�Ա� </td>
      <td width="281">���� </td>
      <td width="81">&nbsp;</td>
    </tr>
    <tr align="center" valign="middle">
      <td height="30"><input name="zname" type="text" id="zname2" size="20" maxlength="100" value="<%=zname%>"></td>
      <td height="30"><input name="zaddr" type="text" id="zaddr2" size="20" maxlength="100" value="<%=zaddr%>"></td>
      <td><select name="zsex" id="zsex">
          <%
if zsex="" then
%>
          <option value="��" selected>��</option>
          <option value="Ů">Ů</option>
          <%
end if
%>
          <%
if zsex="��" then
%>
          <option value="��" selected>��</option>
          <option value="Ů">Ů</option>
          <%
end if
%>
          <%
if zsex="Ů" then
%>
          <option value="��">��</option>
          <option value="Ů" selected>Ů</option>
          <%
end if
%>
        </select></td>
      <td><input name="zothers" type="text" id="zothers2" size="20" maxlength="100" value="<%=zothers%>"></td>
      <td><input type="submit" name="Submit" value="��ѯ>>"></td>
    </tr>
      </form> 
  </table>













<br>
<%
sql="select * from art_artist where allname<> ''"

if zname<>"" then sql=sql & " and allname like '%"&zname&"%'"
if zaddr<>"" then sql=sql & " and addr like '%"&zaddr&"%'"
if zsex<>"" then sql=sql & " and sex like '%"&zsex&"%'"
if zothers<>"" then sql=sql & " and ( simname like '%"&zothers&"%' or nation like '%"&zothers&"%' or native like '%"&zothers&"%' or born_time like '%"&zothers&"%' or graduation like '%"&zothers&"%' or special like '%"&zothers&"%' or job_now like '%"&zothers&"%' or job_old like '%"&zothers&"%' or job_rank like '%"&zothers&"%' or expert like '%"&zothers&"%' or awarded like '%"&zothers&"%' or standest like '%"&zothers&"%' or interest like '%"&zothers&"%' or tel like '%"&zothers&"%' or fax like '%"&zothers&"%' or zip like '%"&zothers&"%' or email like '%"&zothers&"%' or memo like '%"&zothers&"%')"

sql=sql&" order by artist_ID DESC"

Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1 

if rs.eof or rs.bof then 
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' ��ã�û����Ҫ��ѯ�����ݡ���л��������Э ! ');</SCRIPT>")
response.End()
end if
if not isempty(request("page")) then   
pagecount=cint(request("page"))   
else   
pagecount=1
end if

rs.pagesize=20
rs.AbsolutePage=pagecount
%>
<table width="100%" border="1" align="center" cellpadding="2" cellspacing="0" bordercolor="#CCCCCC">
  <tr bgcolor="#FFFFFF"> 
    <td height="30" colspan="8" align="center" class="header"><font color="#FF0000" size="+1"><strong>�������б�</strong></font> 
    </td>
  </tr>
  <tr align="center"> 
    <td width="80" height="30" bgcolor="#FFFFFF">ȫ��</td>
    <td width="80" bgcolor="#FFFFFF">�ź�</td>
    <td width="40" bgcolor="#FFFFFF">�Ա�</td>
    <td bgcolor="#FFFFFF">��ַ</td>
    <td width="100" bgcolor="#FFFFFF">�绰</td>
    <td width="130" bgcolor="#FFFFFF">����</td>
  </tr>
  <%
do while not rs.eof 
%>
  <tr align="center"> 
    <td height="25" bgcolor="#FFFFFF"><%=rs("allname")%></td>
    <td bgcolor="#FFFFFF"><%=rs("simname")%>&nbsp;</td>
    <td bgcolor="#FFFFFF"><%=rs("sex")%></td>
    <td bgcolor="#FFFFFF"><%=rs("addr")%>&nbsp;</td>
    <td bgcolor="#FFFFFF"><%=rs("tel")%>&nbsp;</td>
    <td bgcolor="#FFFFFF"><a href="art_artist_viewall.asp?artist_id=<%=rs("artist_id")%>" target="_blank">��ϸ</a> | <a href="art_artist_edit.asp?artist_id=<%=rs("artist_id")%>">�༭</a> | <a href="art_artist_view.asp?artist_id=<%=rs("artist_id")%>&option=del" onClick="return confirm('ȷ��ɾ��������\n \n ��ɾ�������Ҽ�����Ʒ')">ɾ��</a> </td>
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
          <td><a href="art_artist_view.asp?page=<%=zj%>&zname=<%=zname%>&zsex=<%=zsex%>&zaddr=<%=zaddr%>&zothers=<%=zothers%>">&nbsp;<%=zj%>&nbsp;</a></td>
          <%
else
zk=zk+1
%>
        </tr>
        <tr> 
          <td><a href="art_artist_view.asp?page=<%=zj%>&zname=<%=zname%>&zsex=<%=zsex%>&zaddr=<%=zaddr%>&zothers=<%=zothers%>">&nbsp;<%=zj%>&nbsp;</a></td>
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
%>
</body>
</html>
