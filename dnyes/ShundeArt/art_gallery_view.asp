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
    <td>&nbsp;</td>
  </tr>
</table>

<!--#include file="art_adminconn.asp" -->




<%
if trim(request("option"))="del" then
conn.execute "delete * from art_gallery where art_gallery_id=" & request("art_gallery_id")
response.Write("<SCRIPT LANGUAGE=JavaScript> alert('ɾ���ɹ�');</SCRIPT>")
end if
%>









<%
zgname=trim(request("zgname"))
zgname=replace(zgname,"'","''")

zaddr=trim(request("zaddr"))
zaddr=replace(zaddr,"'","''")

ztel=trim(request("ztel"))
ztel=replace(ztel,"'","''")

zfax=trim(request("zfax"))
zfax=replace(zfax,"'","''")

zprincipal=request("zprincipal")
zprincipal=request("zprincipal")

zothers=trim(request("zothers"))
zothers=replace(zothers,"'","''")
%>








        
  <table width="100%" border="1" align="center" cellpadding="2" cellspacing="0" bordercolor="#CCCCCC">
<form name="form_administrator" id="form_administrator" method="post" action="art_gallery_view.asp">
    <tr align="center" valign="middle"> 
      <td height="30">���ƣ� 
        <input name="zgname" type="text" id="zgname" value="<%=zgname%>" size="20" maxlength="100"></td>
      <td height="30"> ��ַ�� 
        <input name="zaddr" type="text" id="zaddr" size="20" maxlength="100" value="<%=zaddr%>"></td>
      <td> �绰�� 
        <input name="ztel" type="text" id="ztel" size="20" maxlength="100" value="<%=ztel%>"></td>
      <td rowspan="2"><input type="submit" name="Submit" value="��ѯ>>"></td>
    </tr>
    <tr align="center" valign="middle"> 
      <td height="30">���棺 
        <input name="zfax" id="zfax" type="text" size="20" maxlength="100" value="<%=zfax%>">
      </td>
      <td height="30">�᳤�� 
        <input name="zprincipal" type="text" id="zprincipal" size="20" maxlength="100" value="<%=zprincipal%>"></td>
      <td>������ 
        <input name="zothers" type="text" id="zothers" size="20" maxlength="100" value="<%=zothers%>"></td>
    </tr>
       </form>
 </table>

<br>
<%
Set rs=Server.CreateObject("ADODB.RecordSet") 
sql="select * from art_gallery where gname<>'' " 

if zgname<>"" then sql=sql & " and gname like '%"&zgname&"%'"
if zaddr<>"" then sql=sql & " and addr like '%"&zaddr&"%'"
if ztel<>"" then sql=sql & " and tel like '%"&ztel&"%'"
if zfax<>"" then sql=sql & " and fax like '%"&zfax&"%'"
if zprincipal<>"" then sql=sql & " and principal like '%"&zprincipal&"%'"
if zothers<>"" then sql=sql & " and ( zip like '%"&zothers&"%' or email like '%"&zothers&"%' or web_addr like '%"&zothers&"%' or creation_date like '%"&zothers&"%' or briefintro like '%"&zothers&"%' or memo like '%"&zothers&"%')"

sql = sql &" order by art_gallery_id DESC"

rs.open sql,conn,1,1 

if rs.eof or rs.bof then 
%>
<table width="97%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="20" align="center">��������Ҫ������!</td>
  </tr>
</table>
<%
response.End()
end if
if not isempty(request("page")) then   
pagecount=cint(request("page"))   
else   
pagecount=1
end if

rs.pagesize=15
rs.AbsolutePage=pagecount
%>



<table width="100%" border="1" align="center" cellpadding="2" cellspacing="0" bordercolor="#CCCCCC">
  <tr bgcolor="#FFFFFF"> 
    <td height="30" colspan="9" align="center" class="header"><font color="#FF0000" size="+1"><strong>�й������б�</strong></font></td>
  </tr>
  <tr align="center"> 
    <td width="100" bgcolor="#FFFFFF">����</td>
    <td bgcolor="#FFFFFF">��ַ</td>
    <td width="100" bgcolor="#FFFFFF">�绰</td>
    <td width="100" bgcolor="#FFFFFF">����</td>
    <td width="130" bgcolor="#FFFFFF">��ַ</td>
    <td width="100" bgcolor="#FFFFFF">������</td>
    <td width="150" bgcolor="#FFFFFF">����</td>
  </tr>
  <%
do while not rs.eof 
%>
  <tr align="center"> 
    <td height="25" bgcolor="#FFFFFF"><%=rs("gname")%></td>
    <td bgcolor="#FFFFFF"><%=rs("addr")%>&nbsp;</td>
    <td bgcolor="#FFFFFF"><%=rs("tel")%>&nbsp;</td>
    <td bgcolor="#FFFFFF"><%=rs("fax")%>&nbsp;</td>
    <td bgcolor="#FFFFFF"><%=rs("website_addr")%>&nbsp;</td>
    <td bgcolor="#FFFFFF"><%=rs("principal")%>&nbsp;</td>
    <td width="200" bgcolor="#FFFFFF"><a href="art_gallery_viewall.asp?art_gallery_id=<%=rs("art_gallery_id")%>" target="_blank">��ϸ</a>&nbsp; | &nbsp;<a href="art_gallery_edit.asp?art_gallery_id=<%=rs("art_gallery_id")%>">�༭</a>&nbsp; | &nbsp;<a href="art_gallery_view.asp?art_gallery_id=<%=rs("art_gallery_id")%>&option=del" onClick="return confirm('ȷ��ɾ��������')">ɾ��</a>&nbsp;</td>
  </tr>
  <%
rs.movenext
i=i+1                                                                     
if i>=rs.pagesize then exit do                                                           
loop
rs.movefirst
%>
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
          <td><a href="art_gallery_view.asp?zgname=<%=zgname%>&zaddr=<%=zaddr%>&ztel=<%=ztel%>&zfax=<%=zfax%>&zcap=<%=zcap%>&zothers=<%=zothers%>&page=<%=zj%>">&nbsp;<%=zj%>&nbsp;</a></td>
          <%
else
zk=zk+1
%>
        </tr>
        <tr> 
          <td><a href="art_gallery_view.asp?zgname=<%=zgname%>&zaddr=<%=zaddr%>&ztel=<%=ztel%>&zfax=<%=zfax%>&zcap=<%=zcap%>&zothers=<%=zothers%>&page=<%=zj%>">&nbsp;<%=zj%>&nbsp;</a></td>
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