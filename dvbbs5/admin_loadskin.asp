<!--#include file =conn.asp-->
<!-- #include file="inc/const.asp" -->
<title><%=Forum_info(0)%>--����ҳ��</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta NAME=GENERATOR Content="Microsoft FrontPage 3.0" CHARSET=GB2312>
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0" marginheight="0" marginwidth="0" bgcolor="#DDEEFF">
  <script language="JavaScript">
<!--
function CheckAll(form)  {
  for (var i=0;i<form.elements.length;i++)    {
    var e = form.elements[i];
    if (e.name != 'chkall')       e.checked = form.chkall.checked; 
   }
  }
//-->
</script>
 <table border="0"  cellspacing="1" cellpadding="5" height="1" align=center width="95%" class="tableBorder">
    <tr>
      <th  colspan="3" align="center"><a href=?><font color=#ffffff>��̳ģ�浼������</font></a> | <a href=?action=load><font color=#ffffff>��̳ģ�浼�빦��</font></a></th>
    </tr>
  <tr >
    <td   align=center class="forumrow">
      <%
if not master or session("flag")="" then
	Errmsg=Errmsg+"<br>"+"<li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��½</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
	call dvbbs_error()
else
             if request("action")="inputskin" then
                        call inputskin()
                elseif request("action")="loadskin" then
                        call loadskin()
                elseif request("action")="load" then
                        call load()
                elseif request("action")="rename" then
                        call rename()
                elseif request("action")="savenm" then
                        call savenm()
                else
                    
                        CALL MAIN()
                end if
end if
%>
</td>
</tr>
    <tr >
      <th height="1" colspan="3" align="center"><a href=?><font color=#ffffff>��̳ģ�浼������</font></a> | <a href=?action=load><font color=#ffffff>��̳ģ�浼�빦��</font></a></th>
    </tr>
  </table>

</body>
</html>
<%SUB MAIN()
dim sname,act,mdbname,tconn
if request("action")="loadthis" then
    sname="����"
    act="loadskin"
    mdbname=Checkstr(trim(request.form("skinmdb")))
    if mdbname="" then
     		ErrMsg=ErrMsg+"<Br>"+"<li>����д����ģ�汣��ı���"
   		response.write Errmsg
            exit sub
end if
else
    sname="����"
    act="inputskin"
end if     

%>
<table border="0"  cellspacing="1" cellpadding="5" height="1" align=center width="95%" class="tableBorder">
  <tr >
    <td width="100%" height="1" colspan="4" align="center" class="forumrow">
<B><%=sname%>��̳ģ���б�</B></td>
  </tr>
     <tr >
      <td width="10%"  align="center" class="forumrow">���</td>
      <td width="70%"  align="center" class="forumrow">ģ������</td>
      <td width="10%"  align="center" class="forumrow">����</td>
      <td width="10%"  align="center" class="forumrow">ѡ��</td>
    </tr>
 <form action="?action=<%=act%>" method=post name=even>
<%
if act="loadskin" then
        
       	Set tconn = Server.CreateObject("ADODB.Connection")
	tconn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(mdbname)
 set rs=tconn.execute(" select id,Forum_Body,SkinName,active from dvskin  order by id ")
        else
 set rs=conn.execute(" select id,Forum_Body,SkinName,active from config  order by id ")
end if
		do while not rs.eof
response.write "<tr "
   if rs("active")=1 then
response.write "bgcolor="&Forum_body(2)&" "
else
response.write "bgcolor="&Forum_body(2)&" "
end if
response.write ">"
%>

      <td class="forumrow"><%=rs("id")%></td>
      <td class="forumrow"><%=rs("SkinName")%></td>
      <td class="forumrow" align=center><a href="?action=rename&act=<%=act%>&skid=<%=rs("id")%>&mdbname=<%=mdbname%>" >����</a></td>
      <td class="forumrow"><p><input type="checkbox" name="skid" value="<%=rs("id")%>"><input TYPE="hidden" NAME="active" VALUE="<% =rs("active") %>"></p></td>
    </tr>
<%		rs.movenext
		loop
     rs.close
     set rs=nothing
%>
  <tr >
    <td  colspan="4" align=center class="forumrow">

<%=sname%>�����ݿ⣺<input type="text" name="skinmdb" size="30" value="skin/skin.mdb">
<input type="submit" name="submit" value="<%=sname%>"> <input type=submit name=Submit value=ɾ��  onclick="{if(confirm('ע�⣺��ɾ����ģ�潫���ָܻ���')){this.document.even.submit();return true;}return false;}">  <input type=checkbox name=chkall value=on onclick="CheckAll(this.form)">ȫѡ</td>
  </tr>
</form>
  <tr >
    <td class="forumrow" colspan="4" align=left>
ע��<br>
����ȷ��ģ�����ݿ�����ȷ��<br>
������ģ�����ݿ����skinĿ¼�£�����д��skin/skin.mdb��<br>
����ģ�����ݿ��ڱ��ݵı���ΪDVSKIN,�벻Ҫ���ģ�<br>
</td>
  </tr>

</table>
<% END SUB 


sub inputskin()
dim skid,mdbname,TCONN

	skid=checkstr(request("skid"))
   	mdbname=Checkstr(trim(request.form("skinmdb")))
if skid="" or  isnull(skid) then
     		ErrMsg=ErrMsg+"<Br>"+"<li>����δѡȡҪ������ģ��"
   		response.write Errmsg
            exit sub
end if
if mdbname="" then
     		ErrMsg=ErrMsg+"<Br>"+"<li>������д����ģ�����ݿ���"
   		response.write Errmsg
            exit sub
end if
response.write "ģ��ɣģ�"&skid


if request("submit")="ɾ��" then
	set rs=conn.execute("select skinname from config where active=1 and id in ("&skid&")")
	if rs.eof and rs.bof then
	conn.execute("delete from config where id in ("&skid&")")
	response.write "<br>ɾ���ɹ���"
	else
	response.write "<br>������ʾ����ǰʹ�õ���̳ģ�岻��ɾ�����������ø�ģ��Ϊ���ǵ�ǰģ��״̬��"
	end if
	rs.close
	set rs=nothing

else
    
       	Set tconn = Server.CreateObject("ADODB.Connection")
	tconn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(mdbname)
        dim trs,nrs
     set rs=conn.execute(" select Forum_Body,SkinName from config where id in ("&skid&")  order by id ")
		do while not rs.eof
     tconn.execute("insert into dvskin (Forum_Body,SkinName) values ('"&rs("Forum_Body")&"','"&rs("SkinName")&"')")

		rs.movenext
		loop
     rs.close
     set rs=nothing


	if err.number<>0 then
		ErrMsg=ErrMsg+"<Br>"+"<li> ���ݿ����ʧ�ܣ����Ժ����ԣ�"&err.Description
		err.clear
		response.write Errmsg
		exit sub
	else
		response.write "���ݵ����ɹ���<br>"
	end if
end if

end sub


sub load()
%>
<table border="0"  cellspacing="1" cellpadding="5" height="1" align=center width="95%" class="tableBorder">
  <tr >
      <th  colspan="2">
        <p align="center">����ģ������</th>
    </tr>
 <form action="?action=loadthis" method=post >
    <tr >
      <td width="20%" class="forumrow">����ģ�����ݿ�����</td>
      <td width="80%" class="forumrow"><input type="text" name="skinmdb" size="30" value="skin/skin.mdb"></td>
    </tr>
    <tr >
      <th  colspan="2">
        <p align="center"><input type="submit" name="submit" value="��һ��"></th>
    </tr>
</FORM>
  </table>
<%
end sub

sub loadskin()
dim skid,mdbname,TCONN

	skid=checkstr(request("skid"))
   	mdbname=Checkstr(trim(request.form("skinmdb")))
if skid="" or  isnull(skid) then
     		ErrMsg=ErrMsg+"<Br>"+"<li>����δѡȡҪ�����ģ��"
   		response.write Errmsg
            exit sub
end if
if mdbname="" then
     		ErrMsg=ErrMsg+"<Br>"+"<li>����д����ģ�����ݿ���"
   		response.write Errmsg
            exit sub
end if
response.write "ģ��ɣģ�"&skid

     dim trs,ars
       	Set tconn = Server.CreateObject("ADODB.Connection")
	tconn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(mdbname)
   
if request("submit")="ɾ��" then
	tconn.execute("delete from dvskin where id in ("&skid&")")
	response.write "ɾ���ɹ���"

else

     set trs=tconn.execute(" select Forum_Body,SkinName from dvskin where id in ("&skid&")  order by id ")
		do while not trs.eof
     set Ars=conn.execute("select * from config where active=1")

     set rs = server.CreateObject ("adodb.recordset")
     sql="select * from config"
rs.open sql,conn,1,3
rs.addnew
rs("skinname")=trs("skinname")
rs("Forum_body")=trs("Forum_body")
rs("Forum_info")=ars("Forum_info")
rs("Forum_copyright")=ars("Forum_copyright")
rs("Forum_Setting")=ars("Forum_Setting")
rs("StopReadme")=ars("StopReadme")
rs("Forum_upload")=ars("Forum_upload")
rs("Forum_ubb")=ars("Forum_ubb")
rs("Forum_statepic")=ars("Forum_statepic")
rs("Forum_topicpic")=ars("Forum_topicpic")
rs("Forum_boardpic")=ars("Forum_boardpic")
rs("Forum_pic")=ars("Forum_pic")
rs("Forum_ads")=ars("Forum_ads")
rs("Forum_user")=ars("Forum_user")
rs("badwords")=ars("badwords")
rs("splitwords")=ars("splitwords")
rs("birthuser")=ars("birthuser")
rs("Maxonline")=ars("Maxonline")
rs("MaxonlineDate")=ars("MaxonlineDate")
rs("LastPost")="$$"&Now()&"$$$$"
rs("allposttable")=ars("allposttable")
rs("allposttablename")=ars("allposttablename")
rs("NowUseBBS")=ars("NowUseBBS")
rs.update
rs.close
set rs=nothing

		trs.movenext
		loop
     trs.close
     set trs=nothing
	if err.number<>0 then
		ErrMsg=ErrMsg+"<Br>"+"<li> ���ݿ����ʧ�ܣ����Ժ����ԣ�"&err.Description
		err.clear
		response.write Errmsg
		exit sub
	else
		response.write "���ݵ���ɹ���<br>"
	end if
end if

end sub

sub rename()
dim skid,mdbname,tconn
dim srs
	skid=checkstr(request("skid"))
        mdbname=Checkstr(trim(request("mdbname")))
if request("act")="loadskin" then      
       	Set tconn = Server.CreateObject("ADODB.Connection")
	tconn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(mdbname)
 set srs=tconn.execute(" select id,Forum_Body,SkinName from dvskin where id="&skid&" ")
        else
 set srs=conn.execute(" select id,Forum_Body,SkinName from config where id="&skid&" ")
end if
%>
<table border="0"  cellspacing="1" cellpadding="5" height="1" align=center width="95%" class="tableBorder">
  <tr >
      <th   colspan="2">
        <p align="center">����ģ������ ID=<%=srs("ID")%></td>
    </tr>
 <form action="?action=savenm" method=post >
    <tr >
      <td width="20%" class="forumrow">ģ��ԭ����</td>
      <td width="80%" class="forumrow"><%=srs("SkinName")%></td>
    </tr>
   <tr >
      <td width="20%" class="forumrow">ģ��������</td>
      <td width="80%" class="forumrow"><input type="text" name="skinNAME" size="30" value=""></td>
    </tr>
    <tr>
      <th   colspan="2">
        <p align="center"><input type="submit" name="submit" value="����"></th>
    </tr>
<% if request("act")="loadskin" then
%><input TYPE="hidden" NAME="mdbname" VALUE="<% =mdbname %>">
<% end if %>
<input TYPE="hidden" NAME="skid" VALUE="<% =srs("id") %>">
<input TYPE="hidden" NAME="act" VALUE="<% =request("act") %>">
</FORM>
  </table>
<%

srs.close
set srs=nothing
end sub

sub savenm()
dim skid,mdbname,tconn,skinNAME
	skid=checkstr(request("skid"))
        mdbname=Checkstr(trim(request("mdbname")))
        skinNAME=Checkstr(trim(request.form("skinname")))
if request("act")="loadskin" then
        
       	Set tconn = Server.CreateObject("ADODB.Connection")
	tconn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(mdbname)
 tconn.execute(" UPDATE dvskin set SkinName='"&skinNAME&"'  where id="&skid&" ")
        else
 conn.execute(" UPDATE config set SkinName='"&skinNAME&"'  where id="&skid&" ")
end if

response.write "���ݸ��³ɹ���<br>"

end SUB
%>