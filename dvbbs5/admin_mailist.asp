<!--#include file=conn.asp-->
<!-- #include file="inc/const.asp" -->
<!--#include file=inc/email.asp-->
<title><%=Forum_info(0)%>--����ҳ��</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta NAME=GENERATOR Content="Microsoft FrontPage 3.0" CHARSET=GB2312>
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0" marginheight="0" marginwidth="0" bgcolor="#DDEEFF">
<%
	if not master or session("flag")="" then
		Errmsg=Errmsg+"<br>"+"<li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��½</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
		call dvbbs_error()
	else
		dim topic,mailbody,useremail
		i=1
		set rs= server.createobject ("adodb.recordset")
		if request("action")="send" then
			call sendmail()
		else
			call mail()
		end if
	end if

sub mail()
%>
<form action="admin_mailist.asp?action=send" method=post>
<table cellpadding="2" cellspacing="1" border="0" width="95%" class="tableBorder" align=center>
                <tr> 
                  <th colspan="2">��̳�ʼ��б�
                  </th>
                </tr>
                <tr> 
                  <td width="20%" class=forumrow>ע�����</td>
		  <td class=forumrow>��������д���±��������ͣ���Ϣ�����͵�����ע��ʱ������д��������û����ʼ��б��ʹ�ý����Ĵ����ķ�������Դ��������ʹ�á�</td>
                </tr>
                <tr> 
                  <td width="20%" class=forumrow>�ʼ����⣺</td>
		  <td class=forumrow><input type=text name=topic size=60></td>
                </tr>
                <tr> 
                  <td width="20%" class=forumrow>�ʼ����ݣ�</td>
		  <td class=forumrow><textarea cols=80 rows=6 name="content"></textarea></td>
                </tr>
                <tr>  <td width="20%" class=forumrow></td>
                  <td height=20 class=forumrow>
<input type=Submit value="�� ��" name=Submit"> &nbsp; <input type="reset" name="Clear" value="�� ��">
                  </td>
                </tr>
              </table>
</form>
<%
end sub

sub sendmail()
	if request("topic")="" then
		Errmsg=Errmsg+"<br>"+"<li>�������ʼ����⡣"
		founderr=true
	else
		topic=request("topic")
	end if
	if request("content")="" then
		Errmsg=Errmsg+"<br>"+"<li>�������ʼ����ݡ�"
		founderr=true
	else
		mailbody=request("content")
	end if
	if founderr=false then
	on error resume next
	sql="select username,useremail from [user]"
	rs.open sql,conn,1,1
	if not rs.eof and not rs.bof then
		alluser=rs.recordcount
		do while not rs.eof
		if rs("useremail")<>"" then
		useremail=rs("useremail")
			if Forum_Setting(2)=0 then
				errmsg=errmsg+"<br>"+"<li>����̳��֧�ַ����ʼ���"
				exit sub
			elseif Forum_Setting(2)=1 then
				call jmail(useremail,topic,mailbody)
			elseif Forum_Setting(2)=2 then
				call Cdonts(useremail,topic,mailbody)
			elseif Forum_Setting(2)=3 then
				call aspemail(useremail,topic,mailbody)
			end if
		i=i+1
		end if
		rs.movenext
		loop
		Errmsg=Errmsg+"<br>"+"<li>�ɹ�����"&i&"���ʼ���"
	end if
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	end if
	response.write ""&Errmsg&""
end sub
%>