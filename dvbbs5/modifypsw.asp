<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!--#include file="md5.asp"-->
<!--#include file="inc/chkinput.asp"-->
<!--#include file="inc/usercp.asp"-->
<%
'=========================================================
' File: modifypsw.asp
' Version:5.0
' Date: 2002-9-11
' Script Written by satan
'=========================================================
' Copyright (C) 2001,2002 AspSky.Net. All rights reserved.
' Web: http://www.aspsky.net,http://www.dvbbs.net
' Email: info@aspsky.net,eway@aspsky.net
'=========================================================

response.buffer=true
stats="�޸�����"
dim psw
dim password
dim oldpassword
dim quesion
dim answer
dim usercookies

call nav()
if Cint(GroupSetting(16))=0 then
	Errmsg=Errmsg+"<br>"+"<li>��û���ڱ���̳�޸����ϵ�Ȩ�ޣ���<a href=login.asp>��½</a>����ͬ����Ա��ϵ��"
	founderr=true
end if

if not founduser then
	Errmsg=Errmsg+"<br>"+"<li>����û�е�½�����½������޸ġ�"
	founderr=true
end if

if founderr then
  	errmsg=errmsg+"<br>"+"<li>��û��<a href=login.asp target=_blank>��¼</a>��"
	call head_var(1)
	call dvbbs_error()
else
	call cp_var()
	if request("action")="updat" then
		call update()
		if founderr then
			call dvbbs_error()
		else
			sucmsg="<li>�޸�����ɹ���"
			call dvbbs_suc()
		end if
	else
		call userinfo()
	end if
end if
call activeonline()
call footer()

sub userinfo()
set rs=server.createobject("adodb.recordset")
sql="Select * from [User] where userid="&userid
rs.open sql,conn,1,1
if rs.eof and rs.bof then
	errmsg=errmsg+"<br>"+"<li>���û��������ڡ�"
	founderr=true
	exit sub
else
%>
<form action="modifypsw.asp?action=updat&username=<%=htmlencode(membername)%>" method=POST name="theForm">
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
<tr> 
      <th colspan="2" width="100%">�û���������
      </th>
    </tr>
<tr>    
        <td width="40%" class=tablebody1><B>������ȷ��</B>��<BR>��Ҫ�޸���������������ȷ��</td>   
        <td width="60%" class=tablebody1>    
          <input type="password" name="oldpsw" value="" size=30 maxlength=13>   
        </td>   
      </tr>  
    <tr>    
        <td width="40%" class=tablebody1><B>������ȷ��</B>��<BR>��Ҫ�޸���ֱ������������������</td>   
        <td width="60%" class=tablebody1>    
          <input type="password" name="psw" value="" size=30 maxlength=13>   
        </td>   
      </tr>   
      <tr>    
        <td width="40%" class=tablebody1><B>��������</B>��<BR>����д��Ϊ���������</td>   
        <td width="60%" class=tablebody1>    
          <input type=text name="quesion" size=30 value="<%=htmlencode(rs("quesion"))%>">
        </td>   
      </tr>  
      <tr>    
        <td width="40%" class=tablebody1><B>�����</B>��<BR>��������д�Ա����պ�ȡ������<BR>�𰸲�����MD5���ܣ�ֻ��ȡ������ʹ�ã���Ҫ�޸Ŀ�ֱ����д��</td>   
        <td width="60%" class=tablebody1>    
          <input type=text name="answer" size=30 value="<%=htmlencode(rs("answer"))%>">
          <input type=hidden name="oldanswer" value="<%=htmlencode(rs("answer"))%>">
        </td>   
      </tr>
    <tr align="center"> 
      <td colspan="2" width="100%"  class=tablebody2>
            <input type=Submit value="�� ��" name="Submit"> &nbsp; <input type="reset" name="Submit2" value="�� ��">
      </td>
    </tr>

    </table></form>
   
</body>
</html> 
<%
end if
rs.close
set rs=nothing
end sub

sub update()
set rs=server.createobject("adodb.recordset")
sql="Select userpassword from [User] where userid="&userid
set rs=conn.execute(sql)
if rs.eof and rs.bof then
	errmsg=errmsg+"<br>"+"<li>���û��������ڡ�"
	founderr=true
	exit sub
else
	if request("oldpsw")="" then
	  	errmsg=errmsg+"<br>"+"<li>���������ľ�����,��������޸ġ�"
		founderr=true
		exit sub
	elseif md5(trim(request("oldpsw")))<>trim(rs("userpassword")) then
	  	errmsg=errmsg+"<br>"+"<li>����ľ�����������������롣"
		founderr=true
		exit sub
	else
		oldpassword=request("oldpsw")
	end if
				
	if request("psw")<>"" then
		password=md5(request("psw"))
	else
		password=rs("userpassword")
	end if
	if request("quesion")="" then
	  	errmsg=errmsg+"<br>"+"<li>������������ʾ���⡣"
		founderr=true
		exit sub
	else
		quesion=request("quesion")
	end if
	if request("answer")="" then
	  	errmsg=errmsg+"<br>"+"<li>������������ʾ����𰸡�"
		founderr=true
		exit sub
	elseif request("answer")=request("oldanswer") then
		answer=request("answer")
	else
		answer=md5(request("answer"))
	end if
end if
set rs=server.createobject("adodb.recordset")
sql="Select * from [User] where userid="&userid
rs.open sql,conn,1,3
if rs.eof and rs.bof then
	errmsg=errmsg+"<br>"+"<li>���û��������ڡ�"
	founderr=true
	exit sub
else
	rs("userpassword")=password
	
	rs("quesion")=quesion
	rs("answer")=answer
	rs.Update
	Response.Cookies("aspsky")("password") = password
	Response.Cookies("aspsky").path=cookiepath
end if
rs.close
set rs=nothing
end sub
%>   