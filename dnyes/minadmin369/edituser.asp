<!--#include file="SHEI45FECSA.ASP" -->
<%
'��ʼɾ������
if request("action")="deluser" then
conn.execute "delete * from user where id=" & request("userid")
response.write"<SCRIPT language=JavaScript>alert('�����ɹ�������ɾ���˴��ʺţ�');"
response.write"javascript:history.go(-2)</SCRIPT>"
rseponse.end
end if
%>
<%
'��ʼ�༭����
if request("action")="edituser" then
pwd=request("pwd")
userlevel=request("userlevel")
prepay=request("prepay")
sumjifen=request("sumjifen")
sex=request("sex")
age=request("age")
namez=request("namez")
namee=request("namee")
contactz=request("contactz")
contacte=request("contacte")
govz=request("govz")
gove=request("gove")
shengz=request("shengz")
shenge=request("shenge")
cityz=request("cityz")
citye=request("citye")
dizhiz=request("dizhiz")
dizhie=request("dizhie")
tel=request("tel")
tel2=request("tel2")
oicq=request("oicq")
postnumber=request("postnumber")
fax=request("fax")

if fax="" then fax=" "
if oicq="" then oicq=" "
conn.execute "update user set pwd='"&pwd&"',userlevel='"&userlevel&"',prepay='"&prepay&"',sumjifen='"&sumjifen&"',sex='"&sex&"',age='"&age&"',namez='"&namez&"',namee='"&namee&"',contactz='"&contactz&"',contacte='"&contacte&"',govz='"&govz&"',gove='"&gove&"',shengz='"&shengz&"',shenge='"&shenge&"',cityz='"&cityz&"',citye='"&citye&"',dizhiz='"&dizhiz&"',dizhie='"&dizhie&"',tel='"&tel&"',tel2='"&tel2&"',postnumber='"&postnumber&"',oicq='"&oicq&"',fax='"&fax&"' where id=" & request("userid")
response.write"<SCRIPT language=JavaScript>alert('�����ɹ��������޸Ĵ��ʺ����ϣ�');"
response.write"javascript:history.go(-2)</SCRIPT>"
response.end
end if
%>
<%
userid=request("userid")
sql="select * from user where id=" & userid
set rs=server.createobject("ADODB.Recordset")
rs.open sql,conn,1,1
%>
<html>
<head>
<title>�̳�ϵͳ��������</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../inc/southdns.css" type="text/css">
</head>
<body bgcolor="#FFFFCC" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" cellspacing="1" cellpadding="2" bgcolor="#000000" vspace="0" hspace="0">
  <tr bgcolor="#FFCC33"> 
    <td height="27">.:: ����������������û���ز���</td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="50">&nbsp;</td>
  </tr>
</table>

<table border="0" cellspacing="1" cellpadding="4" bgcolor="#000000" width="75%" align="center">
  <form action=EDITUSER.ASP method=post name="edituser">
    <tr> 
      <td height="30" width="12%" bgcolor="#FFCC33">�ʺţ�</td>
      <td bgcolor="#FFCC33" width="38%">&nbsp;<%=rs("uid")%></td>
      <td bgcolor="#FFCC33" width="12%">���룺</td>
      <td bgcolor="#FFCC33" width="38%">
        <input type="text" name="pwd" value="<%=rs("pwd")%>" class="form">
      </td>
    </tr>
    <tr> 
      <td height="30" bgcolor="#FFFF99">����</td>
      <td bgcolor="#FFFF99">
        <select name="userlevel">
          <option value="" selected>��ǰ�û�����</option>
          <%
sqlt="select * from userlevel"
set rst=server.createobject("ADODB.Recordset")
rst.open sqlt,conn,1,1
while not rst.eof
islevel=""
if rst("userlevel")=rs("userlevel") then islevel="selected"
%> 
          <option value="<%=rst("userlevel")%>" <%=islevel%>><%=rst("userlevel")%></option>
          <%
rst.movenext
wend
rst.Close()
%> 
        </select>
      </td>
      <td bgcolor="#FFFF99">Ԥ���</td>
      <td bgcolor="#FFFF99">
        <input type="text" name="prepay" value="<%=rs("prepay")%>" class="form">
        &nbsp; </td>
    </tr>
    <tr> 
      <td height="30" bgcolor="#FFCC33">��֧����</td>
      <td bgcolor="#FFCC33">
<input name="sumjifen" type="text" class="form" id="sumjifen" value="<%=rs("sumjifen")%>"></td>
      <td bgcolor="#FFCC33">&nbsp;</td>
      <td bgcolor="#FFCC33">&nbsp;</td>
    </tr>
    <tr> 
      <td height="30" bgcolor="#FFCC33">�Ա�</td>
      <td bgcolor="#FFCC33">
        <select class=form name=sex ><option selected  value="<%=rs("sex")%>"><%=rs("sex")%></option>
                        <option  value="��">��</option>
                        <option value="Ů">Ů</option>
                      </select>
        &nbsp; </td>
      <td bgcolor="#FFCC33">���䣺</td>
      <td bgcolor="#FFCC33">
        <input type="text" name="age" value="<%=rs("age")%>" class="form">
        &nbsp; </td>
    </tr>
    <tr> 
      <td height="30" bgcolor="#FFFF99">�������У���</td>
      <td bgcolor="#FFFF99">
        <input type="text" name="namez" value="<%=rs("namez")%>" class="form">
      </td>
      <td bgcolor="#FFFF99">������Ӣ����</td>
      <td bgcolor="#FFFF99">
        <input type="text" name="namee" value="<%=rs("namee")%>" class="form">
      </td>
    </tr>
    <tr> 
      <td height="30" bgcolor="#FFFF99">��ϵ�ˣ��У���</td>
      <td bgcolor="#FFFF99">
        <input name="contactz" type="text" class="form" id="contactz" value="<%=rs("contactz")%>">
      </td>
      <td bgcolor="#FFFF99">��ϵ�ˣ�Ӣ����</td>
      <td bgcolor="#FFFF99">
        <input name="contacte" type="text" class="form" id="contacte" value="<%=rs("contacte")%>">
      </td>
    </tr>
    <tr bgcolor="#FFCC33"> 
      <td height="30">���ң��У���</td>
      <td>
        <input type="text" name="govz" value="<%=rs("govz")%>" class="form">
      </td>
      <td>���ң�Ӣ����</td>
      <td> 
        <input type="text" name="gove" value="<%=rs("gove")%>" class="form">
      </td>
    </tr><tr> 
      <td height="30" bgcolor="#FFFF99">ʡ�ݣ��У���</td>
      <td bgcolor="#FFFF99">
        <input type="text" name="shengz" value="<%=rs("shengz")%>" class="form">
      </td>
      <td bgcolor="#FFFF99">ʡ�ݣ�Ӣ����</td>
      <td bgcolor="#FFFF99">
        <input type="text" name="shenge" value="<%=rs("shenge")%>" class="form">
      </td>
    </tr><tr bgcolor="#FFCC33"> 
      <td height="30">���У��У���</td>
      <td>
        <input type="text" name="cityz" value="<%=rs("cityz")%>" class="form">
      </td>
      <td>���У�Ӣ����</td>
      <td> 
        <input type="text" name="citye" value="<%=rs("citye")%>" class="form">
      </td>
    </tr>
	<tr> 
      <td height="30" bgcolor="#FFFF99">��ַ���У���</td>
      <td bgcolor="#FFFF99">
        <input type="text" name="dizhiz" value="<%=rs("dizhiz")%>" class="form">
      </td>
      <td bgcolor="#FFFF99">��ַ��Ӣ����</td>
      <td bgcolor="#FFFF99">
        <input type="text" name="dizhie" value="<%=rs("dizhie")%>" class="form">
      </td>
    </tr>
	<tr> 
      <td height="30" bgcolor="#FFFF99">&nbsp;</td>
      <td bgcolor="#FFFF99">&nbsp; </td>
      <td bgcolor="#FFFF99">��ϵ�绰2��</td>
      <td bgcolor="#FFFF99">
        <input name="tel2" type="text" class="form" id="tel2" value="<%=rs("tel2")%>">
      </td>
    </tr>
	<tr bgcolor="#FFCC33"> 
      <td height="30">�������룺</td>
      <td>
        <input type="text" name="postnumber" value="<%=rs("postnumber")%>" class="form">
      </td>
      <td>��ϵ�绰1��</td>
      <td> 
        <input type="text" name="tel" value="<%=rs("tel")%>" class="form">
      </td>
    </tr>
	<tr> 
      <td height="30" bgcolor="#FFFF99">E-mail��</td>
      <td bgcolor="#FFFF99">
        <input type="text" name="email" value="<%=rs("email")%>" class="form">
      </td>
      <td bgcolor="#FFFF99">���棺</td>
      <td bgcolor="#FFFF99">
        <input type="text" name="fax" value="<%=rs("fax")%>" class="form">
      </td>
    </tr>
    
    <tr bgcolor="#FFCC33"> 
      <td height="30">Oicq:</td>
      <td>
        <input type="text" name="oicq" value="<%=rs("oicq")%>" class="form">
        &nbsp; </td>
      <td colspan="2"> 
        <input type="submit" name="Submit" value="�޸Ĵ��ʺ�����">
        <input type="hidden" name="action" value="edituser">
        <input type="hidden" name="userid" value="<%=rs("id")%>">
      </td>
    </tr>
  </form>
  <tr bgcolor="#FFCC33"> 
    <form action=EDITUSER.ASP method=post name="deluser">
      <td height="22" colspan="2" bgcolor="#FFFF99"> 
        <div align="center"> 
          <input type="submit" name="Submit2" value="ɾ�����ʺ�">
          <input type="hidden" name="action" value="deluser">
          <input type="hidden" name="userid" value="<%=rs("id")%>">
        </div>
      </td>
    </form>
    <form action="USERORDERLOG.ASP" method="post" name="viewuserorderlog">
      <td colspan="2" bgcolor="#FFFF99"> 
        <input type="submit" name="Submit3" value="�鿴���ʺŵĽ�������">
        <input type="hidden" name="uid" value="<%=rs("uid")%>">
      </td>
    </form>
  </tr>
</table>
<%
rs.close
set rs=nothing
%>
<!--#include file="COPYRIGHT.ASP" -->
</body>
</html>

