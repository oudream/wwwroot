<!--#include file="SHEI45FECSA.ASP" -->
<%
'开始删除操作
if request("action")="deluser" then
conn.execute "delete * from user where id=" & request("userid")
response.write"<SCRIPT language=JavaScript>alert('操作成功，您已删除了此帐号！');"
response.write"javascript:history.go(-2)</SCRIPT>"
rseponse.end
end if
%>
<%
'开始编辑操作
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
response.write"<SCRIPT language=JavaScript>alert('操作成功，您已修改此帐号资料！');"
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
<title>商城系统管理中心</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../inc/southdns.css" type="text/css">
</head>
<body bgcolor="#FFFFCC" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" cellspacing="1" cellpadding="2" bgcolor="#000000" vspace="0" hspace="0">
  <tr bgcolor="#FFCC33"> 
    <td height="27">.:: 您可以在这里进行用户相关操作</td>
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
      <td height="30" width="12%" bgcolor="#FFCC33">帐号：</td>
      <td bgcolor="#FFCC33" width="38%">&nbsp;<%=rs("uid")%></td>
      <td bgcolor="#FFCC33" width="12%">密码：</td>
      <td bgcolor="#FFCC33" width="38%">
        <input type="text" name="pwd" value="<%=rs("pwd")%>" class="form">
      </td>
    </tr>
    <tr> 
      <td height="30" bgcolor="#FFFF99">级别：</td>
      <td bgcolor="#FFFF99">
        <select name="userlevel">
          <option value="" selected>当前用户级别</option>
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
      <td bgcolor="#FFFF99">预付款：</td>
      <td bgcolor="#FFFF99">
        <input type="text" name="prepay" value="<%=rs("prepay")%>" class="form">
        &nbsp; </td>
    </tr>
    <tr> 
      <td height="30" bgcolor="#FFCC33">总支付：</td>
      <td bgcolor="#FFCC33">
<input name="sumjifen" type="text" class="form" id="sumjifen" value="<%=rs("sumjifen")%>"></td>
      <td bgcolor="#FFCC33">&nbsp;</td>
      <td bgcolor="#FFCC33">&nbsp;</td>
    </tr>
    <tr> 
      <td height="30" bgcolor="#FFCC33">性别：</td>
      <td bgcolor="#FFCC33">
        <select class=form name=sex ><option selected  value="<%=rs("sex")%>"><%=rs("sex")%></option>
                        <option  value="男">男</option>
                        <option value="女">女</option>
                      </select>
        &nbsp; </td>
      <td bgcolor="#FFCC33">年龄：</td>
      <td bgcolor="#FFCC33">
        <input type="text" name="age" value="<%=rs("age")%>" class="form">
        &nbsp; </td>
    </tr>
    <tr> 
      <td height="30" bgcolor="#FFFF99">姓名（中）：</td>
      <td bgcolor="#FFFF99">
        <input type="text" name="namez" value="<%=rs("namez")%>" class="form">
      </td>
      <td bgcolor="#FFFF99">姓名（英）：</td>
      <td bgcolor="#FFFF99">
        <input type="text" name="namee" value="<%=rs("namee")%>" class="form">
      </td>
    </tr>
    <tr> 
      <td height="30" bgcolor="#FFFF99">联系人（中）：</td>
      <td bgcolor="#FFFF99">
        <input name="contactz" type="text" class="form" id="contactz" value="<%=rs("contactz")%>">
      </td>
      <td bgcolor="#FFFF99">联系人（英）：</td>
      <td bgcolor="#FFFF99">
        <input name="contacte" type="text" class="form" id="contacte" value="<%=rs("contacte")%>">
      </td>
    </tr>
    <tr bgcolor="#FFCC33"> 
      <td height="30">国家（中）：</td>
      <td>
        <input type="text" name="govz" value="<%=rs("govz")%>" class="form">
      </td>
      <td>国家（英）：</td>
      <td> 
        <input type="text" name="gove" value="<%=rs("gove")%>" class="form">
      </td>
    </tr><tr> 
      <td height="30" bgcolor="#FFFF99">省份（中）：</td>
      <td bgcolor="#FFFF99">
        <input type="text" name="shengz" value="<%=rs("shengz")%>" class="form">
      </td>
      <td bgcolor="#FFFF99">省份（英）：</td>
      <td bgcolor="#FFFF99">
        <input type="text" name="shenge" value="<%=rs("shenge")%>" class="form">
      </td>
    </tr><tr bgcolor="#FFCC33"> 
      <td height="30">城市（中）：</td>
      <td>
        <input type="text" name="cityz" value="<%=rs("cityz")%>" class="form">
      </td>
      <td>城市（英）：</td>
      <td> 
        <input type="text" name="citye" value="<%=rs("citye")%>" class="form">
      </td>
    </tr>
	<tr> 
      <td height="30" bgcolor="#FFFF99">地址（中）：</td>
      <td bgcolor="#FFFF99">
        <input type="text" name="dizhiz" value="<%=rs("dizhiz")%>" class="form">
      </td>
      <td bgcolor="#FFFF99">地址（英）：</td>
      <td bgcolor="#FFFF99">
        <input type="text" name="dizhie" value="<%=rs("dizhie")%>" class="form">
      </td>
    </tr>
	<tr> 
      <td height="30" bgcolor="#FFFF99">&nbsp;</td>
      <td bgcolor="#FFFF99">&nbsp; </td>
      <td bgcolor="#FFFF99">联系电话2：</td>
      <td bgcolor="#FFFF99">
        <input name="tel2" type="text" class="form" id="tel2" value="<%=rs("tel2")%>">
      </td>
    </tr>
	<tr bgcolor="#FFCC33"> 
      <td height="30">邮政编码：</td>
      <td>
        <input type="text" name="postnumber" value="<%=rs("postnumber")%>" class="form">
      </td>
      <td>联系电话1：</td>
      <td> 
        <input type="text" name="tel" value="<%=rs("tel")%>" class="form">
      </td>
    </tr>
	<tr> 
      <td height="30" bgcolor="#FFFF99">E-mail：</td>
      <td bgcolor="#FFFF99">
        <input type="text" name="email" value="<%=rs("email")%>" class="form">
      </td>
      <td bgcolor="#FFFF99">传真：</td>
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
        <input type="submit" name="Submit" value="修改此帐号资料">
        <input type="hidden" name="action" value="edituser">
        <input type="hidden" name="userid" value="<%=rs("id")%>">
      </td>
    </tr>
  </form>
  <tr bgcolor="#FFCC33"> 
    <form action=EDITUSER.ASP method=post name="deluser">
      <td height="22" colspan="2" bgcolor="#FFFF99"> 
        <div align="center"> 
          <input type="submit" name="Submit2" value="删除此帐号">
          <input type="hidden" name="action" value="deluser">
          <input type="hidden" name="userid" value="<%=rs("id")%>">
        </div>
      </td>
    </form>
    <form action="USERORDERLOG.ASP" method="post" name="viewuserorderlog">
      <td colspan="2" bgcolor="#FFFF99"> 
        <input type="submit" name="Submit3" value="查看此帐号的交易资料">
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

