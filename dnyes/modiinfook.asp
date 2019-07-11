<!--#include file="mymefaq5411jkjkh/favorible/showme.asp"-->
<%on error resume next%>
<%
if session("y_it_uid")="" then
response.redirect "index.asp"
response.end
end if
%>
<%
'进行安全性监测，看数据来源是否是本服务器页面
if not instr(1,Request.ServerVariables("http_Referer"),Request.ServerVariables ("SERVER_NAME"),1)=8 then
response.write "<b>请不要从非本服务器的页面提交信息</b>"
'response.redirect "error.asp?err=001"
response.end
end if
%>
<%
'从提交表单返回值，并查看是否有非法自符
pwd=replace(request("pwd"),"'","’")
email=replace(request("email"),"'","’")
sex=request("sex")
age=replace(request("age"),"'","’")
namez=replace(request("namez"),"'","’")
namee=replace(request("namee"),"'","’")
govz=replace(request("govz"),"'","’")
gove=replace(request("gove"),"'","’")
shengz=replace(request("shengz"),"'","’")
shenge=replace(request("shenge"),"'","’")
cityz=replace(request("cityz"),"'","’")
citye=replace(request("citye"),"'","’")
dizhiz=replace(request("dizhiz"),"'","’")
dizhie=replace(request("dizhie"),"'","’")
tel=replace(request("tel"),"'","’")
oicq=replace(request("oicq"),"'","’")
postnumber=replace(request("postnumber"),"'","’")
fax=replace(request("fax"),"'","’")

if fax="" then fax=" "
if oicq="" then oicq=" "

conn.execute "update user set pwd='"&pwd&"',email='"&email&"',sex='"&sex&"',age='"&age&"',namez='"&namez&"',namee='"&namee&"',govz='"&govz&"',gove='"&gove&"',shengz='"&shengz&"',shenge='"&shenge&"',cityz='"&cityz&"',citye='"&citye&"',dizhiz='"&dizhiz&"',dizhie='"&dizhie&"',tel='"&tel&"',postnumber='"&postnumber&"',oicq='"&oicq&"',fax='"&fax&"' where uid='"&session("y_it_uid")&"'"
%>
<html>
<head>
<title><%=Application("y_it_itname")%>-修改资料</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="inc/southdns.css" type="text/css">
</head>

<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0">
<!--#include file="top.asp"--> 
<table width="750" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td width=192 valign="top">
<!--#include file="userlogin.asp"-->
<!--#include file="dynamic.asp"-->
<!--#include file="cserv.asp"-->
<!--#include file="link.asp"-->
      <br>
    </td>
    <td width="30">&nbsp;</td>
    <td valign="top" width="576"> 
      <div align="right">
        <table width="563" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td height="8" bgcolor="#FFFFFF"></td>
          </tr>
          <tr> 
            <td height="25" bgcolor="#9CCCFF" style="BORDER-RIGHT: #0079ce 1px solid; BORDER-TOP: #0079ce 1px solid; BORDER-LEFT: #0079ce 1px solid; BORDER-BOTTOM: #0079ce 1px solid"><p
                         align="center" style="margin-top:3px;">
              <div align="center">=== 修改资料 ===</div>
            </td>
          </tr>
          <tr> 
            <td><br>
              &nbsp;&nbsp;&nbsp;&nbsp; 您已经成功的修改了您的个人资料，如果您更改了您的密码，请在下次登陆时使用新密码<br><br>
              <font color="#FF0033">以下是您的新资料</font></td>
          </tr>
        </table>
        <br>
        <table width="563" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td colspan="2" height="1" bgcolor="#9CCCFF"></td>
          </tr>
          <tr> 
            <td width="1" bgcolor="#9CCCFF"></td>
            <td width="562">
              <table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
                <tr> 
                  <td height="30" width="20%">用户帐号：</td>
                  <td width="30%"><%=session("y_it_uid")%> </td>
                  <td width="20%">密码：</td>
                  <td width="30%"><%=pwd%></td>
                </tr>
                <tr> 
                  <td height="30">电子邮件：</td>
                  <td colspan="3"><%=email%></td>
                </tr>
                <tr> 
                  <td height="30">性别：</td>
                  <td><%=sex%> </td>
                  <td>年龄：</td>
                  <td><%=age%></td>
                </tr><tr> 
                  <td height="30">姓名（中文）：</td>
                  <td><%=namez%></td>
                  <td>姓名（英文）： </td>
                  <td><%=namee%></td>
                </tr>
                <tr> 
                  <td height="30">国家（中文）：</td>
                  <td><%=govz%></td>
                  <td>国家（英文）： </td>
                  <td><%=gove%></td>
                </tr><tr> 
                  <td height="30">省份（中文）：</td>
                  <td><%=shengz%></td>
                  <td>省份（英文）： </td>
                  <td><%=shenge%></td>
                </tr>
                <tr><tr> 
                  <td height="30">城市（中文）：</td>
                  <td><%=cityz%></td>
                  <td>城市（英文）： </td>
                  <td><%=citye%></td>
                </tr><tr> 
                  <td height="30">地址（中文）：</td>
                  <td><%=dizhiz%></td>
                  <td>地址（英文）： </td>
                  <td><%=dizhie%></td>
                </tr><tr> 
                  <td height="30">邮政编码：</td>
                  <td><%=postnumber%></td>
                  <td>联系电话： </td>
                  <td><%=tel%></td>
                </tr> 
                  <td height="30">Oicq：</td>
                  <td><%=oicq%></td>
                  <td>传真：</td>
                  <td><%=fax%></td>
                </tr>
                
                <tr> 
                  <td height="40" colspan="4"> 
                    <div align="center"> </div>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
          <tr> 
            <td colspan="2" height="1" bgcolor="#9CCCFF"></td>
          </tr>
        </table>
      </div>
    </td>
  </tr>
</table>
<!--#include file="copyright.asp"-->
</body>
</html>
