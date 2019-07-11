<!--#include file="mymefaq5411jkjkh/favorible/showme.asp"-->
<%on error resume next%>
<%
if request("appeal")="True" then

inBillNo=replace(request("inBillNo"),"'","’")
username=replace(request("username"),"'","’")

'开始向数据库写入数据，并检测是否已有此订单号是否已被申诉
'检测是否有此订单
set rs=Server.CreateObject("ADODB.Recordset")
sql="select * from orders where username= '" & username & "' and inBillNo='"&inBillNo&"'"
rs.open sql,conn,1,1
if rs.bof or rs.eof then
response.redirect "error.asp?err=010"
response.end
end if
rs.close
set rs=nothing

'检测是否已被申诉过
set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from usererror where inBillNo= '" & inBillNo & "'"
rs.open sql,conn,1,1

if not (rs.Bof or rs.eof) then
rs.Close
set rs=nothing
response.redirect "error.asp?err=009"
response.end
else
sql="insert into usererror(username,inBillNo,sdate) values ('"&username&"','"&inBillNo&"','"&now()&"')"
conn.Execute(sql)
rs.close
set rs=nothing
response.redirect "appeal.asp?ok=ok"
end if
end if
%>
<%
if session("y_it_uid")="" then
response.redirect "index.asp"
response.end
end if
%>
<html>
<head>
<title><%=Application("y_it_itname")%>-申诉定单</title>
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
      <br>
    </td>
    <td width="30">&nbsp;</td>
    <td valign="top" width="576"> 
      <div align="right">
        <table width="563" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td height="10"></td>
          </tr>
          <tr bgcolor="#BFCEDD"> 
            <td height="25" bgcolor="#9CCCFF" style="BORDER-RIGHT: #0079ce 1px solid; BORDER-TOP: #0079ce 1px solid; BORDER-LEFT: #0079ce 1px solid; BORDER-BOTTOM: #0079ce 1px solid"><p
                         align="center" style="margin-top:3px;">
              <div align="center">=== 申诉定单 ===</div>
            </td>
          </tr>
          <tr> 
            <td><br>
              如果您确认您已对订单进行支付，请在此提交确认。<font color="#FF3333">&nbsp;</font></td>
          </tr>
        </table>
        <br>
        <table border="0" cellspacing="1" cellpadding="0" width="563" bgcolor="#3D5E7D">
          <tr bgcolor="#FFFFFF"> 
            <td> <%if request("ok")="" then%> 
              <table width="500" border="0" cellspacing="0" cellpadding="0" align="center">
<form name="appeal" action="appeal.asp" method="POST">
                <tr> 
                  <td height="45" colspan="3"><font color="#CC3333"><b>下面是您的帐号和未处理的订单号：</b></font></td>
                </tr>
                <tr> 
                  <td height="30" width="100"> 
                    <div align="center">帐号：</div>
                  </td>
                  <td width="100"> 
                      <input maxlength=30 name=username size=20 class="input" value=<%=session("y_it_uid")%> readonly>
                  </td>
                  <td width="300">&nbsp; </td>
                </tr>
                <tr> 
                  <td height="30"> 
                    <div align="center">订单：</div>
                  </td>
                    <td> 
                      <input maxlength=35 name=inBillNo size=20 type=text class="input" value="<%=request("inBillNo")%>" readonly>
                    </td>
                  <td>&nbsp; </td>
                </tr>
                <tr> 
                  <td height="22">&nbsp;</td>
                  <td>&nbsp;</td>
                  <td>&nbsp;</td>
                </tr>
                <tr> 
                  <td height="40" colspan="3"> 
                    <div align="center"> 
                      <input type="submit" name="Submit" value="确 定">&nbsp; 
                      <input type="reset" name="Submit2" value="取 消">
                      <input type="hidden" name="appeal" value="True">
                    </div>
                  </td>
                </tr>
</form>
              </table>
              <%else%> 
              <table width="500" border="0" cellspacing="0" cellpadding="0" align="center">
                <tr> 
                  <td height="45" colspan="3"><font color="#CC3333"><b>您的订单已提交完毕：</b></font></td>
                </tr>
                <tr> 
                  <td height="25" colspan="3"> 
                    <div align="left">您已经对您订单提交过了</div>
                  </td>
                </tr>
                <tr> 
                  <td height="25" colspan="3"> 
                    <div align="left">我们的工作人员将会在第一时间进行处理，请同时向QQ：6664070 提交你的确认，以便得到更快的处理</div>
                  </td>
                </tr>
                <tr> 
                  <td height="25" colspan="3"> 
                    <div align="left">如果您还有任何疑问，请至信 <a href="mailto:<%=Application("y_it_fromemail")%>"><%=Application("y_it_fromemail")%></a>，或QQ：6664070 。
                    </div>
                  </td>
                </tr>
                <tr> 
                  <td height="30">&nbsp;</td>
                  <td colspan="2"> </td>
                </tr>
              </table>
              <%end if%> </td>
          </tr>
        </table>
      </div>
    </td>
  </tr>
</table>
<!--#include file="copyright.asp"-->
</body>
</html>
