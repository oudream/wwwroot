<!--#include file="mymefaq5411jkjkh/favorible/showme.asp"-->
<%on error resume next%>
<%error=request("error")%>
<html>
<head>
<title><%=Application("y_eshop_esname")%>-������ʾ</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="inc/southdns.css" type="text/css">
</head>

<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0">
<!--#include file="top.asp"--> 
<table width="776" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr valign="top"> 
    <td width="146"><img src="image/in54_r23_c2.gif" width="143" height="31"> 
      <table border="0" cellspacing="1" cellpadding="0" align="center" width="100%" bgcolor="#CCCCCC">
        <tr bgcolor="#FFFFFF"> 
          <td> 
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <form action="userlogin.asp" method="post" name="login">
                <tr> 
                  <td colspan="2" height="20"></td>
                </tr>
                <tr> 
                  <td height="22" width="35%"> 
                    <div align="right"><img src="image/in54_r26_c7.gif" width="39" height="13"></div>
                  </td>
                  <td width="65%"> 
                    <div align="center"> 
                      <input maxlength=30 name=username size=10 class="input">
                    </div>
                  </td>
                </tr>
                <tr> 
                  <td height="22"> 
                    <div align="right"><img src="image/in54_r28_c8.gif" width="27" height="12"></div>
                  </td>
                  <td> 
                    <div align="center"> 
                      <input maxlength=16 name=password size=10 type=password class="input">
                    </div>
                  </td>
                </tr>
                <tr> 
                  <td height="30" colspan="2"> 
                    <div align="center"> <a href="reg.asp"><img src="image/in54_r32_c9.gif" width="47" height="20" border="0"></a>&nbsp;&nbsp; 
                      <input type="image" border="0" name="search2" src="image/in54_r32_c13.gif">
                      <input type="hidden" name="userlogin" value="True">
                    </div>
                  </td>
                </tr>
                <tr> 
                  <td height="30" colspan="2"> 
                    <div align="center"><a href="findpassword.asp"><img src="image/in54_r35_c10.gif" width="92" height="18" border="0"></a></div>
                  </td>
                </tr>
                <tr> 
                  <td height="45" colspan="2"> 
                    <div align="center"><img src="image/in54_r40_c10.gif" width="61" height="26"></div>
                  </td>
                </tr>
              </form>
            </table>
          </td>
        </tr>
      </table>
      <img src="image/in54_r43_c2.gif" width="146" height="25"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <%
sqls="select sarea,photo from sarea order by id" 
Set rss=Server.CreateObject("ADODB.RecordSet") 
rss.Open sqls,conn,1,1
n=0
do while not rss.eof
photo=rss("photo")
photo=photo+".gif"
%> 
        <tr> 
          <td height="40"> 
            <div align="center"> <a href="showsarea.asp?sarea=<%=rss("sarea")%>" title="<%=rss("sarea")%>" target="_blank"> 
              <img src="photo/<%=photo%>" width="136" height="34" border="0"></a></div>
          </td>
        </tr>
        <%
rss.movenext
n=n+1
if n>=7 then exit do
loop
rss.close
set rss=nothing
%> 
        <tr> 
          <td height="25"> 
            <div align="right"><img src="image/in54_r77_c14.gif" width="40" height="13"></div>
          </td>
        </tr>
      </table>
      <br>
    </td>
    <td width="630"> <br>
      <table border="0" cellspacing="1" cellpadding="0" align="center" width="600" bgcolor="#5C9DBE">
              <form action="userlogin.asp" method="post" name="login">
        <tr bgcolor="#FFFFFF"> 
          <td>
            <table width="580" border="0" cellspacing="0" cellpadding="0" align="center">
              <tr><td height="30" colspan="3">&nbsp;</td></tr>
              <tr> 
                <td width="20" height="22">&nbsp;</td>
                  <td bgcolor="#5C9DBE">&nbsp;<b>�Բ������Ĳ�������ȷ������������ʾ</b></td>
                <td width="160">&nbsp;</td>
              </tr>
              <tr><td height="30" colspan="3">&nbsp;</td></tr>
            </table>

            <table width="580" border="0" cellspacing="0" cellpadding="0" align="center">
              <tr>
                <td width="150" valign="top"><img src="image/54-zc1.gif" width="150" height="182"></td>
                <td valign="top" width="350">
      <table border="0" cellspacing="1" cellpadding="0" align="center" width="330" bgcolor="#5C9DBE">
                    <tr bgcolor="#FFFFFF"> 
                        <td> <br>
                          <table width="300" border="0" cellspacing="0" cellpadding="0" align="center">
                            <tr>
                              <td><font color="#CC3333"><%if error="001" then%> 
                                <br>
                                �Բ��������Ǵӱ����������ύ�ı���Ϣ������<br>
                                <br>
                                ������������������򱾲������Է����������ƻ���<br>
                                <br>
                                ���������һ�����ߵ�Ȩ������<br>
                                <%end if%> <%if error="002" then%> <br>
                                �Բ�������ѡ����û����Ѿ���ע����.<br>
                                <br>
                                �����Ե��<a href="reg.asp">����ע��</a>�����ʺ�.<br>
                                <br>
                                лл��ѡ�����ǵķ���<br>
                                <%end if%> <%if error="003" then%> <br>
                                �Բ��𣬲�û��������д���û��ʺ�.<br>
                                <br>
                                �����Ե��<a href="index.asp">���µ�½</a>�����ʺ�.<br>
                                <br>
                                ���ߵ��<a href="reg.asp">����ע��</a>�����ʺ�.<br>
                                <br>
                                лл��ѡ�����ǵķ���<br>
                                <%end if%> <%if error="004" then%> <br>
                                �Բ������������д���ʺŵ�����.<br>
                                <br>
                                �����Ե��<a href="index.asp">���µ�½</a>�����ʺ�.<br>
                                <br>
                                ���ߵ��<a href="findpassword.asp">Ѹ���һ�</a>�����ʺ�.<br>
                                <br>
                                лл��ѡ�����ǵķ���<br>
                                <%end if%> <%if error="005" then%> <br>
                                �Բ��������ʺ��Ѿ���½����.<br>
                                <br>
                                ������Ҫ�ظ���½.���<a href="index.asp">������ҳ</a><br>
                                <br>
                                лл��ѡ�����ǵķ���<br>
                                <%end if%> <%if error="006" then%> 
                                <p><br>
                                  �װ����û�����û��ע�����û�е�½��<br>
                                  <br>
                                  ���½����ע�ᡣ<br>
                                  <br>
                                  <a href="findpassword.asp">�������룿</a> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="reg.asp">��Ҫע�� 
                                  </a> <br>
                                  �����ʺţ� 
                                  <input type="text" name="username" size="12" maxlength="15" class="form">
                                  &nbsp;&nbsp; 
                                  <input type="submit" name="Submit" value="��½">
                                  <input type="hidden" name="userlogin" value="True">
                                  <br>
                                  �������룺 
                                  <input type="password" name="password" size="12" maxlength="15" class="form">
                                  &nbsp;&nbsp; 
                                  <input type="submit" name="Submit2" value="ȡ��">
                                  <br>
                                  <%end if%> <%if error="007" then%> <br>
                                  �Բ������Ĺ��ﳵΪ��.<br>
                                  <br>
                                  ���<a href="index.asp">���¹���.</a><br>
                                  <br>
                                  лл��ѡ�����ǵķ���<br>
                                  <%end if%> <%if error="008" then%> <br>
                                  �Բ��𣬲�û������ʺ�.<br>
                                  <br>
                                  ����������������.<br>
                                  <br>
                                  ���<a href="findpassword.asp">���²�ѯ.</a><br>
                                  <br>
                                  ����������κ����ʣ�������: <a href="mailto:<%=Application("y_eshop_fromemail")%>"><%=Application("y_eshop_fromemail")%></a> 
                                  <%end if%> <%if error="009" then%> <br>
                                  �Բ������Ķ����Ѿ����߹���.<br>
                                  <br>
                                  ���ǻ������ʱ���ڴ������Ķ���.<br>
                                  <br>
                                  ���<a href="index.asp">������ҳ.</a><br>
                                  <br>
                                  ����������κ����ʣ�������: <a href="mailto:<%=Application("y_eshop_fromemail")%>"><%=Application("y_eshop_fromemail")%></a> 
                                  <%end if%> <%if error="010" then%> <br>
                                  �Բ��𣬴˶�����������.<br>
                                  <br>
                                  ������ʵ�����ʺźͶ����ź��ٽ�������.<br>
                                  <br>
                                  ���<a href="usererror.asp">���½�������.</a><br>
                                  <br>
                                  ����������κ����ʣ�������: <a href="mailto:<%=Application("y_eshop_fromemail")%>"><%=Application("y_eshop_fromemail")%></a> 
                                  <%end if%> <%if error="011" then%> <br>
                                  �Բ�������û���Թ���Ա��½.<br>
                                  <br>
                                  ���<a href="admin/adminlogin.asp">���µ�½.</a><br>
                                  <br>
                                  ����������κ����ʣ�������: <a href="mailto:<%=Application("y_eshop_fromemail")%>"><%=Application("y_eshop_fromemail")%></a> 
                                  <%end if%> <%if error="013" then%> <br>
                                  �Բ�����һ�εĹ�����õ���50RMB.<br>
                                  <br>
                                  ���<a href="index.asp">��������.</a><br>
                                  <br>
                                  ����������κ����ʣ�������: <a href="mailto:<%=Application("y_eshop_fromemail")%>"><%=Application("y_eshop_fromemail")%></a> 
                                  <%end if%> <%if error="016" then%> <br>
                                  �Բ�����û��ѡ����Ҫ���۵���Ʒ.<br>
                                  <br>
                                  ���<a href="index.asp">������ҳ.</a><br>
                                  <br>
                                  ����������κ����ʣ�������: <a href="mailto:<%=Application("y_eshop_fromemail")%>"><%=Application("y_eshop_fromemail")%></a> 
                                  <%end if%> 
                                </b></font></td>
                            </tr>
                          </table>
                          <br>
                        </td>
                    </tr>
                  </table>
                </td>
                <td width="80" valign="top">&nbsp;</td>
              </tr>
              <tr>
                  <td height="120">&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
              </tr>
            </table>
          </td>
        </tr></form>
      </table>
    </td>
  </tr>
</table>
<!--#include file="copyright.asp"--> 
</body>
</html>
