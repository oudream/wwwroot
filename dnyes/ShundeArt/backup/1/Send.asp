<%@ Language=VBScript%>

<!--#include file=conn.asp -->
<!--#include file=config.asp -->
<%
NewsID=Request.QueryString("NewsID")

if NewsID="" then
Response.Write "<script>alert('δָ������');history.back()</script>"
response.end
else
if not IsNumeric(NewsID) then
 response.write "<script>alert('�Ƿ�����');history.back()</script>"
 response.end
else

reffer=Request.ServerVariables("HTTP_REFERER")

end if 
end if
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��������_<%=jjgn%></title>
<LINK href=news.css rel=stylesheet>
</head>

<body topmargin="0"><form method="POST" action="Send_submit.asp"><input type=hidden name=NewsID value=<%=NewsID%>><input type=hidden name=reffer value=<%=reffer%>>
  <!--#include file=top.asp -->
  <table width="760" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
    <tr valign="top"> 
      <td height="356"> 
        <table width="102%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="3" ><img src="IMAGES/kb.gif" width="9" height="3"></td>
          </tr>
          <tr> 
            <td width="100%" height="25" background="IMAGES/menu-guestbook.gif" bgcolor="CDCCCC" >&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��Ŀ����<a class=daohang href="./" >&nbsp;</a><strong><b>&nbsp;&nbsp;��ǰλ�ã�<a class="daohang" href="./" >��վ��ҳ</a>&gt;&gt;�����ʼ�</b></strong></td>
          </tr>
          <tr bgcolor="CDCCCC"> 
            <td background="IMAGES/menu-guest-l.gif" bgcolor="#FFFFFF"><div align="center"><a class=daohang href="./" ><br>
                </a>
                <script language=javascript src=./zongg/ad.asp?i=14></script>                
                </div></td>
          </tr>
          <tr bgcolor="CDCCCC"> 
            <td height="25" align="center" background="IMAGES/menu-l-guest.gif" bgcolor="#FFFFFF"><a href="guestadd.asp" class=class>��Ҫ����</a> 
              | <a class=class href="javascript:window.location.reload()">ˢ��</a> 
            </td>
          </tr>
          <tr bgcolor="CDCCCC"> 
            <td height="25" align="center" background="IMAGES/menu-guest-l.gif" bgcolor="#FFFFFF"><table width="95%"  border="0" align="center" cellpadding="0" cellspacing="0" id="AutoNumber3" style="border-collapse: collapse">
                <tr> 
                  <td valign=top>&nbsp;</td>
                </tr>
                <tr> 
                  <td valign=top><table border="0" cellpadding="3" cellspacing="0" width="100%">
                      <tr> 
                        <td width="100%" height="50"><font color="000099">�����������ѵ����֣� 
                          <input type="text" name="name" size="50">
                          </font></td>
                      </tr>
                      <tr> 
                        <td width="100%" height="50"><font color="000099">�����������ѵ�MAIL��</font><font color="000099"> 
                          <input type="text" name="email" size="50">
                          </font></td>
                      </tr>
                      <tr> 
                        <td width="100%" height="50"><font color="000099">���������Լ������֣� 
                          <input type="text" name="myname" size="50">
                          </font></td>
                      </tr>
                      <tr> 
                        <td width="100%" align="center" height="50"> <input type="submit" value="  ����  " name="B1"></td>
                      </tr>
                    </table></td>
                </tr>
                <tr> 
                  <td valign=top>&nbsp;</td>
                </tr>
              </table></td>
          </tr>
          <tr bgcolor="CDCCCC"> 
            <td height="19" align="center" background="IMAGES/menu-guest-t.gif" bgcolor="#FFFFFF">&nbsp;</td>
          </tr>
        </table>        
      </td>
    </tr>
  </table>



<!--#include file=bottom.asp -->
</form>
</body>

</html>