<!--#include file="mymefaq5411jkjkh/favorible/showme.asp" -->
<script language="JavaScript">
var strVariable = "This is a string";
strVariable = strVariable.fontcolor("red");
</script>
<html>
<head>
<title>����-���ſƼ� ���������ռ��������֪ʶ</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="keywords" content="��������֪ʶ,��վ�ռ�֪ʶ,����֪ʶ,����֪ʶ">
<meta name="description" content="��������֪ʶ,��վ�ռ�֪ʶ,����֪ʶ,����֪ʶ">
<link rel="stylesheet" href="css.css" type="text/css">
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0">
<!--#include file="top.asp"-->
<table width="776" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="10" background="images/dnyes-bg-left-and-right.gif"><img src="images/dnyes-bg-left-and-right.gif" width="10" height="1"></td>
    <td><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr> 
          <td colspan="5"><table width="100%" border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td width="1" background="images/2x2.gif"><img src="images/1x2-black.gif" width="1" height="2"></td>
                <td><img src="images/1-x.gif" width="754" height="56"></td>
                <td width="1" background="images/1x2-black.gif"><img src="images/1x2-black.gif" width="1" height="2"></td>
              </tr>
            </table></td>
        </tr>
        <tr> 
          <td width="229" valign="top" background="images/left-228x5.gif"> <table width="229" border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td width="229" height="53" background="images/left-1b.gif"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td width="10%" height="18" valign="bottom">&nbsp; </td>
                      <td width="8%" valign="bottom"><img src="images/gogo.gif" width="6" height="15"></td>
                      <td width="82%" valign="bottom"><font color="#FFFFFF"><b>�� 
                        �� �� ¼</b></font></td>
                    </tr>
                  </table></td>
              </tr>
              <tr> 
                <td valign="top"> <table width="228" border="0" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td background="images/left-1-left-2b.gif"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                          <tr> 
                            <td width="21" valign="top"><img src="images/left-1-left-1b.gif" width="21" height="190"></td>
                            <td valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                                <tr> 
                                  <td> 
                                    <!--#include file="userlogin.asp"-->
                                  </td>
                                </tr>
                                <tr> 
                                  <td> 
                                    <!--#include file="userhelp.asp" -->
                                  </td>
                                </tr>
                              </table></td>
                          </tr>
                        </table></td>
                    </tr>
                    <tr> 
                      <td><img src="images/left-1-left-3b.gif" width="229" height="12"></td>
                    </tr>
                  </table></td>
              </tr>
            </table> 
            <table width="229" border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td width="229" height="28" background="images/left-2.gif"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td width="10%" height="18" valign="bottom">&nbsp; </td>
                      <td width="8%" valign="bottom"><img src="images/gogo.gif" width="6" height="15"></td>
                      <td width="82%" valign="bottom"><font color="#FFFFFF"><b>�� 
                        �� �� ��</b></font></td>
                    </tr>
                  </table></td>
              </tr>
            </table>
            <TABLE width=70% border=0 cellPadding=0 cellSpacing=0>
              <TBODY>
                <TR> 
                  <TD height="3" 
                vAlign=top 
                style="BORDER-left: #000000 1px solid"><img src="1.gif" width="1" height="1"></TD>
                </TR>
              </TBODY>
            </TABLE>
            <table width="228" height="100%" border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td width="1" background="images/1x2-black.gif"><img src="images/1x2-black.gif" width="1" height="2"></td>
                <td valign="top"> 
                  <!--#include file="support.asp" -->
                </td>
              </tr>
            </table></td>
          <td width="7">&nbsp;</td>
          <td width="513" valign="top"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td colspan="2"><div align="center"> 
                    <!--#include file="basetable.asp" -->
                  </div></td>
              </tr>
              <tr> 
                <td width="100%"><br> <table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                      <td colspan="2" valign="top">+ ����֪ʶ +</td>
                    </tr>
                    <tr> 
                      <td colspan="2" valign="top"><img src="images/other-1.gif" width="513" height="12"></td>
                    </tr>
                    <%
sqlnews="select * from myfaq where newstype='����֪ʶ' order by id desc" 
Set rsnews=Server.CreateObject("ADODB.RecordSet") 
rsnews.Open sqlnews,conn,1,1 
n=0
do while not rsnews.eof
%>
                    <tr> 
                      <td width="25" height="18" align="center"><img src="images/dot-6.gif" width="8" height="6" border="0"></td>
                      <td width="488" valign="top"><a href="showbase.asp?id=<%=rsnews("id")%>" title="<%=rsnews("subject")%>" target="_blank"> 
                        <font color=<%=rsnews("color")%>> 
                        <% if len(rsnews("subject"))>100 then%>
                        <%=left(rsnews("subject"),98)%>.. 
                        <%else%>
                        <%=rsnews("subject")%> 
                        <%end if%>
                        </font></a></td>
                    </tr>
                    <%
n=n+1
rsnews.movenext
if n>=10 then exit do
loop
rsnews.close
set rsnews=nothing
%>
                    <tr> 
                      <td>&nbsp;</td>
                      <td><div align="right"><a href="viewbase.asp?newstype=����֪ʶ" >����&gt;&gt;</a></div></td>
                    </tr>
                  </table></td>
              </tr>
              <tr> 
                <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr valign="top"> 
                      <td colspan="2">+ ����֪ʶ +</td>
                    </tr>
                    <tr valign="top"> 
                      <td colspan="2"><img src="images/other-1.gif" width="513" height="12"></td>
                    </tr>
                    <%
sqlnews="select * from myfaq where newstype='��������֪ʶ' order by id desc" 
Set rsnews=Server.CreateObject("ADODB.RecordSet") 
rsnews.Open sqlnews,conn,1,1 
n=0
do while not rsnews.eof
%>
                    <tr> 
                      <td width="25" height="18" align="center"><img src="images/dot-6.gif" width="8" height="6" border="0"></td>
                      <td width="488" valign="top"><a href="showbase.asp?id=<%=rsnews("id")%>" title="<%=rsnews("subject")%>" target="_blank"> 
                        <font color=<%=rsnews("color")%>> 
                        <% if len(rsnews("subject"))>100 then%>
                        <%=left(rsnews("subject"),98)%>.. 
                        <%else%>
                        <%=rsnews("subject")%> 
                        <%end if%>
                        </font></a></td>
                    </tr>
                    <%
n=n+1
rsnews.movenext
if n>=10 then exit do
loop
rsnews.close
set rsnews=nothing
%>
                    <tr> 
                      <td>&nbsp;</td>
                      <td><div align="right"><a href="viewbase.asp?newstype=��������֪ʶ" >����&gt;&gt;</a></div></td>
                    </tr>
                  </table></td>
              </tr>
              <tr> 
                <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr valign="top"> 
                      <td colspan="2">+ ��վ֪ʶ +</td>
                    </tr>
                    <tr valign="top"> 
                      <td colspan="2"><img src="images/other-1.gif" width="513" height="12"></td>
                    </tr>
                    <%
sqlnews="select * from myfaq where newstype='��վ����֪ʶ' order by id desc" 
Set rsnews=Server.CreateObject("ADODB.RecordSet") 
rsnews.Open sqlnews,conn,1,1 
n=0
do while not rsnews.eof
%>
                    <tr> 
                      <td width="25" height="18" align="center"><img src="images/dot-6.gif" width="8" height="6" border="0"></td>
                      <td width="488" valign="top"><a href="showbase.asp?id=<%=rsnews("id")%>" title="<%=rsnews("subject")%>" target="_blank"> 
                        <font color=<%=rsnews("color")%>> 
                        <% if len(rsnews("subject"))>100 then%>
                        <%=left(rsnews("subject"),98)%>.. 
                        <%else%>
                        <%=rsnews("subject")%> 
                        <%end if%>
                        </font></a></td>
                    </tr>
                    <%
n=n+1
rsnews.movenext
if n>=10 then exit do
loop
rsnews.close
set rsnews=nothing
%>
                    <tr> 
                      <td>&nbsp;</td>
                      <td><div align="right"><a href="viewbase.asp?newstype=��վ����֪ʶ" >����&gt;&gt;</a></div></td>
                    </tr>
                  </table></td>
              </tr>
              <tr> 
                <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr valign="top"> 
                      <td colspan="2">+ ��վ�ƹ� +</td>
                    </tr>
                    <tr valign="top"> 
                      <td colspan="2"><img src="images/other-1.gif" width="513" height="12"></td>
                    </tr>
                    <%
sqlnews="select * from myfaq where newstype='��վ�ƹ�֪ʶ' order by id desc" 
Set rsnews=Server.CreateObject("ADODB.RecordSet") 
rsnews.Open sqlnews,conn,1,1 
n=0
do while not rsnews.eof
%>
                    <tr> 
                      <td width="25" height="18" align="center"><img src="images/dot-6.gif" width="8" height="6" border="0"></td>
                      <td width="488" valign="top"><a href="showbase.asp?id=<%=rsnews("id")%>" title="<%=rsnews("subject")%>" target="_blank"> 
                        <font color=<%=rsnews("color")%>> 
                        <% if len(rsnews("subject"))>100 then%>
                        <%=left(rsnews("subject"),98)%>.. 
                        <%else%>
                        <%=rsnews("subject")%> 
                        <%end if%>
                        </font></a></td>
                    </tr>
                    <%
n=n+1
rsnews.movenext
if n>=10 then exit do
loop
rsnews.close
set rsnews=nothing
%>
                    <tr> 
                      <td>&nbsp;</td>
                      <td><div align="right"><a href="viewbase.asp?newstype=�ײ�֪ʶ" >����&gt;&gt;</a></div></td>
                    </tr>
                  </table></td>
              </tr>
              <tr> 
                <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr valign="top"> 
                      <td colspan="2">+ ����֪ʶ +</td>
                    </tr>
                    <tr valign="top"> 
                      <td colspan="2"><img src="images/other-1.gif" width="513" height="12"></td>
                    </tr>
                    <%
sqlnews="select * from myfaq where newstype='��ҵ����֪ʶ' order by id desc" 
Set rsnews=Server.CreateObject("ADODB.RecordSet") 
rsnews.Open sqlnews,conn,1,1 
n=0
do while not rsnews.eof
%>
                    <tr> 
                      <td width="25" height="18" align="center"><img src="images/dot-6.gif" width="8" height="6" border="0"></td>
                      <td width="488" valign="top"><a href="showbase.asp?id=<%=rsnews("id")%>" title="<%=rsnews("subject")%>" target="_blank"> 
                        <font color=<%=rsnews("color")%>> 
                        <% if len(rsnews("subject"))>100 then%>
                        <%=left(rsnews("subject"),98)%>.. 
                        <%else%>
                        <%=rsnews("subject")%> 
                        <%end if%>
                        </font></a></td>
                    </tr>
                    <%
n=n+1
rsnews.movenext
if n>=10 then exit do
loop
rsnews.close
set rsnews=nothing
%>
                    <tr> 
                      <td>&nbsp;</td>
                      <td><div align="right"><a href="viewbase.asp?newstype=��ҵ����֪ʶ" >����&gt;&gt;</a></div></td>
                    </tr>
                  </table></td>
              </tr>
              <tr> 
                <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr valign="top"> 
                      <td colspan="2">+ �����й� +</td>
                    </tr>
                    <tr valign="top"> 
                      <td colspan="2"><img src="images/other-1.gif" width="513" height="12"></td>
                    </tr>
                    <%
sqlnews="select * from myfaq where newstype='�����й�֪ʶ' order by id desc" 
Set rsnews=Server.CreateObject("ADODB.RecordSet") 
rsnews.Open sqlnews,conn,1,1 
n=0
do while not rsnews.eof
%>
                    <tr> 
                      <td width="25" height="18" align="center"><img src="images/dot-6.gif" width="8" height="6" border="0"></td>
                      <td width="488" valign="top"><a href="showbase.asp?id=<%=rsnews("id")%>" title="<%=rsnews("subject")%>" target="_blank"> 
                        <font color=<%=rsnews("color")%>> 
                        <% if len(rsnews("subject"))>100 then%>
                        <%=left(rsnews("subject"),98)%>.. 
                        <%else%>
                        <%=rsnews("subject")%> 
                        <%end if%>
                        </font></a></td>
                    </tr>
                    <%
n=n+1
rsnews.movenext
if n>=10 then exit do
loop
rsnews.close
set rsnews=nothing
%>
                    <tr> 
                      <td>&nbsp;</td>
                      <td><div align="right"><a href="viewbase.asp?newstype=�����й�֪ʶ" >����&gt;&gt;</a></div></td>
                    </tr>
                  </table></td>
              </tr>
              <tr> 
                <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr valign="top"> 
                      <td colspan="2">&nbsp;</td>
                    </tr>
                    <tr> 
                      <td width="29">&nbsp;</td>
                      <td width="547">&nbsp;</td>
                    </tr>
                    <tr> 
                      <td>&nbsp;</td>
                      <td><div align="right"></div></td>
                    </tr>
                  </table></td>
              </tr>
              <tr> 
                <td>&nbsp;</td>
              </tr>
            </table></td>
          <td width="7" background="images/right-7x5.gif">&nbsp;</td>
        </tr>
      </table></td>
    <td width="10" background="images/dnyes-bg-left-and-right.gif"><img src="images/dnyes-bg-left-and-right.gif" width="10" height="1"></td>
  </tr>
</table>
<!--#include file="copyright.asp"-->
</body>
</html>