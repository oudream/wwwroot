<%@ Language=VBScript%>
<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<!--#include file="function.asp"-->
<!--#include file="char.inc"-->
<%
dim linktype
linktype=checkstr(request("linktype"))
if linktype="" then
Response.Write "δָ������"
else
if not IsNumeric(linktype) then
 response.write "<script>alert('�Ƿ�����');history.back()</script>"
 response.end
else



PageShowSize = 10            'ÿҳ��ʾ���ٸ�ҳ
MyPageSize   = 20           'ÿҳ��ʾ����������
	
If Not IsNumeric(Request("page")) Or IsEmpty(Request("page")) Or Request("page") <=0 Then
MyPage=1
Else
MyPage=Int(Abs(Request("page")))
End if
end if


%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��������__<%=jjgn%></title>
<LINK href=news.css rel=stylesheet>
</head>

<body topmargin="0">
<!--#include file=top.asp -->
<table width="760" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <tr valign="top"> 
    <td width="180" background="IMAGES/menu-d.gif"> 
      <table border="0" cellpadding="0" valign=top cellspacing="0" style="border-collapse: collapse" width="100%" id="AutoNumber2" >
        <tr> 
          <td height="3" valign="middle" bgcolor="#FFFFFF"><img src="IMAGES/kb.gif" width="9" height="3"></td>
        </tr>
        <tr> 
          <td width="100%" height="25" valign="middle" background="IMAGES/menu-d-user.gif"> 
            <p align="center"><font color="#000000">��������</font></td>
        </tr>
        <tr> 
          <td width="100%" height="24" align="center" valign="middle" background="IMAGES/menu-d.gif"> 
            <a class=class href="morelink.asp?linktype=2"><br>
            ��ҳͼƬ����</a><br> <a class=class href="morelink.asp?linktype=1"><br>
            ����ͼƬ����</a><br> <br> <a class=class href="morelink.asp?linktype=0">��������</a><br> 
            <br> </td>
        </tr>
        
        
      </table>
    </td>
    <td width="6">&nbsp;</td>
    <td width="574" background="IMAGES/menu-l.gif"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td height="3" bgcolor="#FFFFFF"><img src="IMAGES/kb.gif" width="9" height="3"></td>
        </tr>
        <tr> 
          <td height="25" background="IMAGES/menu-l-m.gif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font class=m_tittle>&nbsp;</font>��Ŀ����<b>&nbsp;&nbsp; 
            ��ǰλ�ã�<a class="daohang" href="./" >��վ��ҳ</a>&gt;&gt;��������<font color=red></font></b></td>
        </tr>
        <tr> 
          <td height="25" background="IMAGES/menu-l.gif"><div align="center"><b> 
              <font color=red><a class=daohang href="./" > 
              <script language=javascript src=./zongg/ad.asp?i=13></script> 
              </a> </font></b></div></td>
        </tr>
        <tr> 
          <td height="25" background="IMAGES/menu-l-zj.gif" class="daohang">&nbsp;</td>
        </tr>
        <tr> 
          <td height="108" background="IMAGES/menu-l.gif"> 
            <div align="center"> 
              <table width="95%"  border="0" align="center" cellpadding="0" cellspacing="0" id="AutoNumber3" style="border-collapse: collapse">
                <tr> 
                  <td valign=top>&nbsp;</td>
                </tr>
                <tr> 
                  <td valign=top><table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse" height="51">
                      <%

set rs3=server.CreateObject("ADODB.RecordSet")

rs3.Source="select * from "& db_Link_Table &" where (linktype="&linktype&" and pass=1) order by id DESC"
rs3.Open rs3.Source,conn,3,1

If Not rs3.eof then
Rs3.PageSize     = MyPageSize
MaxPages         = Rs3.PageCount
Rs3.absolutepage = MyPage
total            = Rs3.RecordCount
%>
                      <tr> 
                        <td width=100% height="20" valign="middle" > <p align="center">&nbsp;�� 
                            <%=total%> ������ǰ�� <%=Mypage%>/<%=Maxpages%> ҳ��ÿҳ <%=MyPageSize%> 
                            ��</td>
                      </tr>
                      <tr> 
                        <td width="100%" height="27"> 
                          <%
i = 0
do until rs3.Eof or i = rs3.PageSize
%>
                          <table width="98%" border="0" cellspacing="0" cellpadding="0">
                            <tr> 
                              <td width="5%" height="31" valign="middle">&nbsp;<img src="images/006.gif" width="8" height="10" border="0">&nbsp; 
                              </td>
                              <%
                              if  rs3("linktype")=0  then
							  %>
							  <td width="18%" valign="middle">��������</td>
							  <%
							  else
							  %>
<%
                              if  rs3("linktype")=2  then
							  %>
							  <td width="18%" valign="middle"><a class=middle href="<%=rs3("weburl")%>" target=_blank title="<%=trim(rs3("content"))%>"><img src="<%=rs3("logo")%>" width="88" height="31"border="0"></a></td>
							  <%
							  else
							  %>

							  <td width="18%" valign="middle"><a class=middle href="<%=rs3("weburl")%>" target=_blank title="<%=trim(rs3("content"))%>"><img src="<%=rs3("logo")%>" width="88" height="31"border="0"></a></td>
<%end if%><%end if%>

                              <td width="40%" valign="middle"><font class=middle>......<a class=middle href="<%=rs3("weburl")%>" target=_blank title="<%=trim(rs3("content"))%>"><%=trim(rs3("webname"))%></font></a></td>
                              <td width="27%" valign="middle"><font class=middle>(<%=trim(rs3("dateandTime"))%>)</font></td>
                            </tr>
                          </table>
                          <br> 
                          <%
rs3.MoveNext
i = i + 1
loop
%>
                        </td>
                      </tr>
                      <tr> 
                        <td width="100%" align=center height="20" valign="middle">�� 
                          <%=total%> ������ǰ�� <%=Mypage%>/<%=Maxpages%> ҳ��ÿҳ <%=MyPageSize%> 
                          �� 
                          <%
url="morelink.asp?linktype="&linktype&""
PageNextSize=int((MyPage-1)/PageShowSize)+1
Pagetpage=int((total-1)/rs3.PageSize)+1

if PageNextSize >1 then
PagePrev=PageShowSize*(PageNextSize-1)
Response.write "<a class=black href='" & Url & "&page=" & PagePrev & "' title='��" & PageShowSize & "ҳ'>��һ��ҳ</a> "
Response.write "<a class=black href='" & Url & "&page=1' title='��1ҳ'>ҳ��</a> "
end if
if MyPage-1 > 0 then
Prev_Page = MyPage - 1
Response.write "<a class=black href='" & Url & "&page=" & Prev_Page & "' title='��" & Prev_Page & "ҳ'>��һҳ</a> "
end if

if Maxpages>=PageNextSize*PageShowSize then
PageSizeShow = PageShowSize
Else
PageSizeShow = Maxpages-PageShowSize*(PageNextSize-1)
End if
If PageSizeShow < 1 Then PageSizeShow = 1
for PageCounterSize=1 to PageSizeShow
PageLink = (PageCounterSize+PageNextSize*PageShowSize)-PageShowSize
if PageLink <> MyPage Then
Response.write "<a class=black href='" & Url & "&page=" & PageLink & "'>[" & PageLink & "]</a> "
else
Response.Write "<B>["& PageLink &"]</B> "
end if
If PageLink = MaxPages Then Exit for
Next

if Mypage+1 <=Pagetpage  then
Next_Page = MyPage + 1
Response.write "<a class=black href='" & Url & "&page=" & Next_Page & "' title='��" & Next_Page & "ҳ'>��һҳ</A>"
end if

if MaxPages > PageShowSize*PageNextSize then
PageNext = PageShowSize * PageNextSize + 1
Response.write " <A class=black href='" & Url & "&page=" & Pagetpage & "' title='��"& Pagetpage &"ҳ'>ҳβ</A>"
Response.write " <a class=black href='" & Url & "&page=" & PageNext & "' title='��" & PageShowSize & "ҳ'>��һ��ҳ</a>"
End if
else
Response.write "<tr><td align=center>&nbsp;����������Ϣ</td></tr>"
				
End If
Rs3.close
set rs3=nothing

%>
                        </td>
                      </tr>
                    </table></td>
                </tr>
                <tr> 
                  <td valign=top>&nbsp;</td>
                </tr>
              </table>
            </div></td>
        </tr>
      </table>
    </td>
  </tr>
  <tr valign="top"> 
    <td height="20" background="IMAGES/menu-d-t.gif" bgcolor="#FFFFFF">&nbsp;</td>
    <td height="20" bgcolor="#FFFFFF">&nbsp;</td>
    <td height="20" background="IMAGES/menu-l-t.gif" bgcolor="#FFFFFF" class="daohang">&nbsp;</td>
  </tr>
</table>
<!--#include file=bottom.asp -->
</body>
</html><%end if%>