<%@ LANGUAGE="VBSCRIPT" %>
<!--#include file=config.asp -->
<script language="javascript">
<!--
function isok(theform)
{
    if (theform.name.value=="")
    {
        alert("����д������ƣ�");
        theform.name.focus();
        return (false);
    }
    if (theform.url.value=="")
    {
        alert("����д����URL��");
        theform.url.focus();
        return (false);
    }
    return (true);
}
-->
</script>
<%
if request.querystring("job")="add" then


if  isnumeric(request.querystring("id"))=false then
%>
<script language="VBScript" type="text/VBScript">
msgbox "�����������"
 history.back(1)
              </script>
<%
response.end
end if

Call  addrk()
end if


%>
<html><head>
<meta http-equiv=Content-Language content=zh-cn>
<meta http-equiv=Content-Type content=text/html; charset=gb2312>
<meta http-equiv=Pragma content=no-cache>
<link href=style.css rel=STYLESHEET type=text/css>
<title>����.������ϵͳ</title>
</script>
<base target=_top>
</head><body marginwidth=0 marginheight=0>


<%
call  Iflogin()
%>

<%

adsconn.open adsdata
dim getid,adsrs,adssql
getid=cint(request.querystring("id"))

set adsrs=server.createobject("adodb.recordset")
adssql="select * from ads where id="&getid
adsrs.open adssql,adsconn,1,1

%>

<div align=center><center>
<table border=0 cellspacing=0 cellpadding=0  width="100%" ><tr><td bgcolor=#304050>

<table border=0 width="100%" cellpadding=2 style="border-collapse: collapse" bordercolor="#111111">
<form method=POST action=edit.asp?job=add&id=<%=adsrs("id")%>>

<tr><td bgcolor=#ffffff colspan=2 align=center height="20"><br><b><font color=red>�޸Ĺ��IDΪ <%=adsrs("id")%> �Ĺ������Ϣ</font></b><hr color="#808080" size="1"></td></tr>

<tr><td bgcolor=#ffffff colspan=2>&nbsp;�������&nbsp;&nbsp;&nbsp;&nbsp; <input type=text name=name size=30 maxlength=30 value=<%=adsrs("sitename")%>>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
  <font color="#FF0000">
  <%if request.querystring("job")="suc" then%>�������Ϣ�޸ĳɹ���<%end if%></font></td></tr>
<tr><td  bgcolor=#ffffff colspan=2>&nbsp;����URL&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <input type=text name=url size=40 value="<%=adsrs("url")%>"></td></tr>
<tr><td  bgcolor=#ffffff>&nbsp;���/����&nbsp;&nbsp;&nbsp; 
                  <textarea rows="5" name="intro" cols="38"><%=adsrs("intro")%></textarea></td>
  <td  bgcolor=#ffffff><font color="#FF0000">&nbsp;��ʾ��</font><br>
                  &nbsp;<font color="#808080">�����Ƕ������뽫������������˴� ����URL��Ч <br>
                  &nbsp;�����ʾ���ı�������ʾΪ�������<br>
                  &nbsp;ֻ��GIFͼƬ��Flash����ʱURL��д��Ч</font></td></tr>
<tr><td bgcolor=#ffffff colspan=2>&nbsp;��ʾ����&nbsp;&nbsp;&nbsp;&nbsp;<input type="radio" name="xslei" value="gif"  <%if adsrs("xslei")="gif" then%>checked<%end if%>>GIFͼƬ 
                  <input type="radio" name="xslei" value="swf" <%if adsrs("xslei")="swf" then%>checked<%end if%>><font siz=3 >Flash���� </font>
                  <input type="radio" name="xslei" value="txt" <%if adsrs("xslei")="txt" then%>checked<%end if%>><font siz=3 >���ı� </font>        
                  <input type="radio" name="xslei" value="dai" <%if adsrs("xslei")="dai" then%>checked<%end if%>>Ƕ�����<font siz=3 >&nbsp; </font>    </td></tr>

<tr><td bgcolor=#ffffff colspan=2>&nbsp;ͼƬURL&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <input type=text name=gif_url size=40 value="<%=adsrs("gif_url")%>">    <font color=red siz=3 >&nbsp;</font><font siz=3 >�߶� </font> 
                  <input type=text name=hei size=3 value="<%=adsrs("hei")%>" maxlength="4"><font siz=3 >&nbsp; 
                  ��� </font> 
                  <input type=text name=wid size=3 value="<%=adsrs("wid")%>" maxlength="4"><font color=red siz=3 > </font></td></tr>

<tr><td bgcolor=#ffffff colspan=2>��</td></tr>

<tr><td bgcolor=#ffffff colspan=2>&nbsp;�������λ&nbsp;&nbsp;
<%'adsconn.Open adsdata
                Call  Ggwxlxx(adsrs("place")) %>
                  
                  
</td></tr>

<tr><td bgcolor=#ffffff colspan=2>&nbsp;���Ӵ򿪷�ʽ <select size=1 name=window><option value=0 <%if adsrs("window")=0 then%>selected<%end if%>>�´��ڴ�</option><option value=1 <%if adsrs("window")=1 then%>selected<%end if%>>ԭ���ڴ�</option></select></td></tr>

<tr><td bgcolor=#DBDBDB colspan=2 height="20">&nbsp;���������飺��ѡ����һ�飬�������Ƹù���Զ���������״̬�������������޸�����������
  ���¼����档</td></tr>


<tr>
<td bgcolor=#ffffff colspan=2>&nbsp;<font color="#808080">** ����/ʱ���ʽ�ϸ�Ϊ:</font><font color=#FF0000>yyyy-mm-dd hh:mm:ss �� yyyy-mm-dd</font>
</td></tr>

<tr><td bgcolor=#ffffff><%
if adsrs("class")=0 then
%>
<table border=0><tr><td><font color="#FF0000">��1��</font><input type=radio value=0 name=class checked></td><td>������ѭ��</td></tr></table>
<%
else
%>
<table border=0><tr><td><font color="#FF0000">��1��</font><input type=radio value=0 name=class></td><td>������ѭ��</td></tr></table>
<%
end if
%>
</td>

<td bgcolor=#ffffee><%
if adsrs("class")=1 then
%>
<table border=0><tr>
<td><font color="#FF0000">��2��</font><input type=radio value=1 name=class checked></td>
<td>�����������</td>
<td><input type=text name=clicks1 size=8 value="<%=adsrs("clicks")%>"></td>
</tr></table>
<%
else
%>
<table border=0><tr>
<td><font color="#FF0000">��2��</font><input type=radio value=1 name=class></td>
<td>�����������</td>
<td><input type=text name=clicks1 size=8></td>
</tr></table>
<%
end if
%>
</td></tr>

<tr><td bgcolor=#ffffee><%
if adsrs("class")=2 then
%>
<table border="0"><tr>
<td><font color="#FF0000">��3��</font><input type=radio value=2 name=class checked></td>
<td>��ʾ��������</td>
<td><input type=text name=shows2 size=8 value="<%=adsrs("shows")%>"></td>
</tr></table>
<%
else
%>
<table border="0"><tr>
<td><font color="#FF0000">��3��</font><input type=radio value=2 name=class></td>
<td>��ʾ��������</td>
<td><input type=text name=shows2 size=8></td>
</tr></table>
<%
end if
%>
</td>

<td bgcolor=#ffffff><%
if adsrs("class")=3 then
%>
<table border="0"><tr>
<td><font color="#FF0000">��4��</font><input type=radio value=3 name=class checked></td>
<td>��ʾ��ֹ��Ϊ</td>
<td><input type=text name=time3 size=20 value="<%=adsrs("lasttime")%>"></td>
</tr></table>
<%
else
%>
<table border="0"><tr>
<td><font color="#FF0000">��4��</font><input type=radio value=3 name=class></td>
<td>��ʾ��ֹ��Ϊ</td>
<td><input type=text name=time3 size=20></td>
</tr></table>
<%
end if
%>
</td></tr>

<tr><td bgcolor=#ffffff colspan=2><%
if adsrs("class")=4 then
%>
<table border="0"><tr>
<td><font color="#FF0000">��5��</font><input type=radio value=4 name=class checked></td>
<td>�����������</td>
<td><input type=text name=clicks4 size=8 value="<%=adsrs("clicks")%>"></td>
<td>��ʾ��������</td>
<td><input type=text name=shows4 size=8 value="<%=adsrs("shows")%>"></td>
</tr></table>
<%
else
%>
<table border="0"><tr>
<td><font color="#FF0000">��5��</font><input type=radio value=4 name=class></td>
<td>�����������</td>
<td><input type=text name=clicks4 size=8></td>
<td>��ʾ��������</td>
<td><input type=text name=shows4 size=8></td>
</tr></table>
<%
end if
%>
</td></tr>

<tr><td bgcolor=#ffffee colspan=2><%
if adsrs("class")=5 then
%>
<table border="0"><tr>
<td><font color="#FF0000">��6��</font><input type=radio value=5 name=class checked></td>
<td>�����������</td>
<td><input type=text name=clicks5 size=8 value="<%=adsrs("clicks")%>"></td>
<td>��ʾ��ֹ��Ϊ</td>
<td><input type=text name=time5 size=20 value="<%=adsrs("lasttime")%>"></td>
</tr></table>
<%
else
%>
<table border="0"><tr>
<td><font color="#FF0000">��6��</font><input type=radio value=5 name=class></td>
<td>�����������</td>
<td><input type=text name=clicks5 size=8></td>
<td>��ʾ��ֹ��Ϊ</td>
<td><input type=text name=time5 size=20></td>
</tr></table>
<%
end if
%>
</td></tr>

<tr><td bgcolor=#ffffff colspan=2><%
if adsrs("class")=6 then
%>
<table border="0"><tr>
<td><font color="#FF0000">��7��</font><input type=radio value=6 name=class checked></td>
<td>��ʾ��������</td>
<td><input type=text name=shows6 size=8 value="<%=adsrs("shows")%>"></td>
<td>��ʾ��ֹ��Ϊ</td>
<td><input type=text name=time6 size=20 value="<%=adsrs("lasttime")%>"></td>
</tr></table>
<%
else
%>
<table border="0"><tr>
<td><font color="#FF0000">��7��</font><input type=radio value=6 name=class></td>
<td>��ʾ��������</td>
<td><input type=text name=shows6 size=8></td>
<td>��ʾ��ֹ��Ϊ</td>
<td><input type=text name=time6 size=20></td>
</tr></table>
<%
end if
%>
</td></tr>

<tr><td bgcolor=#ffffee colspan=2><%
if adsrs("class")=7 then
%>
<table border="0">
<tr>
<td><font color="#FF0000">��8��</font><input type=radio value=7 name=class checked></td>
<td>�����������</td>
<td><input type=text name=clicks7 size=8 value="<%=adsrs("clicks")%>"></td>
<td>��ʾ��������</td>
<td><input type=text name=shows7 size=8 value="<%=adsrs("shows")%>"></td>
</tr>
<tr>
<td></td>
<td>��ʾ��ֹ��Ϊ</td>
<td colspan=3><input type=text name=time7 size=20 value="<%=adsrs("lasttime")%>"></td>
</tr>
</table>
<%
else
%>
<table border="0">
<tr>
<td><font color="#FF0000">��8��</font><input type=radio value=7 name=class></td>
<td>�����������</td>
<td><input type=text name=clicks7 size=8></td>
<td>��ʾ��������</td>
<td><input type=text name=shows7 size=8></td>
</tr>
<tr>
<td></td>
<td>��ʾ��ֹ��Ϊ</td>
<td colspan=3><input type=text name=time7 size=20></td>
</tr>
</table>
<%
end if
%>
</td></tr>

<tr><td align=center bgcolor=#ffffff colspan=2>
  <input type=submit value=�ύ�޸� name=B1> <input type=reset value=��д name=B2></td></tr>

</table>

</td></tr></form></table>

</center></div>

<%
adsrs.close
set adsrs=nothing
adsconn.close
set adsconn=nothing

%>
        
</body>
</html>