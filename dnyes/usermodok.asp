<!--#include file="mymefaq5411jkjkh/favorible/showme.asp"-->
<%on error resume next%>
<%
if session("y_it_uid")="" then
response.redirect "index.asp"
response.end
end if
%>
<%
'���а�ȫ�Լ�⣬��������Դ�Ƿ��Ǳ�������ҳ��
if not instr(1,Request.ServerVariables("http_Referer"),Request.ServerVariables ("SERVER_NAME"),1)=8 then
response.write "<b>�벻Ҫ�ӷǱ���������ҳ���ύ��Ϣ</b>"
'response.redirect "error.asp?err=001"
response.end
end if
%>
<%
'���ύ������ֵ�����鿴�Ƿ��зǷ��Է�

email=request("email")
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

if age="" then age=" "
if namee="" then namee=" "
if contacte="" then contacte=" "
if gove="" then gove=" "
if shenge="" then shenge=" "
if citye="" then citye=" "
if dizhie="" then dizhie=" "
if tel2="" then tel2=" "
if fax="" then fax=" "
if oicq="" then oicq=" "


conn.execute "update user set email='"&email&"',sex='"&sex&"',age='"&age&"',namez='"&namez&"',namee='"&namee&"',contactz='"&contactz&"',contacte='"&contacte&"',govz='"&govz&"',gove='"&gove&"',shengz='"&shengz&"',shenge='"&shenge&"',cityz='"&cityz&"',citye='"&citye&"',dizhiz='"&dizhiz&"',dizhie='"&dizhie&"',tel='"&tel&"',tel2='"&tel2&"',postnumber='"&postnumber&"',oicq='"&oicq&"',fax='"&fax&"' where uid='"&session("y_it_uid")&"'"
%>
<%
response.Redirect("usermodokok.asp")
%>