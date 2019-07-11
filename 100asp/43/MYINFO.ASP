<%@ Language=VBScript %>
<HTML>
<BODY>
<OBJECT RUNAT=server PROGID=MSWC.MyInfo id=mInfo> </OBJECT>
更改前信息内容：
<table border=1 align=center>
<tr><td>我的姓名:</td><td><%=mInfo.PersonalName%></td></tr>
<tr><td>个人电话:</td><td><%=mInfo.PersonalPhone%></td></tr>
<tr><td>E-mail:</td><td><%=mInfo.PersonalMail%></td></tr>
<tr><td>所在公司:</td><td><%=mInfo.CompanyName%></td></tr>
<tr><td>公司电话:</td><td><%=mInfo.CompanyPhone%></td></tr>
</table>
<%
mInfo.PersonalName="小张"
mInfo.PersonalPhone="11111111"
mInfo.PersonalMail="myemail@abc.com.cn"
mInfo.CompanyName="我们公司名称"
mInfo.CompanyPhone="22222222"
%>
更改后信息内容：
<table border=1 align=center>
<tr><td>我的姓名:</td><td><%=mInfo.PersonalName%></td></tr>
<tr><td>个人电话:</td><td><%=mInfo.PersonalPhone%></td></tr>
<tr><td>E-mail:</td><td><%=mInfo.PersonalMail%></td></tr>
<tr><td>所在公司:</td><td><%=mInfo.CompanyName%></td></tr>
<tr><td>公司电话:</td><td><%=mInfo.CompanyPhone%></td></tr>
</table>
</BODY>
</HTML>
