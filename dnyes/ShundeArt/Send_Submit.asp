<%@ Language=VBScript %>
<!--#include file=conn.asp -->
<!--#include file=function.asp -->
<!--#include file=config.asp -->
<!--#include file="char.inc"-->
<%
NewsID=checkstr(Request.Form("NewsID"))
name=trim(Request.Form("name"))
myname=trim(Request.Form("myname"))
toemail=trim(Request.Form("email"))
reffer=trim(Request.Form("reffer"))

set rs=server.CreateObject("ADODB.RecordSet")
rs.Source="select * from "& db_News_Table &" where NewsID=" & NewsID
rs.Open rs.Source,conn,1,1

Title=trim(rs("Title"))
Author=trim(rs("Author"))
Original=trim(rs("Original"))
image=rs("image")
UpdateTime=trim(rs("UpdateTime"))
Content=trim(rs("Content"))
SpecialID=rs("SpecialID")
click=rs("click")

content=htmlencode1(content)
Content="&nbsp;&nbsp;&nbsp;&nbsp;" & Content

rs.Close

Set objCDOMail = Server.CreateObject("CDONTS.NewMail")
With objCDOMail
.From = ""& email &""
.To = toemail
.Subject = ""&name&"您好，您的朋友"& myname &"给你发的文章-" & now()
.BodyFormat = 0
.MailFormat = 0
.Importance = 1 

BodyText="<body>"
BodyText=BodyText & "<table border=1 style=""border-collapse: collapse; font-size: 9pt"" bordercolor="&border&" width=100% align=center cellspacing=0 cellpadding=5>"
BodyText=BodyText & "<tr><td width=100% bgcolor="&m_top&" align=center><font color=#FF9900>" & redcaff & "</font></td></tr>"
BodyText=BodyText & "<tr><td width=100% align=center><b>" & server.HTMLEncode(Title) &"</b></td></tr>"
BodyText=BodyText & "<tr><td width=100% align=center><b>" & updateTime & "&nbsp;&nbsp;" & original &"&nbsp;&nbsp;" & author &"</b></td></tr>"
BodyText=BodyText & "<tr><td width=100% ><a href=" & reffer & " target=""_blank"">点击这里查看原始文件</a></td></tr>"
BodyText=BodyText & "<tr><td width=100% >"& content & "</td></tr>"
BodyText=BodyText & "<tr><td width=100% bgcolor="&m_top&" align=center><a href=http://" & xpurl &" target=_blank><font color=blue>http://" & xpurl &" </font></a><br><a href=mailto:" & email &" ><font color=blue>" & email &"</font></a></td></tr>"
BodyText=BodyText & "</table></body></html>"

.Body = bodytext
.Send	
end with
Set objCDOMail = Nothing
%>
<script language=javascript>
alert("已经将该新闻发送给您的朋友！")
</script>
<body  onload="javascript:window.close()">
</body>