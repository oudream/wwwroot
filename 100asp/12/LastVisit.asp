<%@ LANGUAGE = VBScript %>
<%  Option Explicit %>
<%
      'Cookies通过HTTP Headers来从服务器端返回到浏览器上.
      '在发送Cookies之前,不能向浏览器端发送任何数据.
      Response.Expires = 0
      '从Cookie中取出上一次访问的日期和时间
      Dim LastVisit
      LastVisit = Request.Cookies("LastVisitCookie")
      Response.Cookies("LastVisitCookie") = FormatDateTime(NOW)
%>
<HTML>
     <HEAD>
           <TITLE>上次访问时间</TITLE>
     </HEAD>
     <BODY BGCOLOR="White" TOPMARGIN="10" LEFTMARGIN="10">
     <FONT SIZE="4" FACE="ARIAL, HELVETICA">
     <B>使用Cookies</B></FONT><BR>
         <HR SIZE="1" COLOR="#000000">
           <%           
                If (LastVisit = "") Then
                     '如果Cookie从未被写过,则用户是第一次访问本页
                     Response.Write("欢迎光临本页")
                Else
                     '显示上一次访问日期及时间
                     Response.Write("你上一次访问本页在" + LastVisit)
                End If 
           %>
           <P><A HREF="LastVisit.asp">重新访问本页</A>
      </BODY>
</HTML>
