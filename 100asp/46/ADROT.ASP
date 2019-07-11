<%@ LANGUAGE = VBScript %>
<%  Option Explicit %>
<HTML>
     <HEAD>
         <TITLE>动态广告条</TITLE>
     </HEAD>
     <BODY BGCOLOR="White" TOPMARGIN="10" LEFTMARGIN="10">
         <FONT SIZE="4">
         <B>动态广告条</B></FONT>
        <center>
        <%
           Dim objAd 
           Set objAd = Server.CreateObject("MSWC.AdRotator")
           %>
           <%= objAd.GetAdvertisement("adrot.txt") %><BR>
</center>
<HR SIZE="1" COLOR="#000000">
<Center><A HREF="AdRot.asp">重新访问本页</A></center>
     </BODY>
</HTML>
