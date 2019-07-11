<%@ LANGUAGE = VBScript %>
<%  Option Explicit %>
<SCRIPT  LANGUAGE = VBScript>
      Sub showsessionID()
           MsgBox "你的SessionID是:" & <%= Session.SessionID %>
      End Sub
</SCRIPT>
<HTML>
     <HEAD>
         <TITLE>显示SessionID</TITLE>
     </HEAD>
     <BODY BGCOLOR="White" TOPMARGIN="10" LEFTMARGIN="10">
           <!--  Display header. -->
           <B>本实例显示SessionID</B><P> 
           <INPUT TYPE=Button VALUE="点击此处" ONCLICK=showsessionID>
     </BODY>
</HTML>
