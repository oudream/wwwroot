<%@ LANGUAGE = VBScript %>
<%  Option Explicit %>
<SCRIPT  LANGUAGE = VBScript>
      Sub showsessionID()
           MsgBox "���SessionID��:" & <%= Session.SessionID %>
      End Sub
</SCRIPT>
<HTML>
     <HEAD>
         <TITLE>��ʾSessionID</TITLE>
     </HEAD>
     <BODY BGCOLOR="White" TOPMARGIN="10" LEFTMARGIN="10">
           <!--  Display header. -->
           <B>��ʵ����ʾSessionID</B><P> 
           <INPUT TYPE=Button VALUE="����˴�" ONCLICK=showsessionID>
     </BODY>
</HTML>
