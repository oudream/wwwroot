<%@ LANGUAGE = VBScript %>
<%  Option Explicit %>
<%
      'Cookiesͨ��HTTP Headers���ӷ������˷��ص��������.
      '�ڷ���Cookies֮ǰ,������������˷����κ�����.
      Response.Expires = 0
      '��Cookie��ȡ����һ�η��ʵ����ں�ʱ��
      Dim LastVisit
      LastVisit = Request.Cookies("LastVisitCookie")
      Response.Cookies("LastVisitCookie") = FormatDateTime(NOW)
%>
<HTML>
     <HEAD>
           <TITLE>�ϴη���ʱ��</TITLE>
     </HEAD>
     <BODY BGCOLOR="White" TOPMARGIN="10" LEFTMARGIN="10">
     <FONT SIZE="4" FACE="ARIAL, HELVETICA">
     <B>ʹ��Cookies</B></FONT><BR>
         <HR SIZE="1" COLOR="#000000">
           <%           
                If (LastVisit = "") Then
                     '���Cookie��δ��д��,���û��ǵ�һ�η��ʱ�ҳ
                     Response.Write("��ӭ���ٱ�ҳ")
                Else
                     '��ʾ��һ�η������ڼ�ʱ��
                     Response.Write("����һ�η��ʱ�ҳ��" + LastVisit)
                End If 
           %>
           <P><A HREF="LastVisit.asp">���·��ʱ�ҳ</A>
      </BODY>
</HTML>
