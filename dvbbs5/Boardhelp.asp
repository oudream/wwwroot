<!--#include file=conn.asp-->
<!-- #include file=inc/const.asp -->
<%
	if boardid=0 then
		stats="��̳�ܰ���"
		call nav()
		call head_var(1)
	else
		stats=BoardType & "�������"
		call nav()
		call head_var(2)
	end if
	call main()
	call footer()
sub main()
%>
<table cellspacing=1 cellpadding=3 align=center class=tableborder1>
<tr> 
<td class=tablebody2 width=100% align=center>
<A HREF=#point>��������</A> | <A HREF=#grade>�ȼ�����</A> | <A HREF=#ubb>UBB�﷨</A> | <A HREF=#ubb1>UBB����</A>
</td>
</tr>
  <tr> 
    <th align=left>A. <A name=point>��������</A></th>
  </tr> 
<tr> 
<td width=100% class=tablebody1>
&nbsp;&nbsp;����̳ע�ᡢ��½�����������������뾫����ɾ�����ӵȲ������û���ֵ��Ӱ�죬�����ɸ����û���������������������Ĭ��ֵ���ܰ����ɶ��û���������ֱ�Ӳ�����<BR><BR>
��һ��<B>��Ǯ</B><BR>ע���Ǯ����<font color=<%=Forum_body(8)%>><%=Forum_user(0)%></font>&nbsp;��½���ӽ�Ǯ��<font color=<%=Forum_body(8)%>><%=Forum_user(4)%></font>&nbsp;�������ӽ�Ǯ��<font color=<%=Forum_body(8)%>><%=Forum_user(1)%></font>&nbsp;�������ӽ�Ǯ��<font color=<%=Forum_body(8)%>><%=Forum_user(2)%></font>&nbsp;�������ӽ�Ǯ��<font color=<%=Forum_body(8)%>><%=Forum_user(15)%></font>&nbsp;ɾ�����ٽ�Ǯ��<font color=<%=Forum_body(8)%>><%=Forum_user(3)%></font><BR>
<BR>
������<B>����</B><BR>ע�ᾭ������<font color=<%=Forum_body(8)%>><%=Forum_user(5)%></font>&nbsp;��½���Ӿ��飺<font color=<%=Forum_body(8)%>><%=Forum_user(9)%></font>&nbsp;�������Ӿ��飺<font color=<%=Forum_body(8)%>><%=Forum_user(6)%></font>&nbsp;�������Ӿ��飺<font color=<%=Forum_body(8)%>><%=Forum_user(7)%></font>&nbsp;�������Ӿ��飺<font color=<%=Forum_body(8)%>><%=Forum_user(17)%></font>&nbsp;ɾ�����پ��飺<font color=<%=Forum_body(8)%>><%=Forum_user(8)%></font><BR>
<BR>
������<B>����</B><BR>ע����������<font color=<%=Forum_body(8)%>><%=Forum_user(10)%></font>&nbsp;��½����������<font color=<%=Forum_body(8)%>><%=Forum_user(14)%></font>&nbsp;��������������<font color=<%=Forum_body(8)%>><%=Forum_user(11)%></font>&nbsp;��������������<font color=<%=Forum_body(8)%>><%=Forum_user(12)%></font>&nbsp;��������������<font color=<%=Forum_body(8)%>><%=Forum_user(16)%></font>&nbsp;ɾ������������<font color=<%=Forum_body(8)%>><%=Forum_user(13)%></font><BR>
</td>
</tr>
  <tr> 
    <th align=left><A name=grade>B. �ȼ�����</A></th>
  </tr> 
  <tr> 
    <td class=tablebody1>
	<p style=line-height:15pt>
����Ϊ����̳��Ӧ�ȼ�������֣��Լ���Ӧ�ĵȼ�ͼƬ��<BR><BR>
<%
	set rs=conn.execute("select * from usertitle where not Minarticle=-1 order by MinArticle")
	do while not rs.eof
	response.write "�������������� " & rs("usertitle") & " ��Ҫ " & rs("MinArticle") & " �Ļ��� �ȼ���־Ϊ <img src=pic/"&rs("titlepic")&"><br>" 
	rs.movenext
	loop
	rs.close
	set rs=nothing
%>
�����������  Ϊ����Ա�趨�����Խ����ض����档<BR>
������������  Ϊ����Ա�趨�����Զ���������̳���ӽ��й���<BR>
������������Ա  ��ӵ����̳ȫ��Ȩ�ޡ�
</p>
	</td>
  </tr> 
  <tr> 
    <th align=left>C. <A name=ubb>UBB�﷨</A></th>
  </tr> 
  <tr> 
    <td class=tablebody1>
<p>��̳�����ɹ���Ա�����Ƿ�֧��UBB��ǩ��UBB��ǩ���ǲ�����ʹ��HTML�﷨������£�ͨ����̳������ת��������������֧���������õġ���Σ���Ե�HTMLЧ����ʾ������Ϊ����ʹ��˵����
<p><font color=red>[B]</font><b>����</b><font color=red>[/B]</font>�������ֵ�λ�ÿ��������������Ҫ���ַ�����ʾΪ����Ч����
<p><font color=red>[I]</font><i>����</i><font color=red>[/I]</font>�������ֵ�λ�ÿ��������������Ҫ���ַ�����ʾΪб��Ч����
<p><font color=red>[U]</font><u>����</u><font color=red>[/U]</font>�������ֵ�λ�ÿ��������������Ҫ���ַ�����ʾΪ�»���Ч����
<p><font color=red>[align=center]</font><div align=center>����</div><font color=red>[/align]</font>�������ֵ�λ�ÿ��������������Ҫ���ַ���centerλ��center��ʾ���У�left��ʾ����right��ʾ���ҡ�
<p><font color=red>[URL]</font><A HREF=HTTP://WWW.ASPSKY.NET>HTTP://WWW.ASPSKY.NET</A><font color=red>[/URL]</font>
<P><font color=red>[URL=HTTP://WWW.ASPSKY.NET]</font><A HREF=HTTP://WWW.ASPSKY.NET>�����ȷ�</A><font color=red>[/URL]</font>�������ַ������Լ��볬�����ӣ��������Ӿ����ַ�����������ӡ�
<p><font color=red>[EMAIL]</font><A HREF=mailto:aspmaster@cmmail.com>aspmaster@cmmail.com</A><font color=red>[/EMAIL]</font>
<P><font color=red>[EMAIL=MAILTO:aspmaster@cmmail.com]</font><A HREF=mailto:aspmaster@cmmail.com>ɳ̲С��</A><font color=red>[/EMAIL]</font>�������ַ������Լ����ʼ����ӣ��������Ӿ����ַ�����������ӡ�
<P><font color=red>[img]</font><img src=http://www.aspsky.net/images/asp.gif><font color=red>[/img]</font>���ڱ�ǩ���м����ͼƬ��ַ����ʵ�ֲ�ͼЧ����
<P><font color=red>[flash]</font>Flash���ӵ�ַ<font color=red>[/Flash]</font>���ڱ�ǩ���м����FlashͼƬ��ַ����ʵ�ֲ���Flash��
<P><font color=red>[code]</font>����<font color=red>[/code]</font>���ڱ�ǩ��д�����ֿ�ʵ��html�б��Ч����
<P><font color=red>[quote]</font>����<font color=red>[/quote]</font>���ڱ�ǩ���м�������ֿ���ʵ��HTMl����������Ч����
<P><font color=red>[list]</font>����<font color=red>[/list]</font> <font color=red>[list=a]</font>����<font color=red>[/list]</font>  <font color=red>[list=1]</font>����<font color=red>[/list]</font>������list���Ա�ǩ��ʵ��HTMLĿ¼Ч����
<P><font color=red>[fly]</font>����<font color=red>[/fly]</font>���ڱ�ǩ���м�������ֿ���ʵ�����ַ���Ч������������ơ�
<P><font color=red>[move]</font>����<font color=red>[/move]</font>���ڱ�ǩ���м�������ֿ���ʵ�������ƶ�Ч����Ϊ����Ʈ����
<P><font color=red>[glow=255,red,2]</font>����<font color=red>[/glow]</font>���ڱ�ǩ���м�������ֿ���ʵ�����ַ�����Ч��glow����������Ϊ��ȡ���ɫ�ͱ߽��С��
<P><font color=red>[shadow=255,red,2]</font>����<font color=red>[/shadow]</font>���ڱ�ǩ���м�������ֿ���ʵ��������Ӱ��Ч��shadow����������Ϊ��ȡ���ɫ�ͱ߽��С��
<P><font color=red>[color=��ɫ����]</font>����<font color=red>[/color]</font>������������ɫ���룬�ڱ�ǩ���м�������ֿ���ʵ��������ɫ�ı䡣
<P><font color=red>[size=����]</font>����<font color=red>[/size]</font>���������������С���ڱ�ǩ���м�������ֿ���ʵ�����ִ�С�ı䡣
<P><font color=red>[face=����]</font>����<font color=red>[/face]</font>����������Ҫ�����壬�ڱ�ǩ���м�������ֿ���ʵ����������ת����
<P><font color=red>[DIR=500,350]</font>http://<font color=red>[/DIR]</font>��Ϊ����shockwave��ʽ�ļ����м������Ϊ��Ⱥͳ���
<P><font color=red>[RM=500,350]</font>http://<font color=red>[/RM]</font>��Ϊ����realplayer��ʽ��rm�ļ����м������Ϊ��Ⱥͳ���
<P><font color=red>[MP=500,350]</font>http://<font color=red>[/MP]</font>��Ϊ����Ϊmidia player��ʽ���ļ����м������Ϊ��Ⱥͳ���
<P><font color=red>[QT=500,350]</font>http://<font color=red>[/QT]</font>��Ϊ����ΪQuick time��ʽ���ļ����м������Ϊ��Ⱥͳ���
	</td>
  </tr> 
  <tr> 
    <th align=left>D. <A name=ubb1>UBB����</A></th>
  </tr> 
  <tr> 
    <td class=tablebody1>
&nbsp;&nbsp; ����Ϊ����̳��UBB�﷨���ã�ͨ����Щ���ã�������֪���ڱ����淢��������Щ����ǲ���ʹ�õģ����ﻹ�����˿����û�ǩ����ʹ�õ�UBBѡ�<BR>
&nbsp;&nbsp;<B>�û�����</B>��
<ul>
<li>HTML��ǩ�� <%=iif(Forum_Setting(35),"����","������")%>
<li>UBB��ǩ�� <%=iif(Forum_Setting(34),"����","������")%>
<li>��ͼ��ǩ�� <%=iif(Forum_Setting(37),"����","������")%>
<li>Flash��ǩ��<%=iif(Forum_Setting(38),"����","������")%>
<li>�����ַ�ת����<%=iif(Forum_Setting(36),"����","������")%>
<li>�ϴ�ͼƬ��<%=iif(Forum_Setting(3),"����","������")%>
<li>���<%=Forum_Setting(13)\1024%>KB 
</ul>
<BR>&nbsp;&nbsp;<B>�û�ǩ��</B>��
<ul>
��������<li>HTML��ǩ�� <%=iif(Forum_Setting(66),"����","������")%>
��������<li>UBB��ǩ�� <%=iif(Forum_Setting(65),"����","������")%>
��������<li>��ͼ��ǩ(����ͼƬ��flash����ý��)�� <%=iif(Forum_Setting(67),"����","������")%>
</ul>
˵��������html��ǩָ�Ƿ�����ʹ��html�﷨����ͼ��flash�Լ������ַ�ת��������UBB�﷨���ݣ���ʹ�÷����ɲ鿴UBB�﷨
	</td>
  </tr> 
</table>

<%
	call activeonline()
end sub
%>