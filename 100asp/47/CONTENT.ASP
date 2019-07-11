<Html>
<Head>
<Title>
随机目录条
</Title>
</Head>
<Body>
--------------下面随机显示的内容---------------<br>
<%
set NextTip=Server.CreateObject("MSWC.ContentRotator")
%>
<%=NextTip.ChooseContent("Content.txt")%>
<br>
-------下面是Content.txt中所有的内容-------------
<%=NextTip.GetAllContent("content.txt")%>
</Body>
</Html>