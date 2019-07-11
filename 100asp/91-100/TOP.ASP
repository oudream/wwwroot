<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY bgcolor=#c1f7d8>
<br>
<base target=main>
<table border=1 width=100%>
  <tr>
    <td width=15% align=center>
    <%Response.Write session("name") %>
    </td>
    <td width=17% align=center><a href=changepassword.htm>修改密码</a>
    </td>
    <td width=17% align=center><a href=readmessage.asp>查看留言</a>
    </td>
    <td width=17% align=center><a href=showlist.asp>浏览标题</a>
    </td>
    <td width=17% align=center><a href=publisharticle.htm>发表文章</a>
    </td>
    <td width=17% align=center><a href="http://fang/bbs/" target=_parent>返回主页</a>
    </td>
  </tr>
  <tr>
    <td colspan=6 align=center>欢迎你的光临</td>
  </tr>
</table>
</BODY>
</HTML>
