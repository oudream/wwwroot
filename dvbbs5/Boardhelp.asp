<!--#include file=conn.asp-->
<!-- #include file=inc/const.asp -->
<%
	if boardid=0 then
		stats="论坛总帮助"
		call nav()
		call head_var(1)
	else
		stats=BoardType & "版面帮助"
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
<A HREF=#point>积分设置</A> | <A HREF=#grade>等级设置</A> | <A HREF=#ubb>UBB语法</A> | <A HREF=#ubb1>UBB设置</A>
</td>
</tr>
  <tr> 
    <th align=left>A. <A name=point>积分设置</A></th>
  </tr> 
<tr> 
<td width=100% class=tablebody1>
&nbsp;&nbsp;该论坛注册、登陆、发贴、回帖、加入精华、删除帖子等操作对用户分值的影响，版主可根据用户发贴表现自行增减以下默认值，总版主可对用户威力进行直接操作：<BR><BR>
（一）<B>金钱</B><BR>注册金钱数：<font color=<%=Forum_body(8)%>><%=Forum_user(0)%></font>&nbsp;登陆增加金钱：<font color=<%=Forum_body(8)%>><%=Forum_user(4)%></font>&nbsp;发帖增加金钱：<font color=<%=Forum_body(8)%>><%=Forum_user(1)%></font>&nbsp;跟帖增加金钱：<font color=<%=Forum_body(8)%>><%=Forum_user(2)%></font>&nbsp;精华增加金钱：<font color=<%=Forum_body(8)%>><%=Forum_user(15)%></font>&nbsp;删帖减少金钱：<font color=<%=Forum_body(8)%>><%=Forum_user(3)%></font><BR>
<BR>
（二）<B>经验</B><BR>注册经验数：<font color=<%=Forum_body(8)%>><%=Forum_user(5)%></font>&nbsp;登陆增加经验：<font color=<%=Forum_body(8)%>><%=Forum_user(9)%></font>&nbsp;发帖增加经验：<font color=<%=Forum_body(8)%>><%=Forum_user(6)%></font>&nbsp;跟帖增加经验：<font color=<%=Forum_body(8)%>><%=Forum_user(7)%></font>&nbsp;精华增加经验：<font color=<%=Forum_body(8)%>><%=Forum_user(17)%></font>&nbsp;删帖减少经验：<font color=<%=Forum_body(8)%>><%=Forum_user(8)%></font><BR>
<BR>
（二）<B>魅力</B><BR>注册魅力数：<font color=<%=Forum_body(8)%>><%=Forum_user(10)%></font>&nbsp;登陆增加魅力：<font color=<%=Forum_body(8)%>><%=Forum_user(14)%></font>&nbsp;发帖增加魅力：<font color=<%=Forum_body(8)%>><%=Forum_user(11)%></font>&nbsp;跟帖增加魅力：<font color=<%=Forum_body(8)%>><%=Forum_user(12)%></font>&nbsp;精华增加魅力：<font color=<%=Forum_body(8)%>><%=Forum_user(16)%></font>&nbsp;删帖减少魅力：<font color=<%=Forum_body(8)%>><%=Forum_user(13)%></font><BR>
</td>
</tr>
  <tr> 
    <th align=left><A name=grade>B. 等级设置</A></th>
  </tr> 
  <tr> 
    <td class=tablebody1>
	<p style=line-height:15pt>
以下为该论坛相应等级所需积分，以及相应的等级图片：<BR><BR>
<%
	set rs=conn.execute("select * from usertitle where not Minarticle=-1 order by MinArticle")
	do while not rs.eof
	response.write "升级到 " & rs("usertitle") & " 需要 " & rs("MinArticle") & " 的积分 等级标志为 <img src=pic/"&rs("titlepic")&"><br>" 
	rs.movenext
	loop
	rs.close
	set rs=nothing
%>
贵宾  为管理员设定，可以进入特定版面。<BR>
版主  为管理员设定，可以对其管理的论坛帖子进行管理。<BR>
管理员  ，拥有论坛全部权限。
</p>
	</td>
  </tr> 
  <tr> 
    <th align=left>C. <A name=ubb>UBB语法</A></th>
  </tr> 
  <tr> 
    <td class=tablebody1>
<p>论坛可以由管理员设置是否支持UBB标签，UBB标签就是不允许使用HTML语法的情况下，通过论坛的特殊转换程序，以至可以支持少量常用的、无危害性的HTML效果显示。以下为具体使用说明：
<p><font color=red>[B]</font><b>文字</b><font color=red>[/B]</font>：在文字的位置可以任意加入您需要的字符，显示为粗体效果。
<p><font color=red>[I]</font><i>文字</i><font color=red>[/I]</font>：在文字的位置可以任意加入您需要的字符，显示为斜体效果。
<p><font color=red>[U]</font><u>文字</u><font color=red>[/U]</font>：在文字的位置可以任意加入您需要的字符，显示为下划线效果。
<p><font color=red>[align=center]</font><div align=center>文字</div><font color=red>[/align]</font>：在文字的位置可以任意加入您需要的字符，center位置center表示居中，left表示居左，right表示居右。
<p><font color=red>[URL]</font><A HREF=HTTP://WWW.ASPSKY.NET>HTTP://WWW.ASPSKY.NET</A><font color=red>[/URL]</font>
<P><font color=red>[URL=HTTP://WWW.ASPSKY.NET]</font><A HREF=HTTP://WWW.ASPSKY.NET>动网先锋</A><font color=red>[/URL]</font>：有两种方法可以加入超级连接，可以连接具体地址或者文字连接。
<p><font color=red>[EMAIL]</font><A HREF=mailto:aspmaster@cmmail.com>aspmaster@cmmail.com</A><font color=red>[/EMAIL]</font>
<P><font color=red>[EMAIL=MAILTO:aspmaster@cmmail.com]</font><A HREF=mailto:aspmaster@cmmail.com>沙滩小子</A><font color=red>[/EMAIL]</font>：有两种方法可以加入邮件连接，可以连接具体地址或者文字连接。
<P><font color=red>[img]</font><img src=http://www.aspsky.net/images/asp.gif><font color=red>[/img]</font>：在标签的中间插入图片地址可以实现插图效果。
<P><font color=red>[flash]</font>Flash连接地址<font color=red>[/Flash]</font>：在标签的中间插入Flash图片地址可以实现插入Flash。
<P><font color=red>[code]</font>文字<font color=red>[/code]</font>：在标签中写入文字可实现html中编号效果。
<P><font color=red>[quote]</font>引用<font color=red>[/quote]</font>：在标签的中间插入文字可以实现HTMl中引用文字效果。
<P><font color=red>[list]</font>文字<font color=red>[/list]</font> <font color=red>[list=a]</font>文字<font color=red>[/list]</font>  <font color=red>[list=1]</font>文字<font color=red>[/list]</font>：更改list属性标签，实现HTML目录效果。
<P><font color=red>[fly]</font>文字<font color=red>[/fly]</font>：在标签的中间插入文字可以实现文字飞翔效果，类似跑马灯。
<P><font color=red>[move]</font>文字<font color=red>[/move]</font>：在标签的中间插入文字可以实现文字移动效果，为来回飘动。
<P><font color=red>[glow=255,red,2]</font>文字<font color=red>[/glow]</font>：在标签的中间插入文字可以实现文字发光特效，glow内属性依次为宽度、颜色和边界大小。
<P><font color=red>[shadow=255,red,2]</font>文字<font color=red>[/shadow]</font>：在标签的中间插入文字可以实现文字阴影特效，shadow内属性依次为宽度、颜色和边界大小。
<P><font color=red>[color=颜色代码]</font>文字<font color=red>[/color]</font>：输入您的颜色代码，在标签的中间插入文字可以实现文字颜色改变。
<P><font color=red>[size=数字]</font>文字<font color=red>[/size]</font>：输入您的字体大小，在标签的中间插入文字可以实现文字大小改变。
<P><font color=red>[face=字体]</font>文字<font color=red>[/face]</font>：输入您需要的字体，在标签的中间插入文字可以实现文字字体转换。
<P><font color=red>[DIR=500,350]</font>http://<font color=red>[/DIR]</font>：为插入shockwave格式文件，中间的数字为宽度和长度
<P><font color=red>[RM=500,350]</font>http://<font color=red>[/RM]</font>：为插入realplayer格式的rm文件，中间的数字为宽度和长度
<P><font color=red>[MP=500,350]</font>http://<font color=red>[/MP]</font>：为插入为midia player格式的文件，中间的数字为宽度和长度
<P><font color=red>[QT=500,350]</font>http://<font color=red>[/QT]</font>：为插入为Quick time格式的文件，中间的数字为宽度和长度
	</td>
  </tr> 
  <tr> 
    <th align=left>D. <A name=ubb1>UBB设置</A></th>
  </tr> 
  <tr> 
    <td class=tablebody1>
&nbsp;&nbsp; 下面为本论坛的UBB语法设置，通过这些设置，您可以知道在本版面发言中有哪些语句是不能使用的，这里还包括了控制用户签名里使用的UBB选项。<BR>
&nbsp;&nbsp;<B>用户发贴</B>：
<ul>
<li>HTML标签： <%=iif(Forum_Setting(35),"可用","不可用")%>
<li>UBB标签： <%=iif(Forum_Setting(34),"可用","不可用")%>
<li>贴图标签： <%=iif(Forum_Setting(37),"可用","不可用")%>
<li>Flash标签：<%=iif(Forum_Setting(38),"可用","不可用")%>
<li>表情字符转换：<%=iif(Forum_Setting(36),"可用","不可用")%>
<li>上传图片：<%=iif(Forum_Setting(3),"可用","不可用")%>
<li>最多<%=Forum_Setting(13)\1024%>KB 
</ul>
<BR>&nbsp;&nbsp;<B>用户签名</B>：
<ul>
<li>HTML标签： <%=iif(Forum_Setting(66),"可用","不可用")%>
<li>UBB标签： <%=iif(Forum_Setting(65),"可用","不可用")%>
<li>帖图标签(包括图片、flash、多媒体)： <%=iif(Forum_Setting(67),"可用","不可用")%>
</ul>
说明：这里html标签指是否允许使用html语法，贴图和flash以及表情字符转换都属于UBB语法内容，其使用方法可查看UBB语法
	</td>
  </tr> 
</table>

<%
	call activeonline()
end sub
%>