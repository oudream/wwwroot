<!-- #include file="conn.asp" -->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/char_board.asp" -->
<%
	if BoardID="" or not isInteger(BoardID) or BoardID=0 then
		Errmsg=Errmsg+"<br>"+"<li>错误的版面参数！请确认您是从有效的连接进入。"
		founderr=true
	else
		BoardID=clng(BoardID)
	end if
	if founderr then
		call nav()
		call head_var(1)
		call dvbbs_error()
	else
		stats=boardtype&"发表投票"
		call nav()
		call head_var(2)
		call main()
		if founderr then call dvbbs_error()
	end if
	call footer()
sub main()
	if Cint(GroupSetting(8))=0 then
		Errmsg=Errmsg+"<br>"+"<li>您没有在本论坛发表投票的权限，请<a href=login.asp>登陆</a>或者同管理员联系。"
		founderr=true
		exit sub
	end if
	if cint(lockboard)=2 then
		if not founduser then
			Errmsg=Errmsg+"<br>"+"<li>本论坛为认证论坛，请<a href=login.asp>登陆</a>并确认您的用户名已经得到管理员的认证后进入。"
			founderr=true
			exit sub
		else
			if chkboardlogin(boardid,membername)=false then
			Errmsg=Errmsg+"<br>"+"<li>本论坛为认证论坛，请确认您的用户名已经得到管理员的认证后进入。"
			founderr=true
			exit sub
			end if
		end if
	end if
%>
<script src="inc/ubbcode.js"></script>
<form action=Savevote.asp?boardID=<%=request("boardid")%> method=POST onSubmit=submitonce(this) name=frmAnnounce>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
    <tr>
      <th width=100% colspan=2 align=left>&nbsp;&nbsp;发表新投票 </th>
    </tr>
        <tr>
          <td width=20% class=tablebody1><b>用户名</b></td>
          <td width=80% class=tablebody1><input name=username value=<%=membername%> class=FormClass>&nbsp;&nbsp;<font color=<%=Forum_body(8)%>><b>*</b></font><a href=reg.asp>您没有注册？</a> 
          </td>
        </tr>
        <tr>
          <td width=20% class=tablebody1><b>密码</b></td>
          <td width=80% class=tablebody1><input name=passwd type=password value=<%=htmlencode(memberword)%> class=FormClass><font color=<%=Forum_body(8)%>>&nbsp;&nbsp;<b>*</b></font><a href=lostpass.asp>忘记密码？</a></td>
        </tr>
        <tr>
          <td width=20% class=tablebody1><b>主题标题</b>
              <SELECT name=font onchange=DoTitle(this.options[this.selectedIndex].value)>
              <OPTION selected value="">选择话题</OPTION> <OPTION value=[原创]>[原创]</OPTION> 
              <OPTION value=[转帖]>[转帖]</OPTION> <OPTION value=[灌水]>[灌水]</OPTION> 
              <OPTION value=[讨论]>[讨论]</OPTION> <OPTION value=[求助]>[求助]</OPTION> 
              <OPTION value=[推荐]>[推荐]</OPTION> <OPTION value=[公告]>[公告]</OPTION> 
              <OPTION value=[注意]>[注意]</OPTION> <OPTION value=[贴图]>[贴图]</OPTION>
              <OPTION value=[建议]>[建议]</OPTION> <OPTION value=[下载]>[下载]</OPTION>
              <OPTION value=[分享]>[分享]</OPTION></SELECT>
	  </td>
		            <td width=80% class=tablebody1><input name=subject size=50 maxlength=80 class=FormClass>&nbsp;
<select name=votetimeout size=1>
<option value=0>过期时间</option>
<option value=0>永不过期</option>
<option value=1>一天</option>
<option value=3>三天</option>
<option value=7>一周</option>
<option value=15>半月</option>
<option value=30>一月</option>
<option value=90>三月</option>
<option value=180>半年</option>
</select>
&nbsp;<font color=<%=Forum_body(8)%>><b>*</b></font>不得超过25个汉字或50个英文字符
	 </td>
        </tr>
        <tr>
          <td width=20% class=tablebody1><b>投票项目 </b> <br>
        <br>
        <li>每行一个投票项目，最多<b><%=Forum_Setting(39)%></b>个</li>
        <li>超过自动作废，空行自动过滤</li><br>
        <br>
        <input type=radio name=votetype value=0 checked>
          单选投票<br>
          <input type=radio name=votetype value=1>
          多选投票</font></td>
	  <td width=80% class=tablebody1><textarea name=vote cols=95 rows=8></textarea>
	 </td>
        </tr>
        <tr>
          <td width=20% valign=top class=tablebody1><b>当前心情</b><br><li>将放在帖子的前面</td>
          <td width=80% class=tablebody1>
<%for i=1 to 9%>
	<input type="radio" value="face<%=i%>" name="Expression" <%if i=1 then response.write "checked"%>><img src="face/face<%=i%>.gif" WIDTH="15" HEIGHT="15">&nbsp;&nbsp;
<%next%>
<br>
<%for i=10 to 18%>
	<input type="radio" value="face<%=i%>" name="Expression"><img src="face/face<%=i%>.gif" WIDTH="15" HEIGHT="15">&nbsp;&nbsp;
<%next%>
 </td>
        </tr>
        <tr>
          <td width=20% valign=top class=tablebody1>
<b>内容</b><br>
<!--标签状态-->
<li>HTML标签： <%=iif(Forum_Setting(35),"可用","不可用")%>
<li>UBB标签： <%=iif(Forum_Setting(34),"可用","不可用")%>
<li>贴图标签： <%=iif(Forum_Setting(37),"可用","不可用")%>
<li>Flash标签：<%=iif(Forum_Setting(38),"可用","不可用")%>
<li>表情字符转换：<%=iif(Forum_Setting(36),"可用","不可用")%>
<li>上传图片：<%=iif(Forum_Setting(3),"可用","不可用")%>
<li>最多<%=Forum_Setting(13)\1024%>KB <BR><BR>
<B>特殊内容</B><BR>
<li><a href="javascript:money()" title="使用语法：[money=可浏览该部分内容需要金钱数]内容[/money]">金钱帖</a>
<li><a href="javascript:point()" title="使用语法：[point=可浏览该部分内容需要金钱数]内容[/point]">积分帖</a>
<li><a href="javascript:usercp()" title="使用语法：[usercp=可浏览该部分内容需要金钱数]内容[/usercp]">魅力帖</a>
<li><a href="javascript:power()" title="使用语法：[power=可浏览该部分内容需要金钱数]内容[/power]">威望帖</a>
<li><a href="javascript:article()" title="使用语法：[post=可浏览该部分内容需要金钱数]内容[/post]">文章帖</a>
<li><a href="javascript:replyview()" title="使用语法：[replyview]该部分内容回复后可见[/replyview]">回复可见帖</a>
	  </td>
          <td width=80% class=tablebody1>
<%if Cint(GroupSetting(7))=1 then%>
<iframe name="ad" frameborder=0 width=100% height=30 scrolling=no src=saveannounce_upload.asp?boardid=<%=boardid%>></iframe>
<%end if%>
<%if Forum_Setting(17)=1 then%>
<!--#include file="inc/getubb.asp"-->
<%end if%>
<br>
<textarea class=smallarea cols=95 name=Content rows=12 wrap=VIRTUAL title=可以使用Ctrl+Enter直接提交贴子 class=FormClass onkeydown=ctlent()></textarea>
          </td>
        </tr>

		<tr>
                <td class=tablebody1 valign=top colspan=2><b>点击表情图即可在帖子中加入相应的表情</B><br>&nbsp;
<%
for i=1 to 28
	if len(i)=1 then i="0" & i
	response.write "<img src=""pic/em"&i&".gif"" border=0 onclick=""insertsmilie('[em"&i&"]')"" style=""CURSOR: hand"">&nbsp;"
next
%>
    		</td>
                </tr>
                <tr>
                <td valign=top class=tablebody1><b>选项</b></td>
                <td valign=middle class=tablebody1><input type=checkbox name=signflag value=yes checked>是否显示您的签名？<br>
                <input type=checkbox name=emailflag value=yes>有回复时使用邮件通知您？
<BR><BR></td>
                </tr><tr>
                <td valign=middle colspan=2 align=center class=tablebody2>
                <input type=Submit value='发 表' name=Submit> &nbsp; <input type=button value='预 览' name=Button onclick=gopreview()>&nbsp;
<input type=reset name=Submit2 value='清 除'>
                </td></form></tr>
      </table>
</form>
<form name=preview action=preview.asp method=post target=preview_page>
<input type=hidden name=title value=><input type=hidden name=body value=>
</form>
<%
	call activeonline()
end sub
%>