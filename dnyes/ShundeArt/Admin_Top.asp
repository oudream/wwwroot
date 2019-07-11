<table border="1" cellpadding="6" cellspacing="0" style="border-collapse: collapse" bordercolor="#C0C0C0" width="100%" id="AutoNumber1">
<tr>
<td background="IMAGES/top1.gif">
	<table border="0" cellpadding="6" cellspacing="0" style="border-collapse: collapse" bordercolor="#C0C0C0" width="100%" id="AutoNumber1">
        <tr> 
          <td background="IMAGES/top1.gif">&nbsp;</td>
        </tr>
        <tr> 
          <td background="IMAGES/top1.gif" class="td001"> <font color="#FFFFFF">
		 用户：<b><%=Request.cookies(Forcast_SN)("UserName")%></b>
			<%if db_Birthday_Select="FeiTium" then		'性别字段是原字段%>
				<%=Request.cookies(Forcast_SN)("sex")%>
			<%else
				if db_Birthday_Select="Text" then		'性别字段是BBS的文本型阿拉伯数字
					if Request.cookies(Forcast_SN)("sex")=1 then%>
						男
					<%else
						if Request.cookies(Forcast_SN)("sex")=0 then%>
							女
						<%else%>
							保密
						<%end if
					end if
				end if
			end if%>
					您的权限：<%if Request.cookies(Forcast_SN)("KEY")="super" and Request.cookies(Forcast_SN)("purview")="99999" then%><font color="#FFCC00">超级管理员</font><%end if%>
						<%if Request.cookies(Forcast_SN)("KEY")="super" and Request.cookies(Forcast_SN)("purview")<>"99999" then%><font color="#FFCC00">系统管理员</font><%end if%>
						<%if Request.cookies(Forcast_SN)("KEY")="check" then%><font color="#FFCC00">文章审核员</font><%end if%>
						<%if Request.cookies(Forcast_SN)("KEY")="selfreg" then%><font color="#FFCC00">注册用户</font><%end if%>
						<%if Request.cookies(Forcast_SN)("KEY")="smallmaster" then%><font color="#FFCC00">小类管理员</font><%end if%>
						<%if Request.cookies(Forcast_SN)("KEY")="bigmaster" then%><font color="#FFCC00">大类管理员</font><%end if%>
						<%if Request.cookies(Forcast_SN)("KEY")="typemaster" then%><font color="#FFCC00">总栏管理员</font><%end if%>
					您的等级：<%if Request.cookies(Forcast_SN)("KEY")<>"selfreg" then%><font color="#FFCC00">内部成员</font><%end if%>
						<%if Request.cookies(Forcast_SN)("KEY")="selfreg" and Request.cookies(Forcast_SN)("reglevel")="1" then%><font color="red">普通</font><%end if%>
						<%if Request.cookies(Forcast_SN)("KEY")="selfreg" and Request.cookies(Forcast_SN)("reglevel")="2" then%><font color="red">高级</font><%end if%>
						<%if Request.cookies(Forcast_SN)("KEY")="selfreg" and Request.cookies(Forcast_SN)("reglevel")="3" then%><font color="red">特级</font><%end if%>
						登录次数：<%=Request.cookies(Forcast_SN)("UserLoginTimes")%>
		  </font>
        </td>
        </tr>
      </table>
</td>
</tr>
<tr>
	<td bgcolor="#7C7CB5">
      <div align="right">
	  <A href="Admin_Exit.asp"><FONT color=#ffffff>退出网站</FONT></A>&nbsp;&nbsp;
	  <A onclick="checkclick('您是否要打开本站首页？')" href="index.asp" target=_blank><FONT color=#ffffff>网站首页</FONT></A></div></td>
</tr>
</table>
