<script language="JavaScript" src="../script/ChangeStyle.js"></script>
<script language="JavaScript">
	var CCNS_program = "<%=ProjectName%>";
	var CCNS_version = "<%=Version%>";
</script><script language="JavaScript">
	function doOver(o,n){
		var i;
		var o1 = document.getElementsByName("MainMenu");
		var o2 = document.getElementsByName("MenuContent");
		for(i=0;i<o1.length;++i){
			if(o1[i].className=="MenuItemOnSel"){o1[i].className="MenuItem";o2[i].style.display='none';}
			if(i==n){o1[i].className="MenuItemOnSel";o2[i].style.display='';}
		}
		o1=o2=null;
	}
</script>
<table width="760" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="170"><img src="../images/cccounter6.gif" width="162" height="67" class="HeadTable"></td>
    <td valign="bottom" align="center">
	<div align="right">
		<span style="font-size:8px">Authorize: <%=Accredit%></span>
		<img src="../images/blank.gif" height=1 width=30 align=absmiddle>
		<a href="../register">[注册]</a>
		<img src="../images/blank.gif" height=1 width=5 align=absmiddle>
		<a href="javascript:;" onclick="window.open('../admin/default.asp','','width=300px,height=100')">[管理]</a>
		<img src="../images/blank.gif" height=1 width=5 align=absmiddle>
		<a href="../supervise/login.asp?logout">[退出]</a>
		<img src="../images/blank.gif" height=1 width=5 align=absmiddle>
	</div>
	<div id="divHeadAd"><br><br></div>
	<table border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td class="MenuItem">&nbsp;</td>
      <td id="MainMenu" class="MenuItemOnSel" onMouseOver="doOver(this,0)"><a href="index.asp">统计概况</a></td>
      <td id="MainMenu" class="MenuItem" onMouseOver="doOver(this,1)"><a href="Hour.asp">分时统计</a></td>
      <td id="MainMenu" class="MenuItem" onMouseOver="doOver(this,2)"><a href="Page.asp">内容统计</a></td>
      <td id="MainMenu" class="MenuItem" onMouseOver="doOver(this,3)"><a href="Referer.asp">来源统计</a></td>
      <td id="MainMenu" class="MenuItem" onMouseOver="doOver(this,4)"><a href="Link.asp">客户端统计</a></td>
      <td id="MainMenu" class="MenuItem" onMouseOver="doOver(this,5)">联机帮助</td>
      <td id="MainMenu" class="MenuItem" onMouseOver="doOver(this,6)">定制样式</td>
      <td class="MenuItem" style="width:50px">&nbsp;</td>
    </tr>
    <tr>
      <td colspan="8" height="24" align="left" valign="bottom">
				<div id="menuContent" style="display:">
					<img src="../images/blank.gif" width=56 height=1>
					<span class="ML">
					&nbsp;&nbsp;
					<a href="index.asp">综合统计</a>
					| <a href="Info.asp">修改资料</a>
					| <a href="Help.asp">统计代码</a>
					&nbsp;&nbsp;
					</span>
				</div>
				<div id="menuContent" style="display:none">
					<img src="../images/blank.gif" width=116 height=1>
					<span class="ML">
					&nbsp;&nbsp;<a href="Hour.asp">日报一览</a>
					| <a href="Week.asp">周报一览</a>
					| <a href="Day.asp">月报一览</a>
					| <a href="Month.asp">年报一览</a>
					&nbsp;&nbsp;
					</span>
				</div>
				<div id="menuContent" style="display:none">
					<img src="../images/blank.gif" width=176 height=1>
					<span class="ML">
					&nbsp;&nbsp;<a href="Page.asp">页面统计</a> | <a href="Keyword.asp">关键字统计</a>
					&nbsp;&nbsp;
					</span>
				</div>
				<div id="menuContent" style="display:none;text-align:right">
					<span class="MR">
					&nbsp;&nbsp;&nbsp;&nbsp;
					<a href="Referer.asp">来源统计</a> | <a href="Region.asp">地区统计</a>
					&nbsp;&nbsp;
					</span>
					<img src="../images/blank.gif" width=179 height=1>
				</div>
				<div id="menuContent" style="display:none;text-align:right">
					<span class="MR">
					&nbsp;&nbsp;&nbsp;&nbsp;
					<a href="Link.asp">上网方式统计</a>
					| <a href="ClientSoft.asp">软件统计</a>
					| <a href="ClientColor.asp">屏幕统计</a>
					| <a href="ClientZone.asp">区域统计</a>
					&nbsp;&nbsp;
					</span>
					<img src="../images/blank.gif" width=119 height=1>
				</div>
				<div id="menuContent" style="display:none;text-align:right">
					<span class="MR">
					&nbsp;&nbsp;&nbsp;&nbsp;
					<a href="http://www.ccopus.com/cc6/forum.asp?t=cc6" target="_blank">在线支持</a>
					| <a href="../utilities/update.asp" target="_blank">在线更新</a>
					&nbsp;&nbsp;
					</span>
					<img src="../images/blank.gif" width=59 height=1>
				</div>
				<div id="menuContent" style="display:none;text-align:right">
					<span class="MR">
					&nbsp;&nbsp;&nbsp;&nbsp;
					<a href="javascript:changeStyle(3)">春</a>
					| <a href="javascript:changeStyle(4)">夏</a>
					| <a href="javascript:changeStyle(2)">秋</a>
					| <a href="javascript:changeStyle(1)">冬</a>
					&nbsp;&nbsp;
					</span>
				</div>
			</td>
			<td>&nbsp;</td>
      </tr>
  </table>	</td>
  </tr>
</table>
<br>
