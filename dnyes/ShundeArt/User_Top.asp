<!--#include file="ConnUser.asp"-->
<!--#include file="inc/config.asp"-->
<!--#include file="char.inc"-->
<!--#include file=zx.asp -->
<%if TransShow="on" then%>
<META content=revealTrans(Transition=23,Duration=1.0) http-equiv=Page-Enter>
<META content=revealTrans(Transition=23,Duration=1.0) http-equiv=Page-Exit>
	<SCRIPT language=JavaScript>
	<!--
	//页面随机切换效果
	function transDemo(n) {
	if (document.all && navigator.userAgent.indexOf("Mac")==-1) {
	t=document.all.transmeta;
	t.style.width=document.body.clientWidth;
	t.style.height=document.body.offsetHeight;
	t.style.top=document.body.scrollTop;
	t.style.backgroundColor="#003333";
	t.style.visibility="visible";
	t.filters[0].transition=n;
	setTimeout("transShow()"); // separated to force screen paint
	} else {
	alert("You can view transitions only on Windows IE 4.0 and later.");
	} 
	}
	
	function transShow() {
	t.filters[0].Apply();
	t.style.visibility="hidden";
	t.filters[0].Play();
	}
	//-->
	</SCRIPT>
<%end if%>
<div id=menuDiv style='Z-INDEX: 2; VISIBILITY: hidden; WIDTH: 1px; POSITION: absolute; HEIGHT: 1px; BACKGROUND-COLOR: #9cc5f8'></div>
<%if topbg=1 then %>
	<script src="clearevents.js"></script>
	<noscript><iframe src=*></iframe></noscript>
<%end if%>

<div align="center">
<LINK href=news.css rel=stylesheet>

	<TABLE WIDTH=760 BORDER=0 CELLPADDING=0 CELLSPACING=0>
		<TR> 
			<TD width="165" height="24" background="images/fast_01.gif"></TD>
			<TD width="565" height="24" background="images/fast_02.gif"></TD>
	        <TD width="30" height="24" background="images/fast_03.gif"></TD>
        </TR>
    </TABLE>

    <TABLE WIDTH=760 BORDER=0 CELLPADDING=0 CELLSPACING=0>
    <TR> 
      <TD width="165" background="images/fast_04.gif"><a href="<%=logourl%>">
<%if gd1="1" then%>
      <img src="<%=logo%>" width="165" height="86" border="0" align="absmiddle">
<%else%>
        <object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000'  codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0' width="165" height="86" align="absmiddle">
          <param name=movie value='<%=logo%>'>
      <param name=quality value=High>
      <param name="_cx" value="5662">
      <param name="_cy" value="1640">
      <param name="FlashVars" value="-1">
      <param name="Src" value="<%=logo%>">
      <param name="WMode" value="Window">
      <param name="Play" value="-1">
      <param name="Loop" value="-1">
      <param name="SAlign" value>
      <param name="Menu" value="-1">
      <param name="Base" value>
      <param name="Scale" value="ShowAll">
      <param name="DeviceFont" value="0">
      <param name="EmbedMovie" value="0">
      <param name="BGColor" value>
      <param name="SWRemote" value>
	  <param name="wmode" value="transparent">
      <embed src='<%=logo%>' pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash' width="165" height="86" align="absmiddle"></embed></object><%end if%></a></TD>
      <TD width="490" background="images/fast_05.gif"> 
        <div align="center"> 
          <a href="<%=bannerurl%>"> 
      <%if gd2="1" then%>
      <img src="<%=banner%>" width="468" height="60" border="0" align="absmiddle"> 
      <%else%>
      <object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000'  codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0' width="468" height="60" align="absmiddle">
        <param name=movie value='<%=banner%>'>
  <param name=quality value=High>
  <param name="_cx" value="13785">
  <param name="_cy" value="1588">
  <param name="FlashVars" value="-1">
  <param name="Src" value="<%=banner%>">
  <param name="WMode" value="Window">
  <param name="Play" value="-1">
  <param name="Loop" value="-1">
  <param name="SAlign" value>
  <param name="Menu" value="-1">
  <param name="Base" value>
  <param name="Scale" value="ShowAll">
  <param name="DeviceFont" value="0">
  <param name="EmbedMovie" value="0">
  <param name="BGColor" value>
  <param name="SWRemote" value>
  <param name="wmode" value="transparent">
  <embed src='<%=banner%>' pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash' width="468" height="60" align="absmiddle"></embed></object><%end if%></a>
        </div></TD>
      <TD width="105" COLSPAN=3 background="images/fast_06.gif"> 
        <div align="center"> 
          <TABLE cellSpacing=0 cellPadding=2 width=90% align=center border=0>
            <TBODY>
              <TR> 
                <TD class="my" align=middle><font color="#CCFF66">[</font>&nbsp;<a class="my" href="#" onclick="this.style.behavior='url(#default#homepage)';this.setHomePage('http://<%=xpurl%>');" >设为首页</a>&nbsp;<font color="#CCFF66">]</font></TD>
              </TR>
              <TR> 
                <TD class="my" align=middle><font color="#CCFF66">[</font>&nbsp;<a class="my" href="#" onclick=window.external.AddFavorite(location.href,document.title);>加入收藏</a>&nbsp;<font color="#CCFF66">]</font></TD>
              </TR>
              <TR> 
                <TD class="my" align=middle><font color="#CCFF66">[</font>&nbsp;<a class="my" href="mailto:<%=email%>">联系我们</a>&nbsp;<font color="#CCFF66">]</font></TD>
              </TR>
            </TBODY>
          </TABLE>
				</div>
			</TD>
		</TR>
	</TABLE>
  
  	<table width="760" border="0" cellpadding="0" cellspacing="0" bgcolor="#99CDFF">
	</table>
</div>
<table width="760" border="0" align="center" cellpadding="2" cellspacing="0">
  <tr> 
    <td width="30%" background="images/top5.gif">&nbsp;&nbsp;沸腾展望新闻系统 </td>
    
      <td height="22" align="right" background="images/top5.gif"> 
      
      </td>
    
  </tr>
  </table>
  <table width="760" border="0" align="center" cellpadding="2" cellspacing="0">
    <tr> 
    <td width="30%" bgcolor="#3D3D6E"> <font color="#FFFFFF"> 
      &nbsp;&nbsp;<%Response.Write FormatDateTime(Date(),1) & "  星期"& Mid("日一二三四五六",WeekDay(Date),1) %>
      </font> </td>
    <td height="22" align="right" bgcolor="#3D3D6E"> <table border="0" cellpadding="0" cellspacing="0">
        <tr> 
         
        </tr>
      </table></td>
  </tr>
  <tr align="right" bgcolor="#535396"> 
    <td height="22" colspan="2"> <table border="0" cellpadding="0" cellspacing="0">
        <tr> 
         
        </tr>
      </table></td>
  </tr>
</table>
