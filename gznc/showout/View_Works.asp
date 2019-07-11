<HTML>
<HEAD>
<TITLE>新筑-[NewCreation]-广州新筑建筑设计咨询有限公司</TITLE>
<style type="text/css">
<!--
.opacity {
FILTER: alpha(opacity=100)
}
-->
</style>
<SCRIPT>
<!--



function movstar(a){
movx=setInterval("mov("+a+")",10)
}
function movover(){
clearInterval(movx)
}
function mov(a){
scrollx=new_date.document.body.scrollLeft
scrolly=new_date.document.body.scrollTop
scrolly=scrolly+a
new_date.window.scroll(scrollx,scrolly)
}
function o_down(theobject){
object=theobject
while(object.filters.alpha.opacity>60){
object.filters.alpha.opacity+=-10}
}
function o_up(theobject){
object=theobject
while(object.filters.alpha.opacity<100){
object.filters.alpha.opacity+=10}
}
function wback(){
if(new_date.history.length==0){window.history.back()}
else{new_date.history.back()}
}

function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}
//-->
</SCRIPT>
<SCRIPT language=JavaScript>

imgPath = new Array;
SiClickGoTo = new Array;
if (document.images)
	{
	i0 = new Image;
	i0.src = '../img/nc_works_001.gif';
	SiClickGoTo[0] = "#";
	imgPath[0] = i0.src;
	i1 = new Image;
	i1.src = '../img/nc_works_002.gif';
	SiClickGoTo[1] = "#";
	imgPath[1] = i1.src;
	i2 = new Image;
	i2.src = '../img/nc_works_003.gif';
	SiClickGoTo[2] = "#";
	imgPath[2] = i2.src;
	i3 = new Image;
	i3.src = '../img/nc_works_004.gif';
	SiClickGoTo[3] = "#"
	}
a = 0;
function ejs_img_fx(img)
	{
	if(img && img.filters && img.filters[0])
		{
		img.filters[0].apply();
		img.filters[0].play();
		}
	}

function StartAnim()
	{
	if (document.images)
		{
		document.write('<A HREF="#" onClick="ImgDest();return(false)"><IMG SRC="../img/nc_works_001.gif" BORDER=0 ALT=优良品质-精心挑选 NAME=defil style="filter:progid:DXImageTransform.Microsoft.Pixelate(MaxSquare=100,Duration=1)"></A>');
		defilimg()
		}
	else
		{
		document.write('<A HREF="#"><IMG SRC="../img/nc_works_001.gif" BORDER=0></A>')
		}
	}
function ImgDest()
	{
	document.location.href = SiClickGoTo[a-1];
	}
function defilimg()
	{
	if (a == 3)
		{
		a = 0;
		}
	if (document.images)
		{
		ejs_img_fx(document.defil)
		document.defil.src = imgPath[a];
		tempo3 = setTimeout("defilimg()",8000);
		a++;
		}
	}
</SCRIPT>

<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<META content="Web stategy, design and integration." name=Description>
<LINK href="CSS.css" type=text/css rel=stylesheet>
<META content="MSHTML 6.00.2800.1106" name=GENERATOR></HEAD>
<BODY bgcolor="#4C5155" leftMargin=0 topMargin=0 marginwidth="0" 
marginheight="0">
<map name="Map"><area shape="rect" coords="834,5,977,21" href="Contact_F.asp">
  <area shape="rect" coords="661,6,804,22" href="Share_F.asp">
  <area shape="rect" coords="488,7,631,23" href="InTouch_F.asp">
  <area shape="rect" coords="317,7,460,23" href="Works_F.asp">
  <area shape="rect" coords="143,7,286,23" href="TheFirm_F.asp">
</map>
<!--#include file="dbm_conn.asp" -->
<%
subs_id=request("subs_id")
Set rs=Server.CreateObject("ADODB.RecordSet")
big_sfile=request("big_sfile")

if subs_id="" then
sql="select * from dbm_subs order by subs_id"
rs.open sql,conn,1,1 
if rs.eof or rs.bof then
rs.close 
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' 没有合要求的作品 ! ');</SCRIPT>")
response.End()
else
subs_id=rs("subs_id")
rs.close
end if
end if

sql="select * from dbm_subs where subs_id="&subs_id
rs.open sql,conn,1,1 
if rs.eof or rs.bof then 
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' 没有合要求的作品 ! ');</SCRIPT>")
response.End()
end if 
if big_sfile="" then big_sfile=rs("sfile1")
%>
<table width="990" height="100%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td height="108"><table width="990" height="108" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr> 
          <td width="150" background="../img/logo_bg.gif"><img src="../img/logo_small.gif" width="101" height="108"></td>
          <td align="right" background="../img/logo_bg.gif"><table width="800" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td width="134" align="right"><SCRIPT language=JavaScript>StartAnim()</SCRIPT></td>
              </tr>
            </table></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td height="10"><table width="990" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr> 
          <td width="990" height="10"><img src="../img/1.gif" width="1" height="1"></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td height="2"> <table width="990" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr> 
          <td width="990" background="../img/bg_orange.gif"><img src="../img/bg_orange.gif" width="2" height="2"></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td valign="top"><table width="990" height="100%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td height="10"><table width="473" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td height="10" background="../img/bg_orange.gif"><img src="../img/bg_orange.gif" width="2" height="2"></td>
              </tr>
            </table></td>
        </tr>
        <tr> 
          <td valign="top"><table width="990" height="100%" border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td width="130" valign="top"><table width="100%" height="100%" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td height="130" valign="top"><table width="75%" border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                      <td height="5"><img src="../img/1.gif" width="1" height="1"></td>
                    </tr>
                  </table>
                  <img src="../photo/<%=rs("sfile1")%>" width="130" height="130"></td>
              </tr>
              <tr> 
                <td height="8"><table width="75%" border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                      <td height="8"><img src="../1.gif" width="1" height="1"></td>
                    </tr>
                  </table></td>
              </tr>
              <tr> 
                <td valign="top" background="../img/5x5.gif"> <table width="126" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td width="145">&nbsp;</td>
                    </tr>
                    <tr> 
                      <td height="25" valign="top"><b><a href="#"><%=rs("sname")%></a></b></td>
                    </tr>
                    <tr> 
                      <td valign="top"><font color=#999999> 
<%
if len(rs("memo"))>100 then
memo=left(rs("memo"),95)&"..."
else
memo=rs("memo")
end if
memo=replace(memo,chr(13),"<br><br>")
%>						
						
                        <%=memo%>
                        &nbsp;</font></td>
                    </tr>
                  </table></td>
              </tr>
            </table></td>
                <td width="176" valign="top"><table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td height="125">&nbsp;</td>
              </tr>
              <tr> 
                <td> 
                  <% 
							 if rs("sfile1")<>" " then response.Write("<a href='View_Works.asp?subs_id="&subs_id&"&big_sfile="&rs("sfile1")&"'>1</a>&nbsp;&nbsp;")
							 if rs("sfile2")<>" " then response.Write("<a href='View_Works.asp?subs_id="&subs_id&"&big_sfile="&rs("sfile2")&"'>2</a>&nbsp;&nbsp;")
							 if rs("sfile3")<>" " then response.Write("<a href='View_Works.asp?subs_id="&subs_id&"&big_sfile="&rs("sfile3")&"'>3</a>&nbsp;&nbsp;")
							 if rs("sfile4")<>" " then response.Write("<a href='View_Works.asp?subs_id="&subs_id&"&big_sfile="&rs("sfile4")&"'>4</a>&nbsp;&nbsp;")
							 if rs("sfile5")<>" " then response.Write("<a href='View_Works.asp?subs_id="&subs_id&"&big_sfile="&rs("sfile5")&"'>5</a>&nbsp;&nbsp;")
							 if rs("sfile6")<>" " then response.Write("<a href='View_Works.asp?subs_id="&subs_id&"&big_sfile="&rs("sfile6")&"'>6</a>&nbsp;&nbsp;")
							 if rs("sfile7")<>" " then response.Write("<a href='View_Works.asp?subs_id="&subs_id&"&big_sfile="&rs("sfile7")&"'>7</a>&nbsp;&nbsp;")
							 if rs("sfile8")<>" " then response.Write("<a href='View_Works.asp?subs_id="&subs_id&"&big_sfile="&rs("sfile8")&"'>8</a>")
						  %>
                </td>
              </tr>
            </table></td>
                <td width="167" valign="top"><table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td height="25" align="right" valign="bottom"><img src="../img/zpml.gif" width="167" height="17"></td>
              </tr>
              <tr> 
                <td valign="top" background="../img/5x5.gif"><table width="149" height="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td width="149"><iframe allowtransparency  
      border=0 frameborder=0 framespacing=0 height=100% marginheight=0 
      marginwidth=0 name=new_date noResize scrolling=no 
      src="Projects.asp" width=100% vspale="0"> </iframe></td>
                    </tr>
                  </table></td>
              </tr>
            </table></td>
                <td width="32" valign="bottom"><table width="27" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td>&nbsp;</td>
              </tr>
              <tr> 
                <td height="47"><a href="javascript:void(0)" onFocus="if(this.blur)this.blur()"><img src="../img/up.gif" width="32" height="25" class=opacity  onMouseDown=movover();movstar(-10) onMouseOut=movover();o_up(this) onMouseOver=movstar(-3);o_down(this)  border=0></a></td>
              </tr>
              <tr> 
                <td><a href="javascript:void(0)" onFocus="if(this.blur)this.blur()"><img src="../img/down.gif" width="32" height="25" border="0" class=opacity onMouseDown=movover();movstar(10) onMouseOut=movover();o_up(this) onMouseOver=movstar(3);o_down(this) ></a></td>
              </tr>
            </table></td>
                <td align="center"><img src="../photo/<%=big_sfile%>" width="361" height="361"></td>
              </tr>
            </table></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td height="8"><table width="75%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td height="8"><img src="../1.gif" width="1" height="1"></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td height="33"><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="990" height="33">
        <param name="movie" value="../img/index/works.swf">
        <param name="quality" value="high">
        <embed src="../img/index/works.swf" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" width="990" height="33"></embed></object></td>
  </tr>
  <tr> 
    <td height="45" valign="bottom"> <div id=demo style=overflow:hidden;height:45;width:990;background:#214984;color:#ffffff> 
        <table align=left cellspacing="0" cellpadding=0 cellspace=0 border=0>
          <tr> 
            <td id=demo1 valign=top> <img src="../img/house_001.gif" width="990" height="45"> 
              <img src="../img/house_001.gif" width="990" height="45"> </td>
            <td id=demo2 valign=top></td>
          </tr>
        </table>
      </div>
      <script>
  var speed=30
  demo2.innerHTML=demo1.innerHTML
  function Marquee(){
  if(demo2.offsetWidth-demo.scrollLeft<=0)
  demo.scrollLeft-=demo1.offsetWidth
  else{
  demo.scrollLeft++
  }
  }
  var MyMar=setInterval(Marquee,speed)
  demo.onmouseover=function() {clearInterval(MyMar)}
  demo.onmouseout=function() {MyMar=setInterval(Marquee,speed)}
  </script> </td>
  </tr>
</table>
<%
rs.close
set rs=nothing
%>
</BODY></HTML>
