var imgheight
var imgleft
document.ns = navigator.appName == "Netscape"
window.screen.width>800 ? imgheight=100:imgheight=100
window.screen.width>800 ? imgleft=15:imgleft=120
function myload()
{
if (navigator.appName == "Netscape")
{document.myleft.pageY=pageYOffset+window.innerHeight-imgheight;
document.myleft.pageX=imgleft;
leftmove();
}
else
{
myleft.style.top=document.body.scrollTop+document.body.offsetHeight-imgheight;
myleft.style.left=imgleft;
leftmove();
}
}
function leftmove()
 {
 if(document.ns)
 {
 document.myleft.top=pageYOffset+window.innerHeight-imgheight
 document.myleft.left=imgleft;
 setTimeout("leftmove();",80)
 }
 else
 {
 myleft.style.top=document.body.scrollTop+document.body.offsetHeight-imgheight;
 myleft.style.left=imgleft;
 setTimeout("leftmove();",80)
 }
 }

function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true)

if (navigator.appName == "Netscape")
{
document.write("<layer id=myleft top=300 width=80 height=88><a href=http://ad.cn.doubleclick.net/clk;4026082;6389205;g?http://ad4.sina.com.cn/button/guoxinguanghuan/0315.html target=_blank><img src=http://ad4.sina.com.cn/button/guoxinguanghuan/head.gif border='0'></a></layer>");
myload()}
else
{
document.write("<div id=myleft style='position: absolute;width:80;top:300;left:5;visibility: visible;z-index: 1'><a href=http://ad.cn.doubleclick.net/clk;4026082;6389205;g?http://ad4.sina.com.cn/button/guoxinguanghuan/0315.html target=_blank><img src=http://ad4.sina.com.cn/button/guoxinguanghuan/head.gif border='0'></a></div>");
myload()
}
