document.onselectstart=document.ondragstart=document.oncontextmenu=function(){return false;};
document.onkeydown=function(){event.keyCode=0;return false;}
window.resizeTo(640,480);