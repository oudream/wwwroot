window.ClearEvent=function(){event.cancelBubble=false;var sSrcTagName=event.srcElement.tagName.toLowerCase();
return (sSrcTagName=="textarea" || sSrcTagName=="input" || sSrcTagName=="select");}
window.ClearKey=function(){event.cancelBubble=false;var iKeyCode=event.keyCode;return !(iKeyCode==78 && event.ctrlKey);}
with (window.document){oncontextmenu=onselectstart=ondragstart=window.ClearEvent;onkeydown=window.ClearKey;}