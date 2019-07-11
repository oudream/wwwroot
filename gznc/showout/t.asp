<script language="" type="text/JavaScript" class="pt9">
var marqueeContent=new Array();
marqueeContent[10]='<a href="http://#"><img src="../img/Pic_A_001.gif" border="0"></a>&nbsp;'; 
marqueeContent[9]='<a href="http://#"><img src="../img/Pic_A_002.gif" border="0"></a>&nbsp;'; 
marqueeContent[8]='<a href="http://#"><img src="../img/Pic_A_003.gif" border="0"></a>&nbsp;'; 
marqueeContent[7]='<a href="http://#"><img src="../img/Pic_A_004.gif" border="0"></a>&nbsp;'; 
marqueeContent[6]='<a href="http://#"><img src="../img/Pic_A_005.gif" border="0"></a>&nbsp;'; 
marqueeContent[5]='<a href="http://#"><img src="../img/Pic_A_006.gif" border="0"></a>&nbsp;'; 
marqueeContent[4]='<a href="http://#"><img src="../img/Pic_A_007.gif" border="0"></a>&nbsp;'; 
marqueeContent[3]='<a href="http://#"><img src="../img/Pic_A_008.gif" border="0"></a>&nbsp;'; 
marqueeContent[2]='<a href="http://#"><img src="../img/Pic_A_009.gif" border="0"></a>&nbsp;'; 
marqueeContent[1]='<a href="http://#"><img src="../img/Pic_A_010.gif" border="0"></a>&nbsp;'; 
marqueeContent[0]='<a href="http://#"><img src="../img/Pic_A_011.gif" border="0"></a>&nbsp;'; 
var marqueeInterval=new Array();  
var marqueeId=0;
var marqueeDelay=5000;
var marqueeHeight=112;

Array.prototype.random=function() {
	var a=this;
	var l=a.length;
	for(var i=0;i<l;i++) {
		var r=Math.floor(Math.random()*(l-i));
		a=a.slice(0,r).concat(a.slice(r+1)).concat(a[r]);
		}
	return a;
	}
function initMarquee() {
	marqueeContent=marqueeContent.random();
	var str='';
	for(var i=0;i<Math.min(1,marqueeContent.length);i++) str+=(i>0?'':'')+marqueeContent[i];
	document.write('<span id=marqueeBox style="overflow:hidden;height:'+marqueeHeight+'px" onmouseover="clearInterval(marqueeInterval[0])" onmouseout="marqueeInterval[0]=setInterval(\'startMarquee()\',marqueeDelay)"><span>'+str+'</span></span>');
	marqueeId+=2;
	if(marqueeContent.length>1)marqueeInterval[0]=setInterval("startMarquee()",marqueeDelay);
	}
function startMarquee() {
	var str='';
	for(var i=0;(i<1)&&(marqueeId+i<marqueeContent.length);i++) {
		str+=(i>0?'':'')+marqueeContent[marqueeId+i];
		}
	marqueeId+=1;
	if(marqueeId>marqueeContent.length)marqueeId=0;

	if(marqueeBox.childNodes.length==1) {
		var nextLine=document.createElement('DIV');
		nextLine.innerHTML=str;
		marqueeBox.appendChild(nextLine);
		}
	else {
		marqueeBox.childNodes[0].innerHTML=str;
		marqueeBox.appendChild(marqueeBox.childNodes[0]);
		marqueeBox.scrollTop=0;
		}
	clearInterval(marqueeInterval[1]);
	marqueeInterval[1]=setInterval("scrollMarquee()",20);
	}
function scrollMarquee() {
	marqueeBox.scrollTop++;
	if(marqueeBox.scrollTop%marqueeHeight==(marqueeHeight-1)){
		clearInterval(marqueeInterval[1]);
		}
	}
initMarquee();
</script> 
