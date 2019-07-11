<%
dim conn,sqlstr,DBPath
dim bodymax,name,url,gbook_url,password,count,today_count,today_time,setup_ip,setup_time

DBPath = Server.MapPath("data/#dnyes##jj.asp")

Connstr = "DRIVER={Microsoft Access Driver (*.mdb)};DBQ="& DBPath
Set conn = Server.CreateObject("ADODB.Connection")
Conn.open Connstr
sql="select * from admin"
set Rs = Server.CreateObject("ADODB.Recordset")
Rs.open sql,conn,1,1

gbook_name = Rs("gbook_name")
passwords = Rs("password")
name = Rs("name")
url = Rs("url")
bodymax = Rs("bodymax")
pagesize = Rs("pagesize")
today_count = Rs("today_count")
today_time = Rs("today_time")
setup_ip = Rs("setup_ip")
setup_time = Rs("setup_time")
Rs.close

set Rs = nothing

sub error(message)
%>
<script>alert('<%=message%>');history.back();</script><script>window.close();</script>
<%
response.end
end sub

sub htmlend()
%>

<script language=JavaScript>
<!-- // BannerAD

   var bannerAD=new Array();
   var bannerADlink=new Array();
   var adNum=0;

   bannerAD[0]="http://www.sczz.com/book/images/book.gif";
   bannerADlink[0]="http://www.sczz.com/book/register.asp";
   bannerAD[1]="http://www.sczz.com/book/images/banner.gif";
   bannerADlink[1]="http://www.sczz.com/";
   bannerAD[2]="http://www.sczz.com/book/images/banner1.gif";
   bannerADlink[2]="http://www.sczz.com/";
   
   var preloadedimages=new Array();
   for (i=1;i<bannerAD.length;i++){
      preloadedimages[i]=new Image();
      preloadedimages[i].src=bannerAD[i];
   }

function setTransition(){
   if (document.all){
      bannerADrotator.filters.revealTrans.Transition=Math.floor(Math.random()*23);
      bannerADrotator.filters.revealTrans.apply();
   }
}

function playTransition(){
   if (document.all)
      bannerADrotator.filters.revealTrans.play()
}

function nextAd(){
   if(adNum<bannerAD.length-1)adNum++ ;
      else adNum=0;
   setTransition();
   document.images.bannerADrotator.src=bannerAD[adNum];
   playTransition();
   theTimer=setTimeout("nextAd()", 12000);
}

function jump2url(){
   jumpUrl=bannerADlink[adNum];
   jumpTarget='_blank';
   if (jumpUrl != ''){
      if (jumpTarget != '')window.open(jumpUrl,jumpTarget);
      else location.href=jumpUrl;
   }
}
function displayStatusMsg() { 
   status=bannerADlink[adNum];
   document.returnValue = true;
}

//-->
</script>
<div align="center">
  <center>
<table cellpadding="3" cellspacing="0" style="border-collapse: collapse" bordercolor="#4F4F4F" width="500">
  <tr>
 
    <table cellspacing=0 cellpadding=0 width=520 border=0 align="center">
      <tr>
          <td width="25"><img height=35 src="images/f.gif" width=25 border=0></td> 
          <td width=464 background=images/f1.gif><img height=1  
            src="images/space.gif" width=1 border=0></td> 
          <td width="25"><img height=35 src="images/f2.gif" width=25  
        border=0></td> 
      </tr>
      <tr>
          <td background=images/l.gif width="25"><img height=1  
            src="images/space.gif" width=1 border=0></td> 
        <td width=464 bgcolor="#F8FCF8" height="70">  
          <p align="center"><A onmouseover="displayStatusMsg();return document.returnValue" 
      href="javascript:jump2url()"><IMG 
      style="FILTER: revealTrans(duration=1,transition=20)" height=60 
      src="http://www.sczz.com/images/banner.gif" width=468 border=0 
      name=bannerADrotator>
</A> 
<SCRIPT language=JavaScript>nextAd()</SCRIPT>        
        </td>         
      </center>  
  <center> 
          <td background=images/r.gif width="25"><img height=1          
            src="images/space.gif" width=1 border=0></td>         
      </tr> 
      <tr>        
          <td width="25"><img height=11 src="images/bl.gif" width=25 border=0></td>         
          <td width=464 background=images/bm.gif><img height=1          
            src="images/space.gif" width=1 border=0></td>         
          <td width="25"><img height=11 src="images/br.gif" width=25          
        border=0></td>         
      </tr>        
    </table>                
  </center>                
  </tr>   
</table>   
</div>   
   
<% 
end sub 
%>