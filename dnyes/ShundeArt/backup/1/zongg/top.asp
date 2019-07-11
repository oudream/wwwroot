<!-- 
'*****************************  本版发行日期：2004-2-20  ——  版权说明  ——  请保留此信息 ****************************************
'***** 程序名称:忠网广告管理系统                               ****                     English Name:ZonGG                      *****
'***** 当前版本:ZonGG Version 1.1.0.0             ******************************        Powered By:ZonGG Version 1.1.0.0        *****
'***** 开发制作:忠网.黑子 2003-2004               ***          ****          ***        Do By:Zon.Hizi 2003-2004                *****
'***** 版权所有:忠网.黑子 2003-2004               ******************************        Copyright:Zon.Hizi 2003-2004            *****
'***** 忠网域名:http://www.zon.cn                 *****        ****         ****        URL:http://www.zon.cn                   *****
'***** 网络实名:忠网                              ***   **********************          忠网                                    *****
'***** 程序下载地址:http://zon.cn/down                                    ******        DownLoad:http://zon.cn/down             *****
'***** 技术支持信箱:hizi@zon.cn                                                         E-mail:hizi@zon.cn                      *****
'***** 3721中文邮:黑子@忠网                                                             C-mail:黑子@忠网                        *****
'************************************************************************************************************************************
-->
<html><head>
<META content="忠网广告管理系统,ZonGG Version 1.1.0.0,hizi@zon.cn,http://www.zon.cn,http://zon.cn,http://zon.cn/down,测试：http://zon.cn/cs/gg,广告,  ,忠网.黑子 2003年-2004年 " name=KEYWORDS>
<meta http-equiv=Content-Language content=zh-cn>
<meta http-equiv=Content-Type content=text/html; charset=gb2312>
<meta http-equiv=Pragma content=no-cache>
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link href=style.css rel=STYLESHEET type=text/css>

<title>&#24544;&#32593;.&#24191;&#21578;&#31649;&#29702;&#31995;&#32479;</title>
<base target=_top>
</head><body marginwidth=0 marginheight=0>
   <div align="center">
      <center>
<table style="BORDER-COLLAPSE: collapse" cellspacing=0 
cellpadding=0 width="789" align=center border=0>
  <tbody> 
  <tr> 
    <td width=1% height=65 background="xppic/top_left.gif"> 
      <table border="0" cellpadding="0" cellspacing="0" width="20" class=link>
        <!-- fwtable fwsrc="Untitled" fwbase="top_left.gif" fwstyle="Dreamweaver" fwdocid = "742308039" fwnested="0" -->
        <tr> 
          <td width="5" height="65"></td>
          <td width="10" height="65" background="xppic/top_mid.jpg"> 
            　</td>
        </tr>
        </table>
    </td>
    <td background="xppic/top_mid.jpg" width="99%" align="center"> <br><font size="6" color="#FFFFFF" face="华文行楷">G</font>&nbsp; &nbsp;
    

   
    <%if session("masterlogin")<>"superadvertadmin" then%><font size="3" color="#FFFFFF" face="&#40657;&#20307;"> </font>
    <font size="3" color="#FFFFFF" face="&#40657;&#20307;"> &#24544;&#32593;.&#24191;&#21578;&#31649;&#29702;&#31995;&#32479;</font>  
<%else%>
      <a href="index.asp"><font color="#FFFFFF">&#22238;&#39318;&#39029;</font></a><font color="#FFFFFF">&nbsp;
    <a href="addadm.asp"><font color="#FFFFFF">&#31649;&#29702;&#21592;</font></a>&nbsp; </font> <a href=addadw.asp ><font color="#FFFFFF">
    &#24191;&#21578;&#20301;</font></a><font color="#FFFFFF">&nbsp;
        </font>
        <a href=addads.asp ><font color="#FFFFFF">&#28155;&#21152;&#24191;&#21578;</font></a><font color="#FFFFFF">&nbsp; </font> <a href=list.asp ><font color="#FFFFFF">&#27491;&#24120;&#24191;&#21578;</font></a><font color="#FFFFFF">&nbsp;
    </font> <a href="list.asp?type=img"><font color="#FFFFFF">&#22270;&#29255;&#24191;&#21578;</font></a><font color="#FFFFFF">&nbsp;
    </font> <a href=list.asp?type=txt ><font color="#FFFFFF">&#25991;&#26412;&#24191;&#21578;</font></a><font color="#FFFFFF">&nbsp;
    </font> <a href="list.asp?type=click"><font color="#FFFFFF">&#28857;&#20987;&#25490;&#34892;</font></a><font color="#FFFFFF">&nbsp;
    </font> <a href="list.asp?type=show"><font color="#FFFFFF">&#26174;&#31034;&#25490;&#34892;</font></a><font color="#FFFFFF">&nbsp;
    </font> <a href=list.asp?type=close ><font color="#FFFFFF">&#26242;&#20572;&#24191;&#21578;</font></a><font color="#FFFFFF">&nbsp; </font> <a href=list.asp?type=lose ><font color="#FFFFFF">&#22833;&#25928;&#24191;&#21578;</font></a><font color="#FFFFFF">&nbsp;
    </font> <a href=logout.asp ><font color="#FFFFFF">&#36864;&#20986;</font></a>
        <%end if%>
        </td>
    <td width=1%>
    <img height=65 
      src="xppic/top_right.gif" 
      width=36></td>
  </tr>
  </tbody>
</table>

<div align="center">
  <center>


<TABLE style="BORDER-COLLAPSE: collapse" cellSpacing=0 
cellPadding=0 width="789" border=0 bgcolor="#FFFFFF">
  <TBODY> 
  <TR> <TD bgColor=#3373ce width="2">
</TD>


    <TD> 
<table width="98%" border="0" cellspacing="0" cellpadding="5" align="center">
<tr>
<td><%call  Iflogin()%>