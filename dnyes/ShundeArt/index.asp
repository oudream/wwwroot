<%@ Language=VBScript%>
<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="function.asp"-->
<!--#include file="char.inc"-->

<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="keywords" content="顺德国际艺术交流中心http://www.shundeart.com/">
<title>顺德国际艺术交流中心_首页</title>
<LINK href=news.css rel=stylesheet>
<script language="JavaScript">
<!--
if (self != top) top.location.href = window.location.href
//-->
</script>
<script language=JavaScript>
<!--
//
var version = "other"
browserName = navigator.appName;   
browserVer = parseInt(navigator.appVersion);
if (browserName == "Netscape" && browserVer >= 3) version = "n3";
else if (browserName == "Netscape" && browserVer < 3) version = "n2";
else if (browserName == "Microsoft Internet Explorer" && browserVer >= 4) version = "e4";
else if (browserName == "Microsoft Internet Explorer" && browserVer < 4) version = "e3";
function marquee1()
{
	if (version == "e4" | version == "n3")
	{
		document.write("<marquee style='BOTTOM: 0px; FONT-WEIGHT: 100px; HEIGHT:160px;  TEXT-ALIGN: left; TOP: 0px' id='news' scrollamount='1' scrolldelay='10' behavior='loop' direction='up' border='0' onmouseover='this.stop()' onmouseout='this.start()'>")
	}
}

function marquee2()
{
	if (version == "e4" | version == "n3")
	{
		document.write("</marquee>")
	}
}

function marquee_logo_news()
{
	if (version == "e4")
	{
		document.write("<marquee style='BOTTOM: 0px; FONT-WEIGHT: 1200px; HEIGHT:31px;  TEXT-ALIGN: left; TOP: 0px' id='link_map' scrollamount='2' scrolldelay='10' behavior='alternate' direction='right' border='0' onmouseover='this.stop()' onmouseout='this.start()'>")
	}
}

function marquee3()
{
	if (version == "e4" | version == "n3")
	{
		document.write("<marquee direction='left' border='0' onmouseover='this.stop()' onmouseout='this.start()'>")
	}
}

function marquee4()
{
	if (version == "e4" | version == "n3")
	{
		document.write("</marquee>")
	}
}

function marquee5()
{
	if (version == "e4" | version == "n3")
	{
		document.write("<marquee style='BOTTOM: 0px; FONT-WEIGHT: 100px; HEIGHT:100px;  TEXT-ALIGN: left; TOP: 0px' id='news' scrollamount='1' scrolldelay='10' behavior='loop' direction='up' border='0' onmouseover='this.stop()' onmouseout='this.start()'>")
	}
}

function marquee6()
{
	if (version == "e4" | version == "n3")
	{
		document.write("</marquee>")
	}
}

//-->
</script>
<script language="JavaScript1.2">
function makevisible(cur,which){
if (which==0)
cur.filters.alpha.opacity=100
else
cur.filters.alpha.opacity=20
}
</script>

<style type="text/css">
</style>
</head>
<body topmargin="0" >
<!--#include file="top.asp"-->
<%
dim typeid
dim typename
dim typecontent
set rs=server.CreateObject("ADODB.RecordSet")
rs.Source="select * from "& db_Type_Table &" where typeview=1 order by typeorder"
rs.Open rs.Source,conn,1,1
i=1
Dim ArraytypeID(10000),ArraytypeName(10000)

rs.close
set rs=nothing

set rs=server.CreateObject("ADODB.RecordSet")
rs.Source="select * from "& db_Type_Table &" where typeview=1 order by typeorder"
rs.Open rs.Source,conn,1,1
i=1
Dim ArraytyID(10000),ArraytyName(10000),Arraytyview(10000)
if not rs.EOF then
rseof=1

while not rs.EOF
RecordCount=rs.RecordCount
tyID=rs("typeID")
tyName=trim(rs("typeName"))
tycontent=rs("typecontent")
tyview=trim(rs("typeview"))
ArraytyID(i)=tyID
ArraytyName(i)=tyName
Arraytyview(i)=tyview
i=i+1

rs.MoveNext
wend
rs.close
%>
  
<table width="760" border="0" align="center" cellpadding="0" cellspacing="0">
<tr align="right" valign="bottom">
    <td height="30" colspan="3" background="pic/news_0001.gif"><table width="590" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <TD width="440" height=19>
          <script language="JavaScript">marquee3();</script>
          <%
set rs9=server.CreateObject("ADODB.RecordSet") 
rs9.Source="select top 6 * from "& db_News_Table &" where checkked=1 order by newsid desc"
rs9.Open rs9.Source,conn,1,1
do while not rs9.EOF
title=trim(rs9("title"))
newsurl="ReadNews.asp?NewsID=" & rs9("NewsID")
newswwwurl=rs9("titleface")
%>
          <a class=middle target="_blank" href="<%if rs9("titleface")="无" then %><%=newsurl%><% else %> <%=newswwwurl%><%end if%>" title="<%=htmlencode4(title)%>">
          <img src="images/d.gif" width="13" height="10" border="0"><font color=<%=rs9("titlecolor")%>><%=CutStr(htmlencode4(rs9("title")),14)%></font>&nbsp;&nbsp;
          <%rs9.movenext
loop
rs9.close
set rs9=nothing%>
      <script language="JavaScript">marquee4();</script>
        </a> </TD>
        <TD width="150">&nbsp;</TD>
      </tr>
    </table></td>
  </tr>
  <tr> 
    <td width="160" rowspan="2" valign="top" background="images/left.gif"> 
      <table width="100%" border="0" cellpadding="3" cellspacing="0">
        <tr> 
          <td valign="top"> </td>
        </tr>
        <tr > 
          <td height="34" background="pic/left1-m3.gif">&nbsp;</td>
        </tr>
        <tr > 
          <td align="center" valign=top  ><br>
            <table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="98%" id="AutoNumber5">
<% 
 dim ii
ii = 0
if uselevel=1 then
if Request.cookies(Forcast_SN)("key")="" then
rs.Source="select top " & top_txt & " * from "& db_News_Table &" where updatetime>now()-30 and checkked=1 and newslevel=0 order by click DESC"   '选择本月
end if
if Request.cookies(Forcast_SN)("key")="selfreg" then
	if Request.cookies(Forcast_SN)("reglevel")="3" then
	rs.Source="select top " & top_txt & " * from "& db_News_Table &" where updatetime>now()-30 and checkked=1 order by click DESC"   '选择本月
	end if
	if Request.cookies(Forcast_SN)("reglevel")="2" then
	rs.Source="select top " & top_txt & " * from "& db_News_Table &" where updatetime>now()-30 and checkked=1 order by click DESC"   '选择本月
	end if
	if Request.cookies(Forcast_SN)("reglevel")="1" then
	rs.Source="select top " & top_txt & " * from "& db_News_Table &" where updatetime>now()-30 and checkked=1 order by click DESC"   '选择本月
	end if
end if
if Request.cookies(Forcast_SN)("key")="super" or Request.cookies(Forcast_SN)("key")="typemaster" or Request.cookies(Forcast_SN)("key")="bigmaster" or Request.cookies(Forcast_SN)("key")="smallmaster" or Request.cookies(Forcast_SN)("key")="check" then
rs.Source="select top " & top_txt & " * from "& db_News_Table &" where updatetime>now()-30 and checkked=1 order by click DESC"   '选择本月
end if
else
rs.Source="select top " & top_txt & " * from "& db_News_Table &" where updatetime>now()-30 and checkked=1 order by click DESC"   '选择本月
end if
rs.Open rs.Source,conn,1,1
	if rs.bof and rs.eof then 
		response.write "<td align=center><br>暂 无<br><br></td>" 
	else 
	do while not rs.eof 
%>
              <tr> 
                <td height=18> · 
                  <%if rs("picnews")=1 then%>
                  <img src="images/news_img.gif"> 
                  <%end if%>
                  <a class=middle href="ReadNews.asp?NewsID=<%=rs("NewsID")%>" title="<%=htmlencode4(rs("title"))%>" target="_blank" >
				  <%=CutStr(htmlencode4(rs("title")),14)%>
                  </a></td>
			  </tr>
<%  ii = ii + 1
    if ii>top_txt-1 then exit do
	rs.movenext     
	loop
	end if  
	rs.close   
	set rs=nothing
%>
</table>
            <br>
          </td>
</tr>
        <tr >
          <%if showspecial="1" then%>
		  <td height="34" background="pic/left1-m4.gif" ></td>
        </tr>
          <tr>
            
          <td height="50" ><br>
            <%if showspecial=1 then%>
            <%set rs2=server.CreateObject("ADODB.RecordSet")  '专题
rs2.Source="select Top " & top_sp & " * from "& db_Special_Table &"  order by SpecialID DESC "
rs2.Open rs2.Source,conn,1,1
if not rs2.EOF then
while not rs2.EOF

TrString="&nbsp;·&nbsp;<a class=class href='Special_News.asp?SpecialID=" & rs2("SpecialID") &"'>" & trim(CutStr(htmlencode4(rs2("SpecialName")),14)) & "</a><br>"
Response.Write TrString

rs2.MoveNext
wend
%>
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a class=class href=Special.asp>更多专题...</a> 
            <%
else
Response.Write "<tr><td width=100% align=center height=40>暂无专题</td></tr>"
end if
rs2.Close
set rs2=nothing

  
%>
            <br>
            <br>
          </td>
        </tr><%end if%><%end if%>
          
          <tr ><% if showdata=1 then%> 
            <td height="34" align="center" background="pic/left1-m5.gif">&nbsp;</td>
          </tr>
          <tr > 
            <td > <br> 

		<%								
		function gettipnum()
			dim tmprs
			tmprs=conn.execute("Select Count(NewsID) from "& db_News_Table &" where checkked=1")
			gettipnum=tmprs(0)
			set tmprs=nothing
			if isnull(gettipnum) then gettipnum=0
		end function

		function todays()
			dim tmprs
			tmprs=conn.execute("Select count(NewsID) from "& db_News_Table &" Where checkked=1 and year(updatetime)=year(date()) and month(updatetime)=month(date()) and day(updatetime)=day(date())")
			todays=tmprs(0)
			set tmprs=nothing
			if isnull(todays) then todays=0
		end function

		function getusernum()
			dim rs
			rs=ConnUser.execute("Select Count("& db_User_ID &") from "& db_User_Table &"")
			getusernum=rs(0)
			set rs=nothing
			if isnull(getusernum) then getusernum=0
		end function

		function getgg()
			dim rs
			if db_Sex_Select = "FeiTium" then
				rs=ConnUser.execute("Select Count("& db_User_Id &") from "& db_User_Table &" where  "& db_User_Sex &"='男' ")
				getgg=rs(0)
				set rs=nothing
			else
				if db_Sex_Select = "Number" then
					rs=ConnUser.execute("Select Count("& db_User_Id &") from "& db_User_Table &" where  "& db_User_Sex &"=1")
					getgg=rs(0)
					set rs=nothing
				end if
			end if
			if isnull(getgg) then getgg=0
		end function

		function getmm()
			dim rs
			if db_Sex_Select = "FeiTium" then
				rs=ConnUser.execute("Select Count("& db_User_Id &") from "& db_User_Table &" where  "& db_User_Sex &"='女' ")
				getmm=rs(0)
				set rs=nothing
			else
				if db_Sex_Select = "Number" then
					rs=ConnUser.execute("Select Count("& db_User_Id &") from "& db_User_Table &" where  "& db_User_Sex &"=0")
					getmm=rs(0)
					set rs=nothing
				end if
			end if
			if isnull(getmm) then getmm=0
		end function

		function getother()
			dim rs
			if db_Sex_Select = "FeiTium" then
				rs=ConnUser.execute("Select Count("& db_User_ID &") from "& db_User_Table &" where  "& db_User_Sex &" = '保密' ")
				getother=rs(0)
				set rs=nothing
			else
				if db_Sex_Select = "Number" then
					rs=ConnUser.execute("Select Count("& db_User_ID &") from "& db_User_Table &" where  "& db_User_Sex &" <>1 and  "& db_User_Sex &"<>0 ")
					getother=rs(0)
					set rs=nothing
				end if
			end if
			if isnull(getother) then getother=0
		end function
		%>
		&nbsp;&nbsp;○- 今日文章：<font color=red><%=todays()%></font><br> 
		&nbsp;&nbsp;○- 文章总数：<font color=red><%=gettipnum()%></font><br>
		<%if showcount=1 then%>
			<script src=Cnt.asp></script>
			<!--#include file=zx.asp -->
				  <br> &nbsp;&nbsp;○- 当前在线：<font color=red><%=i%></font><br><br>
		<%end if%>
            </td>
          </tr>
          <%end if%>
		  <tr> 
            <td valign="top"> </td>
          </tr>
          
          <%if showuser=1 then%>
          <tr > 
            <td height="34" valign="middle" background="pic/left1-m6.gif"> <p align="center"></td>
          </tr>
          
        <tr > 
          <td valign="middle" background="images/left.gif"  ><br>
          权希军&nbsp;&nbsp; | &nbsp;&nbsp;刘炳森<br>
          田英章&nbsp;&nbsp; | &nbsp;&nbsp;旭&nbsp;&nbsp;宇<br>
          邹德忠&nbsp;&nbsp; | &nbsp;&nbsp;刘&nbsp;&nbsp;江<br>
          王祥之&nbsp;&nbsp; | &nbsp;&nbsp;肖&nbsp;&nbsp;峰<br>
          孟天宇&nbsp;&nbsp; | &nbsp;&nbsp;王渊华<br>
          何昌贵&nbsp;&nbsp; | &nbsp;&nbsp;丘程光<br>
            崔学路&nbsp;&nbsp; | &nbsp;&nbsp;沈鸿根<br>
            <br>
          </td>
            <%end if%>
          </tr>

        
              <tr> 
                <td height="34" background="pic/left1-m7.gif"> </td>
		<tr>
          <TD height="30"><br>
            本站编辑部<br>
            名誉站长:X X<br>
            站长总编:陈先生<br>
            作品编辑:林先生<br>
            页面维护:林 某<br>
            技术支持:林森森<br>
            广告部主任:王小姐<br>
            拓展部主任:王玉<br>
            <br>
            <br>
            地 址：中国·广东省·佛山<br>
            顺德顺风山公园青龙阁.<br>
            邮 编：528300 　<br>
            <br>
            电 话：0757-22222222 <br>
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;22222223<br>
            传 真：0757-22222228</TD>
        </tr>
      </table>
</td>
    <td width="10" rowspan="2" background="images/left-s.gif"></td>
    <td width="590" align="center" valign="top">      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="450" valign="top" bgcolor="#FFFFFF"> 
            <!--#include file=DaoDu.asp -->
          </td>
          <td width="140" align="right" valign="top" background="images/right.gif" bgcolor="#ECECEC">
            <table width="100%" border="0" cellpadding="0" cellspacing="0">
              <tr> 
          <td width="100%" height="20" valign="middle" background="pic/bzgg_0001.gif"> 
            <p align="center"></td>
        </tr>
			  <%
set rsb=server.CreateObject("ADODB.RecordSet") 
rsb.Source="select * from "& db_Board_Table &" where inuse=1"
rsb.Open rsb.Source,conn,1,1
if not rsb.EOF then
%>
        
        <tr> 
          <td height="118" align=center> <br> 
            <table width="90%" height="60" border="0" cellpadding="0" cellspacing="0">
		<tr> 
                <td align="center"> 
		
		<%=CutStr(htmlencode4(rsb("title")),20)%>

		<script language=JavaScript>marquee1();</script> <%=CheckStr(rsb("content"))%> 
		<script language=JavaScript>marquee2();</script> <%=rsb("dateandtime")%><br> 
		<a href="Board.asp" target="_blank" class=class>以前公告</a></td>
		</tr>
	</table>
	</td>
        </tr>
        <%else Response.Write "<tr><td align=center><br>暂  无</td></tr>"
		end if
     rsb.close
      set rsb=nothing%>
          </table>
          <br>
          <table width="100%"  border="0" cellpadding="0" cellspacing="0">
            <tr>
              <td height="31" background="pic/rwts_0001.gif">&nbsp;</td>
            </tr>
              <tr> 
                <td align="right"><table width="75%" border="0" cellspacing="0" cellpadding="3">
                    <tr> 
                      <td><img src="pic/temp_001.gif" width="130" height="107"></td>
                    </tr>
                    <tr>
                      <td align="center">风情人物画家刘选让</td>
                    </tr>
                  </table></td>
            </tr>
          </table></td>
        </tr>
        <tr> 
          <td height="32" colspan="2" background="images/photonews.gif">&nbsp;</td>
        </tr>
        <tr> 
          <td height="100" colspan="2" align="center" bgcolor="#FFFFFF">
<!--#include file="NewsTopPhoto.asp"-->
					</td>
</tr>
<tr valign="top">
	<td background="images/t1.gif" height="8" colspan="2">&nbsp;</td>
</tr>
<tr valign="top" bgcolor="#FFFFFF"> 
	<td colspan="2">
<%
for i=1 to RecordCount
typeID=Arraytyid(i)
typeName=ArraytyName(i)
%>
	<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%" id="AutoNumber6">
		<tbody>
			<tr> 
				<td width="49%" valign="top" bgcolor="#FFFFFF">
					<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="100%" id="AutoNumber12">
						<tbody>
							<tr valign="middle"> 
								<td background="images/lmtop.gif" height="30" colspan="2"><font class="m_tittle">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=typeName%></font></td>
							</tr>
							<tr> 
								<td width="10%" height="118" align="left" valign="top" bgcolor="#CCCCCC"> 
<%
set rs=server.CreateObject("ADODB.RecordSet")
if uselevel=1 then
	if Request.cookies(Forcast_SN)("key")="super" or Request.cookies(Forcast_SN)("key")="typemaster" or Request.cookies(Forcast_SN)("key")="bigmaster" or Request.cookies(Forcast_SN)("key")="smallmaster" or Request.cookies(Forcast_SN)("key")="check" then
		rs.Source="select top 1 * from "& db_News_Table &" where typeid=" & typeid &" and checkked=1 and picnews=1 and picname is not null order by newsid DESC"
	end if
	if Request.cookies(Forcast_SN)("key")="" then
		rs.Source="select top 1 * from "& db_News_Table &" where (typeid=" & typeid &" and checkked=1 and newslevel=0 and picnews=1 and picname is not null) order by newsid DESC"
	end if
	if Request.cookies(Forcast_SN)("key")="selfreg" then
		if Request.cookies(Forcast_SN)("reglevel")=3 then
			rs.Source="select top 1 * from "& db_News_Table &" where (typeid=" & typeid &" and checkked=1 and picnews=1 and picname is not null) order by newsid DESC"
		end if
		if Request.cookies(Forcast_SN)("reglevel")=2 then
			rs.Source="select top 1 * from "& db_News_Table &" where (typeid=" & typeid &" and checkked=1 and picnews=1 and picname is not null) order by newsid DESC"
		end if
		if Request.cookies(Forcast_SN)("reglevel")=1 then
			rs.Source="select top 1 * from "& db_News_Table &" where (typeid=" & typeid &" and checkked=1 and picnews=1 and picname is not null) order by newsid DESC"
		end if
	end if
else
	rs.Source="select top 1 * from "& db_News_Table &" where typeid=" & typeid &" and checkked=1 and picnews=1 and picname is not null order by newsid DESC"
end if
rs.Open rs.Source,conn,1,1
if rs.EOF then
	Response.Write "<img src='images/notopic.gif' width=56 border=0 height=56>"
else
	title=rs("title")
	fileExt=lcase(getFileExtName(rs("picname")))
	%>
	<a class="class" target="_blank" href="ReadNews.asp?newsid=<%=rs("newsid")%>" title="<%=htmlencode4(title)%>"> 
	<%if fileext="jpg" or fileext="bmp" or fileext="png" or fileext="gif" then%>
		<img src="<%=FileUploadPath & rs("picname")%>" width="56" border=1 style=border-color:#ffffff height="56">
	</a> 
	<%end if%>
	<%if fileext="swf" then%>
		<object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0" width="56" height="56">
		<param name="movie" value="<%=FileUploadPath & rs("picname")%>">
		<param name="quality" value="high">
		<param name="Play" value="-1">
		<param name="Loop" value="0">
		<param name="Menu" value="-1">
		<embed src="<%=FileUploadPath & rs("picname")%>" width="56" pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash"></embed></object> 
	<%end if%>
	<br>
	<%
	end if
	rs.close
	set rs=nothing
	%>
	</td>
	<td width="90%" valign="top" bgcolor="#ECECEC">
		<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber11">
	<% dim rs3
	set rs3=server.CreateObject("ADODB.RecordSet")
	if uselevel=1 then
		if Request.cookies(Forcast_SN)("key")="super" or Request.cookies(Forcast_SN)("key")="typemaster" or Request.cookies(Forcast_SN)("key")="bigmaster" or Request.cookies(Forcast_SN)("key")="smallmaster" or Request.cookies(Forcast_SN)("key")="check" then
			rs3.Source="select top 6 * from "& db_News_Table &" where (typeid=" & typeid &" and checkked=1) order by newsid DESC"
		end if
		if Request.cookies(Forcast_SN)("key")="" then
			rs3.Source="select top 6 * from "& db_News_Table &" where (typeid=" & typeid &" and checkked=1) order by newsid DESC"
		end if
		if Request.cookies(Forcast_SN)("key")="selfreg" then
			if Request.cookies(Forcast_SN)("reglevel")=3 then
				rs3.Source="select top 6 * from "& db_News_Table &" where (typeid=" & typeid &" and checkked=1 ) order by newsid DESC"
			end if
			if Request.cookies(Forcast_SN)("reglevel")=2 then
				rs3.Source="select top 6 * from "& db_News_Table &" where (typeid=" & typeid &" and checkked=1 ) order by newsid DESC"
			end if
			if Request.cookies(Forcast_SN)("reglevel")=1 then
				rs3.Source="select top 6 * from "& db_News_Table &" where (typeid=" & typeid &" and checkked=1 ) order by newsid DESC"
			end if
		end if
	else
		rs3.Source="select top 6 * from "& db_News_Table &" where (typeid=" & typeid &" and checkked=1) order by newsid DESC"
	end if
	rs3.Open rs3.Source,conn,1,1
	while not rs3.EOF
		newsurl="ReadNews.asp?NewsID=" & rs3("NewsID")
		newswwwurl=rs3("titleface")
		if showyear=1 then
			datetime="<font class=middle>(" & year(rs3("UpdateTime"))  &"年"& Month(rs3("UpdateTime"))  &"月"& Day(rs3("UpdateTime")) &"日)</font>"
		else
			datetime="<font class=middle>("& Month(rs3("UpdateTime"))  &"月"& Day(rs3("UpdateTime")) &"日)</font>"
		end if
		if rs3("picnews")=1 then
			img=" <img src='images/news_img.gif' border='0'>"
		else
			img=""
		end if

		'增加小类名称程序开始
		B=rs3("bigclassID")
		set rs22=server.createobject("adodb.recordset")
		sql="SELECT bigclassID,bigclassname FROM "& db_BigClass_Table &" WHERE bigclassID="&B
		Set rs22=Conn.Execute(sql)
		A=rs22("bigclassname")
		rs22.close
		set rs22=Nothing
		'增加小类名称程序结束
		%>
                              <tbody>
                                <tr> 
                                  <td> <table width="100%" border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111">
                                      <tbody>
                                        <tr> 
                                          <td height="18">[<%=A%>]<%=img%>
                                          <a class="middle" href="<%if rs3("titleface")="无" then %><%=newsurl%><% else %> <%=newswwwurl%><%end if%>" title="<%=htmlencode4(rs3("title"))%>" target="_blank">
										  <font color="<%=rs3("titlecolor")%>">
										  <%=CutStr(htmlencode4(rs3("title")),14)%>
										  </font>
										  </a>
						  <!--标题后评论提示-->

						  <% if rs3("titlesize")>=1 then %>
						  <A class=middle HREF="<%=path%>Review.asp?NewsID=<%=rs3("NewsID")%>" target="_blank" ><font color=red><b>评</b></font></A>
						  <%end if %>

						  <!--标题后评论提示-->

											<%''if showauthor="1" then%>
											<%''=rs3("Author")%>
											<%''end if%>
											
										  </td>
                                        </tr>
                                      </tbody>
                                    </table></td>
                                </tr>
                                <%rs3.MoveNext
wend
%>
                              </tbody>
                            </table> </td>
                        </tr>
                        <tr> 
                          <td width="10%" align="right" bgcolor="#CCCCCC"></td>
                          <td width="90%" align="right" bgcolor="#ECECEC"><a class="class" href="Type.asp?typeid=<%=typeid%>"><img src="images/more.gif" border="0" alt="更多<%=typeName%>" /></a></td>
                        </tr>
                      </tbody>
                  </table></td>
                  <%i=i+1
              typeID=Arraytyid(i)
              typeName=ArraytyName(i)
              if i<=RecordCount then
              %>
                  <td width="2%" valign="top" background="images/t.gif"></td>
                  <td width="49%" valign="top" bgcolor="#FFFFFF"> <table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="100%" id="AutoNumber13">
                      <tbody>
                        <tr > 
                          <td background="images/lmtop.gif" height="30" colspan="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font class="m_tittle"><%=typeName%></font></td>
                        </tr>
                        <tr> 
                          <td width="10%" height="118" align="left" valign="top" bgcolor="#CCCCCC"> 
                            <%
set rs=server.CreateObject("ADODB.RecordSet")
if uselevel=1 then
if Request.cookies(Forcast_SN)("key")="super" or Request.cookies(Forcast_SN)("key")="typemaster" or Request.cookies(Forcast_SN)("key")="bigmaster" or Request.cookies(Forcast_SN)("key")="smallmaster" or Request.cookies(Forcast_SN)("key")="check" then
rs.Source="select top 1 * from "& db_News_Table &" where typeid=" & typeid &" and checkked=1 and picnews=1 and picname is not null order by newsid DESC"
end if
if Request.cookies(Forcast_SN)("key")="" then
rs.Source="select top 1 * from "& db_News_Table &" where (typeid=" & typeid &" and checkked=1 and newslevel=0 and picnews=1 and picname is not null) order by newsid DESC"
end if
if Request.cookies(Forcast_SN)("key")="selfreg" then
	if Request.cookies(Forcast_SN)("reglevel")=3 then
	rs.Source="select top 1 * from "& db_News_Table &" where (typeid=" & typeid &" and checkked=1 and picnews=1 and picname is not null) order by newsid DESC"
	end if
	if Request.cookies(Forcast_SN)("reglevel")=2 then
	rs.Source="select top 1 * from "& db_News_Table &" where (typeid=" & typeid &" and checkked=1 and picnews=1 and picname is not null) order by newsid DESC"
	end if
	if Request.cookies(Forcast_SN)("reglevel")=1 then
	rs.Source="select top 1 * from "& db_News_Table &" where (typeid=" & typeid &" and checkked=1 and picnews=1 and picname is not null) order by newsid DESC"
	end if
end if
else
rs.Source="select top 1 * from "& db_News_Table &" where typeid=" & typeid &" and checkked=1 and picnews=1 and picname is not null order by newsid DESC"
end if
rs.Open rs.Source,conn,1,1
if rs.EOF then
Response.Write "<img src=images/notopic.gif width=56 border=0 height=56>"
else
title=rs("title")
fileExt=lcase(getFileExtName(rs("picname")))
%>
<a class="class" target="_blank" href="ReadNews.asp?newsid=<%=rs("newsid")%>" title="<%=htmlencode4(title)%>"> 
				<%if fileext="jpg" or fileext="bmp" or fileext="png" or fileext="gif" then%>
					<img src="<%=FileUploadPath & rs("picname")%>" width="56" border=1 style=border-color:#ffffff height="56" /></a> 
				<%end if%>
				<%if fileext="swf" then%>
					<object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0" width="56" height="56">
					<param name="movie" value="<%=FileUploadPath & rs("picname")%>" />
					<param name="quality" value="high" />
					<param name="Play" value="-1" />
					<param name="Loop" value="0" />
					<param name="Menu" value="-1" />
					<embed src="<%=FileUploadPath & rs("picname")%>" width="56" pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash"></embed>
					</object> 
				<%end if%>
				<br>
				<%
end if
rs.close
set rs=nothing
%>
</td>
                          <td width="90%" valign="top" bgcolor="#ECECEC"> <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber14">
                              <%
set rs3=server.CreateObject("ADODB.RecordSet")
if uselevel=1 then
if Request.cookies(Forcast_SN)("key")="super" or Request.cookies(Forcast_SN)("key")="typemaster" or Request.cookies(Forcast_SN)("key")="bigmaster" or Request.cookies(Forcast_SN)("key")="smallmaster" or Request.cookies(Forcast_SN)("key")="check" then
rs3.Source="select top 6 * from "& db_News_Table &" where (typeid=" & typeid &" and checkked=1) order by newsid DESC"
end if
if Request.cookies(Forcast_SN)("key")="" then
rs3.Source="select top 6 * from "& db_News_Table &" where (typeid=" & typeid &" and checkked=1) order by newsid DESC"
end if
if Request.cookies(Forcast_SN)("key")="selfreg" then
	if Request.cookies(Forcast_SN)("reglevel")=3 then
	rs3.Source="select top 6 * from "& db_News_Table &" where (typeid=" & typeid &" and checkked=1 ) order by newsid DESC"
	end if
	if Request.cookies(Forcast_SN)("reglevel")=2 then
	rs3.Source="select top 6 * from "& db_News_Table &" where (typeid=" & typeid &" and checkked=1 ) order by newsid DESC"
	end if
	if Request.cookies(Forcast_SN)("reglevel")=1 then
	rs3.Source="select top 6 * from "& db_News_Table &" where (typeid=" & typeid &" and checkked=1 ) order by newsid DESC"
	end if
end if
else
	rs3.Source="select top 6 * from "& db_News_Table &" where (typeid=" & typeid &" and checkked=1) order by newsid DESC"
end if
rs3.Open rs3.Source,conn,1,1
while not rs3.EOF
if showyear=1 then
newsurl="ReadNews.asp?NewsID=" & rs3("NewsID")
newswwwurl=rs3("titleface")
datetime="<font class=middle>(" & year(rs3("UpdateTime"))  &"年"& Month(rs3("UpdateTime"))  &"月"& Day(rs3("UpdateTime")) &"日)</font>"
else
newsurl="ReadNews.asp?NewsID=" & rs3("NewsID")
newswwwurl=rs3("titleface")
datetime="<font class=middle>("& Month(rs3("UpdateTime"))  &"月"& Day(rs3("UpdateTime")) &"日)</font>"
end if
if rs3("picnews")=1 then
img=" <img src='images/news_img.gif' border='0'>"
else
img=""
end if
%>
                              <%
''''增加小类名称程序开始
B=rs3("bigclassID")
set rs22=server.createobject("adodb.recordset")
sql="SELECT bigclassID,bigclassname FROM "& db_BigClass_Table &" WHERE bigclassID="&B
Set rs22=Conn.Execute(sql)
A=rs22("bigclassname")
rs22.close
set rs22=Nothing
''''增加小类名称程序结束
%>
                              <tbody>
                                <tr> 
                                  <td><table width="100%" border="0" cellpadding="2" cellspacing="0">
                                      <tbody>
                                        <tr> 
                                          <td height="18">[<%=A%>]<%=img%>
                                          <a class="middle" href="<%if rs3("titleface")="无" then %><%=newsurl%><% else %> <%=newswwwurl%><%end if%>" title="<%=htmlencode4(rs3("title"))%>" target="_blank">
										  <font color="<%=rs3("titlecolor")%>"> 
                                          <%=CutStr(htmlencode4(rs3("title")),14)%>
                                          </font>
										  </a>

<!--标题后评论提示-->
<% if rs3("titlesize")>=1 then %>
<A class=middle HREF="<%=path%>review.asp?NewsID=<%=rs3("NewsID")%>" target="_blank" ><font color=red><b>评</b></font></A>
<%end if %>

<!--标题后评论提示-->
<%''if showauthor="1" then%>
<%''=rs3("Author")%>
<%''end if%>
</td>
                                        </tr>
                                      </tbody>
                                    </table></td>
                                </tr>
                                <%rs3.MoveNext
wend
%>
                              </tbody>
                            </table></td>
                        </tr>
                        <tr> 
                          <td width="10%" align="right" bgcolor="#CCCCCC"></td>
                          <td width="90%" align="right" bgcolor="#ECECEC"><a class="class" href="Type.asp?typeid=<%=typeid%>"><img src="images/more.gif" border="0" alt="更多<%=typeName%>" /></a></td>
                        </tr>
                      </tbody>
                  </table><br></td>
                  <%end if%>
                </tr>
              </tbody>
            </table>
          <%
next
rs3.close
set rs3=nothing
%> </td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td align="center" valign="top" bgcolor="#FFFFFF">　</td>
  </tr>
</table>
  <table width="760" border="0"  align="center" cellspacing="0" cellpadding="0">
    <tr>
      <td width="160" valign="top" background="images/left.gif"></td>
      <td width="10" background="images/left-s.gif">　</td>
      <td width="590" valign="top" bgcolor="#FFFFFF">&nbsp;</td>
    </tr>
    <tr>
      <td background="images/db.gif" height="30" colspan="3" valign="top">&nbsp;</td>
    </tr>
  </table>

  <table width="760" align="center" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td valign="bottom"><div align="center"></div></td>
    </tr>
  </table>
  
  <table width="760" height="31" align="center" border="0" cellpadding="0" cellspacing="0">
    <tr>
<%if showlinkmap=1 then%><td colspan="2" align="center" height="31">
	<%if linkshownum >8 then
		Response.Write "<script language=JavaScript>marquee_logo_news();</script><P align=left>"& vbcrlf
	end if
	set rs10=server.CreateObject("ADODB.RecordSet") 
	rs10.Source="select top "& linkshownum &" * from "& db_Link_Table &" where linktype=2 and pass=1 order by ID DESC "
	rs10.Open rs10.Source,conn,1,1
	for i=1 to linkshownum
		if not rs10.EOF then%>
<a href="<%=rs10("weburl")%>" target="_blank" title="<%=rs10("webname")%>&#13;&#10;简介：<%=rs10("content")%>&#13;&#10;站长：<%=rs10("webmaster")%>&#13;&#10;申请时间：<%=rs10("dateandtime")%>"><img src="<%=rs10("logo")%>" width="88" height="31" border="0" align=left></a>
		<%rs10.MoveNext
		else%>
<a href="#" onclick="javascript:linkreg()"><img src="images/logo.gif" width=88 height=31 border=0 align=left></a>
		<%end if
	Next%>
	</td>
	<%
	rs10.Close
	set rs10=nothing
end if%>
	</tr>
</table>
<table width="760" border="0" align="center" cellpadding="0" cellspacing="0" >
  <tr>
<%if showlink=1 then
set rs10=server.CreateObject("ADODB.RecordSet") 
rs10.Source="select top "& linkshownum &" * from "& db_Link_Table &" where linktype=0 and pass=1 order by ID DESC "
rs10.Open rs10.Source,conn,1,1
for i=1 to 6
	if not rs10.EOF then
	%>
<td height=31 align="center"><a class="middle" href="<%=rs10("weburl")%>" target="_blank" title="<%=rs10("content")%>
站长：<%=rs10("webmaster")%>
申请时间：<%=rs10("dateandtime")%>"><%=rs10("webname")%></a></td>
<%
rs10.MoveNext
else
%>
      <td height=31 align="center"><a class="middle" href="#" onclick="javascript:linkreg()">您的位置</a></td>
      <%end if
	Next
rs10.Close
set rs10=nothing
%>
      
    <td align="center">&nbsp;</td>
    </tr><%end if%>
  </table>
<%else
Response.Write "<table width=760 align=center border=0 height=50><tr><td align=center>暂 无 文 章 类 别，请 <a href=login.asp>登 陆</a> 后 添 加！</td></tr></table>"
end if%>
<!--#include file=Bottom.asp -->
