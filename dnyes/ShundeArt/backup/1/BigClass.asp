<%@ Language=VBScript%>

<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<!--#include file="function.asp"-->
<!--#include file="char.inc"-->
<%

typeid=checkstr(request("typeid"))
request_BigClassID=Request.QueryString("BigClassID")
if typeid="" or request_bigclassid="" then
	Response.Write "<script>alert('未指定参数');history.back()</script>"
	response.end
else
	if not IsNumeric(request_BigClassID) or not IsNumeric(typeid) then
		response.write "<script>alert('非法参数');history.back()</script>"
		response.end
	else
		set rs=server.CreateObject("ADODB.RecordSet")
		rs.Source="select * from "& db_BigClass_Table &" where bigclassid=" & request_bigclassid &" order by BigClassorder"
		rs.Open rs.Source,conn,1,1
		bigclassname=rs("bigclassname")
		rs.close
		set rs=nothing
		PageShowSize = 10            '每页显示多少个页
		MyPageSize   = 20           '每页显示多少条新闻
			
		If Not IsNumeric(Request("page")) Or IsEmpty(Request("page")) Or Request("page") <=0 Then
			MyPage=1
		Else
			MyPage=Int(Abs(Request("page")))
		End if
		set rs=server.CreateObject("ADODB.RecordSet")
		rs.Source="select * from "& db_BigClass_Table &" where typeid=" & typeid &" order by BigClassorder"
		rs.Open rs.Source,conn,1,1
		bigclasszs=rs("bigclasszs")
		
		dim ArrayBigClassID(10000),ArrayBigClassName(10000)
		BigClassCount=rs.RecordCount
		
		for i=1 to BigClassCount
			ArrayBigClassID(i)=rs("BigClassID")
			ArrayBigClassName(i)=rs("BigClassName")
			rs.MoveNext
		next
		
		rs.close
		set rs=nothing
		
		set rs2=server.CreateObject("ADODB.RecordSet")  '专题
		rs2.Source="select Top " & top_sp & " * from "& db_Special_Table &"  order by SpecialID DESC "
		rs2.Open rs2.Source,conn,1,1
		rs2.close
		set rs2=nothing
		
		set rs4=server.CreateObject("Adodb.RecordSet")
		rs4.source="select * from "& db_SmallClass_Table &" Where BigClassid=" & request_BigClassid &" order by SmallClassorder"
		rs4.open rs4.source,conn,1,1
		
		SmallClassCount=rs4.RecordCount
		dim ArraySmallClassID(10000),ArraySmallClassName(10000)
		
		for i=1 to SmallClassCount
			ArraySmallClassID(i)=rs4("SmallClassID")
			ArraySmallClassName(i)=rs4("SmallClassName")
			rs4.MoveNext
		next
		
		rs4.Close
		set rs4=nothing
		
		dim typename
		set rs5=server.CreateObject("ADODB.RecordSet")
		rs5.Source="select * from "& db_Type_Table &" where typeid=" & typeid &" order by typeorder"
		rs5.Open rs5.Source,conn,1,1
		typename=rs5("typename")
		rs5.Close
		set rs5=nothing
		
	end if
	%>
<html>

<head>
<meta http-equiv="Content-Type" content="">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=BigClassName%>__<%=typeName%>__<%=jjgn%></title>
<LINK href=news.css rel=stylesheet>
</head>

<body topmargin="0">
<!--#include file="top.asp"-->
<table width="760" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <tr valign="top"> 
    <td width="180" background="IMAGES/menu-d.gif" bgcolor="#FFFFFF"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td></td>
        </tr>
        <tr> 
          <td height="645" valign="top" background="IMAGES/menu-d.gif" bgcolor="#FFFFFF"><table width="102%" border="0" cellspacing="0" cellpadding="2">
              <tr bgcolor="#FFFFFF"> 
                <td height="25" align="center" background="IMAGES/menu-d-fl.gif" bgcolor="efefef"><font color="#000000">分类栏目</font></td>
              </tr>
              <tr bgcolor="#FFFFFF"> 
                <td width="30%" height="154" align="center" valign="top" background="IMAGES/menu-d.gif"><table width="92%" border="0" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td> 
                        <%
    dim mode
    typeid=checkstr(request("typeid"))
    set rs=server.CreateObject("ADODB.RecordSet")
rs.Source="select * from "& db_Type_Table &" where typeid="&typeid&" order by typeorder"
rs.Open rs.Source,conn,1,1
mode=rs("mode")
rs.close
set rs=nothing
    if mode=5 then
    set rs=conn.execute("SELECT * FROM "& db_BigClass_Table &" where typeid="&typeid&" order by bigclassorder") 
    if not rs.EOF then
	do while not rs.eof

    %>
                        <table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="100%" id="AutoNumber2">
                          <tr> 
                            <td colspan="2"> <li></li>
                              <%if not Rs.eof then%>
                              <font class=m_tittle>&nbsp;</font><a class=middle href="bigclass.asp?typeid=<%=typeid%>&bigclassid=<%=Rs("bigclassid")%>"><font color="#000000"><%=Rs("bigclassname")%></font></a>&nbsp;</td>
                          </tr>
                          <%set nrs=conn.execute("SELECT * FROM "& db_SmallClass_Table &" where bigclassid="&cstr(rs("bigclassid"))&" order by smallclassorder")%>
                          <%do while not nrs.eof%>
                          <tr> 
                            <td width=100% align=center valign="middle"> 
                              <%if not nRs.eof then%>
                              <a class=class href="SmallClass.asp?typeid=<%=typeid%>&BigClassID=<%=nrs("bigclassid")%>&smallclassid=<%=nrs("smallclassid")%>"><%=nRs("smallclassname")%></a>　</td>
                            <%nrs.movenext   
		     end if %>
                          </tr>
                          <%loop
            nRs.Close
           set nRs=nothing
            %>
                        </table>
                        <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="<%=border%>" width="105%" id="AutoNumber2" height="11">
                          <tr> 
                            <td height=1 align=center></td>
                          </tr>
                        </table>
                        <%rs.movenext   
	end if
	  loop
	rs.close
%>
                        <%else%>
                        <table cellpadding=3 cellspacing=0 width=100% border=0 style="border-collapse: collapse">
                          <tr> 
                            <td height="22">&nbsp;&nbsp;&nbsp;&nbsp;<font color="#000000">暂无大类</font>&nbsp;</td>
                          </tr>
                        </table>
                        <%end if%>
                        <%else%>
                        <%  
typeid=checkstr(request("typeid"))
	set rs=conn.execute("SELECT * FROM "& db_BigClass_Table &" where typeid="&typeid&" order by bigclassorder") 
	do while not rs.eof

%>
                        <table width="100%" border="0" cellpadding="3" cellspacing="0" id="AutoNumber2" style="border-collapse: collapse">
                          <tr> 
                            <td height=22 colspan="2"> <li></li>
                              <%if not Rs.eof then%>
                              <a class=middle href="bigclass.asp?typeid=<%=typeid%>&bigclassid=<%=Rs("bigclassid")%>"><font color="#000000"><%=Rs("bigclassname")%></font></a></td>
                          </tr>
                          <%set nrs=conn.execute("SELECT * FROM "& db_SmallClass_Table &" where bigclassid="&cstr(rs("bigclassid"))&" order by smallclassorder")%>
                          <%do while not nrs.eof%>
                          <tr> 
                            <td width=50% align=center valign="middle"> 
                              <%if not nRs.eof then%>
                              <a class=class href="SmallClass.asp?typeid=<%=typeid%>&BigClassID=<%=nrs("bigclassid")%>&smallclassid=<%=nrs("smallclassid")%>"><%=nRs("smallclassname")%></a>　</td>
                            <%nrs.movenext   
		     end if %>
                            <td width=50% align=center valign="middle"> 
                              <%if not nRs.eof then%>
                              <a class=class href="SmallClass.asp?typeid=<%=typeid%>&BigClassID=<%=nrs("bigclassid")%>&smallclassid=<%=nrs("smallclassid")%>"><%=nRs("smallclassname")%></a>　</td>
                            <%nRs.movenext
             end if%>
                          </tr>
                          <%loop
            nRs.Close
           set nRs=nothing
            %>
                        </table>
                        <%rs.movenext   
	end if
	loop  
	rs.close
%>
                        <%end if%>
                      </td>
                    </tr>
                  </table></td>
              </tr>
            </table>
            <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#999999" width="100%" id="AutoNumber4" height="25">
              <tr> 
                <td height="25" align=center background="IMAGES/menu-d-pic.gif"><font color="#000000">最新图文</font></td>
              </tr>
              <tr> 
                <td width="100%" height="20" align=center><br> 
                  <%
set rs3=server.CreateObject("ADODB.RecordSet")
if uselevel=1 then
if Request.cookies(Forcast_SN)("key")="super" or Request.cookies(Forcast_SN)("key")="typemaster" or Request.cookies(Forcast_SN)("key")="bigmaster" or Request.cookies(Forcast_SN)("key")="smallmaster" or Request.cookies(Forcast_SN)("key")="check" then
rs3.Source ="select top " & top_img & " * from "& db_News_Table &" where picnews=1 and checkked=1 and typeid="&typeid&" and picname is not null order by NewsID DESC"
end if
if Request.cookies(Forcast_SN)("key")="" then
rs3.Source ="select top " & top_img & " * from "& db_News_Table &" where picnews=1 and checkked=1 and newslevel=0 and typeid="&typeid&" and picname is not null order by NewsID DESC"
end if
if Request.cookies(Forcast_SN)("key")="selfreg" then
	if Request.cookies(Forcast_SN)("reglevel")=3 then
rs3.Source ="select top " & top_img & " * from "& db_News_Table &" where picnews=1 and checkked=1  and typeid="&typeid&" and picname is not null order by NewsID DESC"
	end if
	if Request.cookies(Forcast_SN)("reglevel")=2 then
rs3.Source ="select top " & top_img & " * from "& db_News_Table &" where picnews=1 and checkked=1  and typeid="&typeid&" and picname is not null order by NewsID DESC"
	end if
	if Request.cookies(Forcast_SN)("reglevel")=1 then
rs3.Source ="select top " & top_img & " * from "& db_News_Table &" where picnews=1 and checkked=1  and typeid="&typeid&" and picname is not null order by NewsID DESC"
	end if
end if
else
rs3.Source ="select top " & top_img & " * from "& db_News_Table &" where picnews=1 and checkked=1 and typeid="&typeid&" and picname is not null order by NewsID DESC"
end if
rs3.Open rs3.Source,conn,1,1
if not rs3.EOF then
while not rs3.EOF
fileExt=lcase(getFileExtName(rs3("picname")))
Content=htmlencode4(rs3("Content"))

%>
                  <table width="90%" border="0" cellspacing="0" cellpadding="3" align="center" style="TABLE-LAYOUT: fixed">
                    <tr> 
                      <td style="WORD-WRAP: break-word"><a class=middle href="ReadNews.asp?NewsID=<%=rs3("NewsID")%>" target=_blank title="<%=htmlencode4(rs3("title"))%>"> 
                        <%if fileext="jpg" or fileext="bmp" or fileext="png" or fileext="gif" then%>
                        <img  src="<%=FileUploadPath & rs3("picname")%>" width="65" height="65" border=0 align="left"> 
                        <%end if%>
                        <%if fileext="swf" then%>
                        <object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000'  codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0' width="65" height="65" border=0 align="left">
                          <param name=movie value="<%=FileUploadPath & rs3("picname")%>">
                          <param name=quality value=high>
                          <param name='Play' value='-1'>
                          <param name='Loop' value='0'>
                          <param name='Menu' value='-1'>
                          <embed src="<%=FileUploadPath & rs3("picname")%>" width="65" height="65" border=0 align="left" pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash'></embed></object>
                        <%end if%>
                        <%=rs3("title")%></a></div></td>
                    </tr>
                  </table>
                  <%
rs3.MoveNext
wend
else
Response.Write "<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"" align=""center"" style=""TABLE-LAYOUT: fixed""><tr><td width=100% align=center height=18 background='IMAGES/menu-d.gif'>暂无</td></tr></table>"
end if
rs3.close
set rs3=nothing

%>
                </td>
              </tr>
            </table>
            
                    <%end if%>
                    </td>
        </tr>
      </table>
    </td>
    <td width="6">&nbsp;</td>
    <td background="IMAGES/menu-l.gif"> 
      <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td height="25" background="IMAGES/menu-l-m.gif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font class=m_tittle>&nbsp;</font>栏目导航&nbsp;&nbsp;<a href="./" class="daohang">网站首页</a>&gt;&gt;<a href=type.asp?typeid=<%=typeid%> class="daohang"><%=typename%></a> 
            <% Response.Write "&gt;&gt;" & BigClassName & "" %>
          </td>
        </tr>
        <tr> 
          <td background="IMAGES/menu-l.gif"><div align="center"> 
              <script language=javascript src=./zongg/ad.asp?i=13></script>
            </div></td>
        </tr>
      </table>
                  
      <!--模版开始-->
      <% select case mode%>
      <% case "1" %>
      <!--图片模版-->
      <%
set rs3=server.CreateObject("ADODB.RecordSet")
%>
      <table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="100%" id="AutoNumber4">
        <tr> 
          <td height="25" background="IMAGES/menu-l-zj.gif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font class=body><%=BigClassName%></font></td>
        </tr>
        <%if uselevel=1 then
if Request.cookies(Forcast_SN)("key")="super" or Request.cookies(Forcast_SN)("key")="typemaster" or Request.cookies(Forcast_SN)("key")="bigmaster" or Request.cookies(Forcast_SN)("key")="smallmaster" or Request.cookies(Forcast_SN)("key")="check" then
rs3.Source="select  * from "& db_News_Table &" where (BigClassid=" & request_BigClassid & " and checkked=1) order by newsid DESC"
end if
if Request.cookies(Forcast_SN)("key")="" then
rs3.Source="select  * from "& db_News_Table &" where (BigClassid=" & request_BigClassid & " and checkked=1 ) order by newsid DESC"
end if
if Request.cookies(Forcast_SN)("key")="selfreg" then
	if Request.cookies(Forcast_SN)("reglevel")=3 then
rs3.Source="select  * from "& db_News_Table &" where (BigClassid=" & request_BigClassid & " and checkked=1 ) order by newsid DESC"
	end if
	if Request.cookies(Forcast_SN)("reglevel")=2 then
rs3.Source="select  * from "& db_News_Table &" where (BigClassid=" & request_BigClassid & " and checkked=1 ) order by newsid DESC"
	end if
	if Request.cookies(Forcast_SN)("reglevel")=1 then
rs3.Source="select  * from "& db_News_Table &" where (BigClassid=" & request_BigClassid & " and checkked=1 ) order by newsid DESC"
	end if
end if
else
rs3.Source="select  * from "& db_News_Table &" where (BigClassid=" & request_BigClassid & " and checkked=1) order by newsid DESC"
end if
rs3.Open rs3.Source,conn,1,1

if not rs3.EOF then
Rs3.PageSize     = MyPageSize
MaxPages         = Rs3.PageCount
Rs3.absolutepage = MyPage
total            = Rs3.RecordCount
i = 0

%>
        <tr> 
          <td> <table width="90%" border="0" cellspacing="0" cellpadding="3" align="center">
              <!--图片换行显示1-->
<% do until rs3.Eof or i = rs3.PageSize %>
              <tr> 
<%
Content=htmlencode4(rs3("Content"))
content=replace(content,"[---分页---]","")
%>
<% if not rs3.EOF then %>
                <td width=25% align=center valign="top" style="table-layout:fixed; word-break:break-all"> 

                  <div align="center"> 
                    <% content=htmlencode4(rs3("Content"))
                       content=replace(content,"[---分页---]","")%>
                    <a class=middle href="ReadNews.asp?NewsID=<%=rs3("NewsID")%>" target=_blank title="<%=CutStr(nohtml(Content),150)%>"> 
                    <%if   rs3("picname")=("") then%>
                    <img  src="IMAGES/flashorno.gif" width="110" height="80" border=1 style=border-color:#000000 align=top> 
                    <%else%>
                    <%if fileext="jpg" or fileext="bmp" or fileext="png" or fileext="gif" then%>
                    <img  src="<%=FileUploadPath & rs3("picname")%>" width="110" height="80" border=1 style=border-color:#000000 align=top> 
                    <%end if%>

                    <%if fileext="swf" then%>
					<object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000'  codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0' width="110" height="80" border=0 >
					<param name=movie value="<%=FileUploadPath & rs3("picname")%>">
					<param name=quality value=high>
					<param name='Play' value='-1'>
					<param name='Loop' value='0'>
					<param name='Menu' value='-1'>
					<embed src="<%=FileUploadPath & rs3("picname")%>" width="110" height="80" border=0  pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash'></embed>
					</object>
					<%end if%>

                    <%end if%>
                    </a><br>
                    <br>
                    <a class=middle href="ReadNews.asp?NewsID=<%=rs3("NewsID")%>" target=_blank title="<%=htmlencode4(rs3("title"))%>"> 
                    <%=CutStr(rs3("title"),18)%> </a> </div></td>
<%rs3.movenext
i = i + 1
end if
%>
                <!--图片换行显示2-->
                <td width=25% align=center valign="top" style="table-layout:fixed; word-break:break-all"> 
<% if not rs3.EOF then %>
                  <div align="center"> 
                    <%
Content=htmlencode4(rs3("Content"))
content=replace(content,"[---分页---]","")
%>
                    <a class=middle href="ReadNews.asp?NewsID=<%=rs3("NewsID")%>" target=_blank title="<%=CutStr(nohtml(Content),150)%>"> 
                    <%if rs3("picname")=("") then%>
                    <img  src="IMAGES/flashorno.gif" width="110" height="80" border=1 style=border-color:#000000 align=top> 
                    <%else%>
                    <%if fileext="jpg" or fileext="bmp" or fileext="png" or fileext="gif" then%>
                    <img  src="<%=FileUploadPath & rs3("picname")%>" width="110" height="80" border=1 style=border-color:#000000 align=top> 
                    <%end if%>

                    <%if fileext="swf" then%>
					<object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000'  codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0' width="110" height="80" border=0 >
					<param name=movie value="<%=FileUploadPath & rs3("picname")%>">
					<param name=quality value=high>
					<param name='Play' value='-1'>
					<param name='Loop' value='0'>
					<param name='Menu' value='-1'>
					<embed src="<%=FileUploadPath & rs3("picname")%>" width="110" height="80" border=0  pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash'></embed>
					</object>
					<%end if%>

                    <%end if%>
                    </a><br>
                    <br>
                    <a class=middle href="ReadNews.asp?NewsID=<%=rs3("NewsID")%>" target=_blank title="<%=htmlencode4(rs3("title"))%>"> 
                    <%=CutStr(rs3("title"),18)%> </a> </div></td>
<%rs3.movenext
i = i + 1
end if
%>
                <!--图片换行显示3-->
                <td width=25% align=center valign="top" style="table-layout:fixed; word-break:break-all"> 
<% if not rs3.EOF then %>
                  <div align="center"> 
                    <%
Content=htmlencode4(rs3("Content"))
content=replace(content,"[---分页---]","")
%>
                    <a class=middle href="ReadNews.asp?NewsID=<%=rs3("NewsID")%>" target=_blank title="<%=CutStr(nohtml(Content),150)%>"> 
                    <%if rs3("picname")=("") then%>
                    <img  src="IMAGES/flashorno.gif" width="110" height="80" border=1 style=border-color:#000000 align=top> 
                    <%else%>
                    <%if fileext="jpg" or fileext="bmp" or fileext="png" or fileext="gif" then%>
                    <img  src="<%=FileUploadPath & rs3("picname")%>" width="110" height="80" border=1 style=border-color:#000000 align=top> 
                    <%end if%>

                    <%if fileext="swf" then%>
					<object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000'  codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0' width="110" height="80" border=0 >
					<param name=movie value="<%=FileUploadPath & rs3("picname")%>">
					<param name=quality value=high>
					<param name='Play' value='-1'>
					<param name='Loop' value='0'>
					<param name='Menu' value='-1'>
					<embed src="<%=FileUploadPath & rs3("picname")%>" width="110" height="80" border=0  pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash'></embed>
					</object>
					<%end if%>

                    <%end if%>
                    </a><br>
                    <br>
                    <a class=middle href="ReadNews.asp?NewsID=<%=rs3("NewsID")%>" target=_blank title="<%=htmlencode4(rs3("title"))%>"> 
                    <%=CutStr(rs3("title"),18)%> </a> </div></td>
<%rs3.movenext
i = i + 1
end if
%>
                <!--图片换行显示4-->
                <td width=25% align=center valign="top" style="table-layout:fixed; word-break:break-all"> 
<% if not rs3.EOF then %>
                  <div align="center"> 
                   <%
Content=htmlencode4(rs3("Content"))
content=replace(content,"[---分页---]","")
%>
                    <a class=middle href="ReadNews.asp?NewsID=<%=rs3("NewsID")%>" target=_blank title="<%=CutStr(nohtml(Content),150)%>"> 
                    <%if rs3("picname")=("") then%>
                    <img  src="IMAGES/flashorno.gif" width="110" height="80" border=1 style=border-color:#000000 align=top> 
                    <%else%>
                    <%if fileext="jpg" or fileext="bmp" or fileext="png" or fileext="gif" then%>
                    <img  src="<%=FileUploadPath & rs3("picname")%>" width="110" height="80" border=1 style=border-color:#000000 align=top> 
                    <%end if%>

                    <%if fileext="swf" then%>
					<object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000'  codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0' width="110" height="80" border=0 >
					<param name=movie value="<%=FileUploadPath & rs3("picname")%>">
					<param name=quality value=high>
					<param name='Play' value='-1'>
					<param name='Loop' value='0'>
					<param name='Menu' value='-1'>
					<embed src="<%=FileUploadPath & rs3("picname")%>" width="110" height="80" border=0  pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash'></embed>
					</object>
					<%end if%>

                    <%end if%>
                    </a><br>
                    <br>
                    <a class=middle href="ReadNews.asp?NewsID=<%=rs3("NewsID")%>" target=_blank title="<%=htmlencode4(rs3("title"))%>"> 
                    <%=CutStr(rs3("title"),18)%> </a> </div></td>
<%rs3.movenext
i = i + 1
end if
%>
              </tr>
              <!--图片换行结束-->
<%loop%>
            </table></tr>
        
        <tr> 
          <td> <TABLE cellSpacing=0 cellPadding=3 width="100%" align=center border=0 >
              <tr> 
                <td width="100%" align=center valign="middle">共 <%=total%> 
                        条，当前第 <%=Mypage%>/<%=Maxpages%> 页，每页 <%=MyPageSize%> 条 
                        <%
url="bigclass.asp?typeid=" & typeid & "&BigClassid=" & request_BigClassid									
PageNextSize=int((MyPage-1)/PageShowSize)+1
Pagetpage=int((total-1)/rs3.PageSize)+1

if PageNextSize >1 then
PagePrev=PageShowSize*(PageNextSize-1)
Response.write "<a class=black href='" & Url & "&page=" & PagePrev & "' title='上" & PageShowSize & "页'>上一翻页</a> "
Response.write "<a class=black href='" & Url & "&page=1' title='第1页'>页首</a> "
end if
if MyPage-1 > 0 then
Prev_Page = MyPage - 1
Response.write "<a class=black href='" & Url & "&page=" & Prev_Page & "' title='第" & Prev_Page & "页'>上一页</a> "
end if

if Maxpages>=PageNextSize*PageShowSize then
PageSizeShow = PageShowSize
Else
PageSizeShow = Maxpages-PageShowSize*(PageNextSize-1)
End if
If PageSizeShow < 1 Then PageSizeShow = 1
for PageCounterSize=1 to PageSizeShow
PageLink = (PageCounterSize+PageNextSize*PageShowSize)-PageShowSize
if PageLink <> MyPage Then
Response.write "<a class=black href='" & Url & "&page=" & PageLink & "'>[" & PageLink & "]</a> "
else
Response.Write "<B>["& PageLink &"]</B> "
end if
If PageLink = MaxPages Then Exit for
Next

if Mypage+1 <=Pagetpage  then
Next_Page = MyPage + 1
Response.write "<a class=black href='" & Url & "&page=" & Next_Page & "' title='第" & Next_Page & "页'>下一页</A>"
end if

if MaxPages > PageShowSize*PageNextSize then
PageNext = PageShowSize * PageNextSize + 1
Response.write " <A class=black href='" & Url & "&page=" & Pagetpage & "' title='第"& Pagetpage &"页'>页尾</A>"
Response.write " <a class=black href='" & Url & "&page=" & PageNext & "' title='下" & PageShowSize & "页'>下一翻页</a>"
End if
else
Response.write "<tr><td align=center>&nbsp;本类暂无信息</td></tr>"
				
End If
Rs3.close
set rs3=nothing
%>
                      </td>
              </tr>
            </table>
            </td>
        </tr>
      </table>
      <% case "2" %>
      <!--新闻模版-->
      <%
set rs3=server.CreateObject("ADODB.RecordSet")
%>
      <table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="100%" id="AutoNumber4">
        <tr> 
          <td height="25" background="IMAGES/menu-l-zj.gif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font class=body><%=BigClassName%></font></td>
        </tr>
        <%if uselevel=1 then
if Request.cookies(Forcast_SN)("key")="super" or Request.cookies(Forcast_SN)("key")="typemaster" or Request.cookies(Forcast_SN)("key")="bigmaster" or Request.cookies(Forcast_SN)("key")="smallmaster" or Request.cookies(Forcast_SN)("key")="check" then
rs3.Source="select  * from "& db_News_Table &" where (BigClassid=" & request_BigClassid & " and checkked=1) order by newsid DESC"
end if
if Request.cookies(Forcast_SN)("key")="" then
rs3.Source="select  * from "& db_News_Table &" where (BigClassid=" & request_BigClassid & " and checkked=1 ) order by newsid DESC"
end if
if Request.cookies(Forcast_SN)("key")="selfreg" then
	if Request.cookies(Forcast_SN)("reglevel")=3 then
rs3.Source="select  * from "& db_News_Table &" where (BigClassid=" & request_BigClassid & " and checkked=1 ) order by newsid DESC"
	end if
	if Request.cookies(Forcast_SN)("reglevel")=2 then
rs3.Source="select  * from "& db_News_Table &" where (BigClassid=" & request_BigClassid & " and checkked=1 ) order by newsid DESC"
	end if
	if Request.cookies(Forcast_SN)("reglevel")=1 then
rs3.Source="select  * from "& db_News_Table &" where (BigClassid=" & request_BigClassid & " and checkked=1 ) order by newsid DESC"
	end if
end if
else
rs3.Source="select  * from "& db_News_Table &" where (BigClassid=" & request_BigClassid & " and checkked=1) order by newsid DESC"
end if
rs3.Open rs3.Source,conn,1,1

if not rs3.EOF then
Rs3.PageSize     = MyPageSize
MaxPages         = Rs3.PageCount
Rs3.absolutepage = MyPage
total            = Rs3.RecordCount

i = 0
do until rs3.Eof or i = rs3.PageSize

newsurl="ReadNews.asp?NewsID=" & rs3("NewsID")
newswwwurl=rs3("titleface")
datetime="<font class=middle>(" & year(rs3("UpdateTime"))  &"年"& Month(rs3("UpdateTime"))  &"月"& Day(rs3("UpdateTime")) &"日)</font>"

if rs3("picname")<>"" then
img="<img src='images/news_img.gif' border='0'>"
else
img=""
end if

title=trim(rs3("title"))
title=replace(title,"<br>","")
%>
        <tr> 
          <td height=1><div align="center"> 
              <table width="98%" border="0" cellpadding="0" cellspacing="0">
                <tr> 
                  <td width="60%">&nbsp;&nbsp;<img src="images/006.gif" width="8" height="10" border="0"> 
                   <%=img%> <a class=middle href="<%if rs3("titleface")="无" then %><%=newsurl%><% else %> <%=newswwwurl%><%end if%>" target="_blank" title="<%=htmlencode4(title)%>"><font color="<%=rs3("titlecolor")%>"><%=CutStr(title,46)%></font></a> 
                    <!--标题后评论提示-->
                    <% if rs3("titlesize")>=1 then %>
                    <A class=middle HREF="<%=path%>Review.asp?NewsID=<%=rs3("NewsID")%>" target="_blank" ><font color=red><b>评</b></font></A> 
                    <%end if %>
                    <!--标题后评论提示-->
                  </td>
                  <td width="7%" align="right"> 
                    <%if showauthor="1" then%>
                    <%=rs3("author")%> 
                    <%end if%>
                  </td>
                  <td width="20%" align="right"> 
                    <%if showtime="1" then%>
                    <%=datetime%> 
                    <%end if%>
                  </td>
                  <td width="5%" align="right"> 
                    <%if showclick="1" then%>
                    <font color=#666666><%=rs3("click")%></font> 
                    <%end if%>
                  </td>
                  <td width="8%" align="right"> 
                    <%if year(rs3("updatetime"))=year(date()) and month(rs3("updatetime"))=month(date()) and day(rs3("updatetime"))=day(date()) then%>
                    <img src="images/new.gif"> 
                    <%end if%>
                    <%if rs3("goodnews")="1" then%>
                    <img src="images/g.gif" > 
                    <%end if%>
              </table>

<%
rs3.MoveNext
i = i + 1
loop
%>

            </div></td>
        </tr>
<tr>
<td>
        <TABLE cellSpacing=0 cellPadding=3 width="100%" align=center border=0 >
              <tr> 
                <td width="100%" align=center valign="middle">共 <%=total%> 
                        条，当前第 <%=Mypage%>/<%=Maxpages%> 页，每页 <%=MyPageSize%> 条 
                        <%
url="bigclass.asp?typeid=" & typeid & "&BigClassid=" & request_BigClassid									
PageNextSize=int((MyPage-1)/PageShowSize)+1
Pagetpage=int((total-1)/rs3.PageSize)+1

if PageNextSize >1 then
PagePrev=PageShowSize*(PageNextSize-1)
Response.write "<a class=black href='" & Url & "&page=" & PagePrev & "' title='上" & PageShowSize & "页'>上一翻页</a> "
Response.write "<a class=black href='" & Url & "&page=1' title='第1页'>页首</a> "
end if
if MyPage-1 > 0 then
Prev_Page = MyPage - 1
Response.write "<a class=black href='" & Url & "&page=" & Prev_Page & "' title='第" & Prev_Page & "页'>上一页</a> "
end if

if Maxpages>=PageNextSize*PageShowSize then
PageSizeShow = PageShowSize
Else
PageSizeShow = Maxpages-PageShowSize*(PageNextSize-1)
End if
If PageSizeShow < 1 Then PageSizeShow = 1
for PageCounterSize=1 to PageSizeShow
PageLink = (PageCounterSize+PageNextSize*PageShowSize)-PageShowSize
if PageLink <> MyPage Then
Response.write "<a class=black href='" & Url & "&page=" & PageLink & "'>[" & PageLink & "]</a> "
else
Response.Write "<B>["& PageLink &"]</B> "
end if
If PageLink = MaxPages Then Exit for
Next

if Mypage+1 <=Pagetpage  then
Next_Page = MyPage + 1
Response.write "<a class=black href='" & Url & "&page=" & Next_Page & "' title='第" & Next_Page & "页'>下一页</A>"
end if

if MaxPages > PageShowSize*PageNextSize then
PageNext = PageShowSize * PageNextSize + 1
Response.write " <A class=black href='" & Url & "&page=" & Pagetpage & "' title='第"& Pagetpage &"页'>页尾</A>"
Response.write " <a class=black href='" & Url & "&page=" & PageNext & "' title='下" & PageShowSize & "页'>下一翻页</a>"
End if
else
Response.write "<tr><td align=center>&nbsp;本类暂无信息</td></tr>"
				
End If
Rs3.close
set rs3=nothing
%>
                      </td>
              </tr>
            </table>
</td>
     </tr>
      </table>

      <% case "3" %>
      <!--网址模版-->
      <%
set rs3=server.CreateObject("ADODB.RecordSet")
%>
      <table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="100%" id="AutoNumber4">
        <tr> 
          <td height="25" background="IMAGES/menu-l-zj.gif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font class=body><%=BigClassName%></font></td>
        </tr>
        <%if uselevel=1 then
if Request.cookies(Forcast_SN)("key")="super" or Request.cookies(Forcast_SN)("key")="typemaster" or Request.cookies(Forcast_SN)("key")="bigmaster" or Request.cookies(Forcast_SN)("key")="smallmaster" or Request.cookies(Forcast_SN)("key")="check" then
rs3.Source="select  * from "& db_News_Table &" where (BigClassid=" & request_BigClassid & " and checkked=1) order by newsid DESC"
end if
if Request.cookies(Forcast_SN)("key")="" then
rs3.Source="select  * from "& db_News_Table &" where (BigClassid=" & request_BigClassid & " and checkked=1 ) order by newsid DESC"
end if
if Request.cookies(Forcast_SN)("key")="selfreg" then
	if Request.cookies(Forcast_SN)("reglevel")=3 then
rs3.Source="select  * from "& db_News_Table &" where (BigClassid=" & request_BigClassid & " and checkked=1 ) order by newsid DESC"
	end if
	if Request.cookies(Forcast_SN)("reglevel")=2 then
rs3.Source="select  * from "& db_News_Table &" where (BigClassid=" & request_BigClassid & " and checkked=1 ) order by newsid DESC"
	end if
	if Request.cookies(Forcast_SN)("reglevel")=1 then
rs3.Source="select  * from "& db_News_Table &" where (BigClassid=" & request_BigClassid & " and checkked=1 ) order by newsid DESC"
	end if
end if
else
rs3.Source="select  * from "& db_News_Table &" where (BigClassid=" & request_BigClassid & " and checkked=1) order by newsid DESC"
end if
rs3.Open rs3.Source,conn,1,1

if not rs3.EOF then
Rs3.PageSize     = MyPageSize
MaxPages         = Rs3.PageCount
Rs3.absolutepage = MyPage
total            = Rs3.RecordCount
i = 0
%>
<td>
<table width="98%" border="1" cellspacing="0" cellpadding="3" bordercolorlight=#cccccc  bordercolordark=#ffffff align="center">
              <!--网址换行显示1-->
<% do until rs3.Eof or i = rs3.PageSize %>
<tr>
<%
Content=htmlencode4(rs3("Content"))
%>
<% if not rs3.EOF then %>
<td width=50% align=center valign="middle">
                <div align="center"> <a class=middle href="<%=rs3("Original")%>" target=_blank title="<%=CutStr(nohtml(rs3("Content")),80)%>"> 
                  <%=CutStr(rs3("title"),30)%> </a></div></td>
<%rs3.movenext
i = i + 1
end if
%>

              <!--网址换行显示2-->
<% if not rs3.EOF then %>
<td width=50% align=center valign="middle"> 

<div align="center"> <a class=middle href="<%=rs3("Original")%>" target=_blank title="<%=CutStr(nohtml(rs3("Content")),80)%>"> 
                  <%=CutStr(rs3("title"),30)%> </a> </div></td>
<%rs3.movenext
i = i + 1
end if
%>
</tr>

            <!--网址换行结束-->
<%loop%>

          </table></td>

        <tr> 
          <td> <TABLE cellSpacing=0 cellPadding=3 width="100%" align=center border=0 >
              <tr> 
                <td width="100%" align=center valign="middle">共 <%=total%> 
                        条，当前第 <%=Mypage%>/<%=Maxpages%> 页，每页 <%=MyPageSize%> 条 
                        <%
url="bigclass.asp?typeid=" & typeid & "&BigClassid=" & request_BigClassid									
PageNextSize=int((MyPage-1)/PageShowSize)+1
Pagetpage=int((total-1)/rs3.PageSize)+1

if PageNextSize >1 then
PagePrev=PageShowSize*(PageNextSize-1)
Response.write "<a class=black href='" & Url & "&page=" & PagePrev & "' title='上" & PageShowSize & "页'>上一翻页</a> "
Response.write "<a class=black href='" & Url & "&page=1' title='第1页'>页首</a> "
end if
if MyPage-1 > 0 then
Prev_Page = MyPage - 1
Response.write "<a class=black href='" & Url & "&page=" & Prev_Page & "' title='第" & Prev_Page & "页'>上一页</a> "
end if

if Maxpages>=PageNextSize*PageShowSize then
PageSizeShow = PageShowSize
Else
PageSizeShow = Maxpages-PageShowSize*(PageNextSize-1)
End if
If PageSizeShow < 1 Then PageSizeShow = 1
for PageCounterSize=1 to PageSizeShow
PageLink = (PageCounterSize+PageNextSize*PageShowSize)-PageShowSize
if PageLink <> MyPage Then
Response.write "<a class=black href='" & Url & "&page=" & PageLink & "'>[" & PageLink & "]</a> "
else
Response.Write "<B>["& PageLink &"]</B> "
end if
If PageLink = MaxPages Then Exit for
Next

if Mypage+1 <=Pagetpage  then
Next_Page = MyPage + 1
Response.write "<a class=black href='" & Url & "&page=" & Next_Page & "' title='第" & Next_Page & "页'>下一页</A>"
end if

if MaxPages > PageShowSize*PageNextSize then
PageNext = PageShowSize * PageNextSize + 1
Response.write " <A class=black href='" & Url & "&page=" & Pagetpage & "' title='第"& Pagetpage &"页'>页尾</A>"
Response.write " <a class=black href='" & Url & "&page=" & PageNext & "' title='下" & PageShowSize & "页'>下一翻页</a>"
End if
else
Response.write "<tr><td align=center>&nbsp;本类暂无信息</td></tr>"
				
End If

Rs3.close
set rs3=nothing
%>
                      </td>
              </tr>
            </table></tr></table>
      <% case "4" %>
      <!--软件模版-->
<%
set rs3=server.CreateObject("ADODB.RecordSet")
%>
      <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" >
        <tr> 
          <td> <table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="100%" id="AutoNumber4">
              <tr> 
                <td height="25" background="IMAGES/menu-l-zj.gif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font class=body><%=BigClassName%></font> 
                </td>
              </tr>
            </table>
            <%if uselevel=1 then
if Request.cookies(Forcast_SN)("key")="super" or Request.cookies(Forcast_SN)("key")="typemaster" or Request.cookies(Forcast_SN)("key")="bigmaster" or Request.cookies(Forcast_SN)("key")="smallmaster" or Request.cookies(Forcast_SN)("key")="check" then
rs3.Source="select  * from "& db_News_Table &" where (BigClassid=" & request_BigClassid & " and checkked=1) order by newsid DESC"
end if
if Request.cookies(Forcast_SN)("key")="" then
rs3.Source="select  * from "& db_News_Table &" where (BigClassid=" & request_BigClassid & " and checkked=1 ) order by newsid DESC"
end if
if Request.cookies(Forcast_SN)("key")="selfreg" then
	if Request.cookies(Forcast_SN)("reglevel")=3 then
rs3.Source="select  * from "& db_News_Table &" where (BigClassid=" & request_BigClassid & " and checkked=1 ) order by newsid DESC"
	end if
	if Request.cookies(Forcast_SN)("reglevel")=2 then
rs3.Source="select  * from "& db_News_Table &" where (BigClassid=" & request_BigClassid & " and checkked=1 ) order by newsid DESC"
	end if
	if Request.cookies(Forcast_SN)("reglevel")=1 then
rs3.Source="select  * from "& db_News_Table &" where (BigClassid=" & request_BigClassid & " and checkked=1 ) order by newsid DESC"
	end if
end if
else
rs3.Source="select  * from "& db_News_Table &" where (BigClassid=" & request_BigClassid & " and checkked=1) order by newsid DESC"
end if
rs3.Open rs3.Source,conn,1,1

if not rs3.EOF then
Rs3.PageSize     = MyPageSize
MaxPages         = Rs3.PageCount
Rs3.absolutepage = MyPage
total            = Rs3.RecordCount

i = 0
do until rs3.Eof or i = rs3.PageSize

newsurl="ReadNews.asp?NewsID=" & rs3("NewsID")
newswwwurl=rs3("titleface")
datetime="<font class=middle>(" & Month(rs3("UpdateTime"))  &"月"& Day(rs3("UpdateTime")) &"日)</font>"


title=htmlencode4(trim(rs3("title")))
%>

            <div align="center"> 
              <table width="95%" border="0" cellpadding="3" cellspacing="0">
                <tr bgcolor="#EFEFEF"> 

                  <td colspan="2">&nbsp;<img src="images/news_img.gif" width="9" height="9" border="0">&nbsp; <a class=middle href="<%if rs3("titleface")="无" then %><%=newsurl%><% else %> <%=newswwwurl%><%end if%>" target="_blank" title="<%=htmlencode4(title)%>"><strong><font color="<%=rs3("titlecolor")%>"> 
                    <%=CutStr(title,40)%> 
                    </font></strong></a></td>
                  <td width="22%" align="right" bgcolor="#EFEFEF"> 
                    <%if showtime="1" then%>
                    <%=datetime%> 
                    <%end if%>
                    <%if showclick="1" then%>
                    <font color=#666666>[<%=rs3("click")%>]</font> 
                    <%end if%>
                  </td>
                  <td width="13%" align="center" bgcolor="#EFEFEF"><div align="right"> 
                      <%if year(rs3("updatetime"))=year(date()) and month(rs3("updatetime"))=month(date()) and day(rs3("updatetime"))=day(date()) then%>
                      <img src="images/new.gif" border="0"> 
                      <%end if%>
                      <%if rs3("goodnews")="1" then%>
                      <img src="images/g.gif" border="0"> 
                      <%end if%>
                    </div></td>
                </tr>
                <tr> 
<%
fileExt=lcase(getFileExtName(rs3("picname")))
Content=htmlencode4(rs3("Content"))
content=replace(content,"[---分页---]","")%>
                  <td width="16%"><a class=middle href="<%if rs3("titleface")="无" then %><%=newsurl%><% else %> <%=newswwwurl%><%end if%>" target=_blank title="<%=htmlencode4(rs3("title"))%>"> 
                    <%if   rs3("picname")=("")   then%>
                    <img  src="IMAGES/softno.gif" width="65" height="65" border=0 align="left"> 
                    <%else%>
                    <%if fileext="jpg" or fileext="bmp" or fileext="png" or fileext="gif" then%>
                    <img  src="<%=FileUploadPath & rs3("picname")%>" width="65" height="65" border=0 align="left"> 
                    <%end if%>

                    <%if fileext="swf" then%>
					<object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000'  codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0' width="65" height="65" border=0 >
					<param name=movie value="<%=FileUploadPath & rs3("picname")%>">
					<param name=quality value=high>
					<param name='Play' value='-1'>
					<param name='Loop' value='0'>
					<param name='Menu' value='-1'>
					<embed src="<%=FileUploadPath & rs3("picname")%>" width="6" 5height="65" border=0  pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash'></embed>
					</object>
					<%end if%>

                    <%end if%>
                    </a></td>
                  <td colspan="3" valign="top"> <a class=middle href="<%if rs3("titleface")="无" then %><%=newsurl%><% else %> <%=newswwwurl%><%end if%>" target="_blank" title="<%=htmlencode4(title)%>"><%=CutStr(nohtml(rs3("Content")),250)%></a> 
                  </td>
                </tr>
              </table>
              <%
rs3.MoveNext
i = i + 1
loop
%>

              <TABLE cellSpacing=0 cellPadding=3 width="100%" align=center border=0 >
              <tr> 
                <td width="100%" align=center valign="middle">共 <%=total%> 
                        条，当前第 <%=Mypage%>/<%=Maxpages%> 页，每页 <%=MyPageSize%> 条 
                        <%
url="bigclass.asp?typeid=" & typeid & "&BigClassid=" & request_BigClassid									
PageNextSize=int((MyPage-1)/PageShowSize)+1
Pagetpage=int((total-1)/rs3.PageSize)+1

if PageNextSize >1 then
PagePrev=PageShowSize*(PageNextSize-1)
Response.write "<a class=black href='" & Url & "&page=" & PagePrev & "' title='上" & PageShowSize & "页'>上一翻页</a> "
Response.write "<a class=black href='" & Url & "&page=1' title='第1页'>页首</a> "
end if
if MyPage-1 > 0 then
Prev_Page = MyPage - 1
Response.write "<a class=black href='" & Url & "&page=" & Prev_Page & "' title='第" & Prev_Page & "页'>上一页</a> "
end if

if Maxpages>=PageNextSize*PageShowSize then
PageSizeShow = PageShowSize
Else
PageSizeShow = Maxpages-PageShowSize*(PageNextSize-1)
End if
If PageSizeShow < 1 Then PageSizeShow = 1
for PageCounterSize=1 to PageSizeShow
PageLink = (PageCounterSize+PageNextSize*PageShowSize)-PageShowSize
if PageLink <> MyPage Then
Response.write "<a class=black href='" & Url & "&page=" & PageLink & "'>[" & PageLink & "]</a> "
else
Response.Write "<B>["& PageLink &"]</B> "
end if
If PageLink = MaxPages Then Exit for
Next

if Mypage+1 <=Pagetpage  then
Next_Page = MyPage + 1
Response.write "<a class=black href='" & Url & "&page=" & Next_Page & "' title='第" & Next_Page & "页'>下一页</A>"
end if

if MaxPages > PageShowSize*PageNextSize then
PageNext = PageShowSize * PageNextSize + 1
Response.write " <A class=black href='" & Url & "&page=" & Pagetpage & "' title='第"& Pagetpage &"页'>页尾</A>"
Response.write " <a class=black href='" & Url & "&page=" & PageNext & "' title='下" & PageShowSize & "页'>下一翻页</a>"
End if
else
Response.write "<tr><td align=center>&nbsp;本类暂无信息</td></tr>"
				
End If
Rs3.close
set rs3=nothing
%>
                      </td>
              </tr>
            </table>
                          
              <!--模版结束-->
            </div></td>
        </tr>
      </table><% end select%>
</td>
</tr>
  <tr valign="top">
    <td height="20" background="IMAGES/menu-d-t.gif" bgcolor="#FFFFFF">&nbsp;</td>
    <td>&nbsp;</td>
    <td height="20" background="IMAGES/menu-l-t.gif" class="daohang">&nbsp;</td>
  </tr>
</table>
<!--#include file="bottom.asp"-->
</body>
</html>
