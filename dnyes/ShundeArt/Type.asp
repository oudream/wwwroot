<%@ Language=VBScript%>
<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<!--#include file="function.asp"-->
<!--#include file="char.inc"-->
<%
typeid=checkstr(request("typeid"))
if typeid="" then
	Response.Write "<script>alert('未指定参数');history.back()</script>"
	response.end
else
	if not IsNumeric(typeid) then
		response.write "<script>alert('非法参数');history.back()</script>"
		response.end
	else
		dim ttypename
		set rs5=server.CreateObject("ADODB.RecordSet")
		rs5.Source="select * from "& db_Type_Table &" where typeid=" & typeid &" order by typeorder"
		rs5.Open rs5.Source,conn,1,1
		ttypename=rs5("typename")
		rs5.Close
		set rs5=nothing
		%>
		<%
		dim typeid
		typeid=checkstr(request("typeid"))
		set rs=server.CreateObject("ADODB.RecordSet")
		rs.Source="select * from "& db_BigClass_Table &" where typeid="& typeid &" order by bigclassorder"
		rs.Open rs.Source,conn,1,1
		i=1
		Dim ArrayBigClassID(10000),ArrayBigClassName(10000),ArrayBigClassView(10000)
		if not rs.EOF then
			rseof=1
		end if
	end if
	while not rs.EOF
		abcount=rs.RecordCount
		BigClassView=rs("BigClassView")
		BigClassID=rs("BigClassID")
		
		BigClassName=trim(rs("BigClassName"))
		BigClasszs=rs("BigClasszs")
		
		ArrayBigClassView(i)=BigClassView
		ArrayBigClassID(i)=BigClassID
		ArrayBigClassName(i)=BigClassName

		i=i+1
		
		rs.MoveNext
	wend
	rs.close
	%>

<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=ttypeName%>__<%=jjgn%></title>
<LINK href=news.css rel=stylesheet>
<script language="JavaScript1.2">
<!--
function makevisible(cur,which){
if (which==0)
cur.filters.alpha.opacity=100
else
cur.filters.alpha.opacity=20
}

function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}
//-->
</script>
</head>

<body topmargin="0" marginheight="0">
<!--#include file="top.asp"-->
<%
dim typename
typeid=checkstr(request("typeid"))
set rs=server.CreateObject("ADODB.RecordSet")
rs.Source="select * from "& db_Type_Table &" where typeid="& typeid &" order by typeorder"
rs.Open rs.Source,conn,1,1
mode=rs("mode")
typename=rs("typename")
rs.close
set rs=nothing%>
<table width="760" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
	<tr valign="top"> 
		<td width="180" background="IMAGES/menu-d.gif"> 
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr bordercolor="#999999"> 
					<td height="25" align="center" valign="middle" bordercolor="#999999" background="IMAGES/menu-d-fl.gif"><font color="#000000">分类栏目</font></td>
				</tr>
				<tr> 
					<td height="490" valign="top" background="IMAGES/menu-d.gif"> 
						<table width="100%" border="0" cellspacing="0" cellpadding="3">
							<tr> 
								<td height="25" align="center" valign="middle" background="IMAGES/menu-d.gif"> 
									<table width="92%" border="0" cellpadding="0" cellspacing="0">
										<tr> 
											<td> 
												<%
												'dim mode
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
																			<font class=m_tittle>&nbsp;</font>
																			<a class=middle href="BigClass.asp?typeid=<%=typeid%>&bigclassid=<%=Rs("bigclassid")%>">
																				<font color="#000000"><%=Rs("bigclassname")%></font>
																			</a>&nbsp;
																	</td>
																</tr>
																			<%set nRs=conn.execute("SELECT * FROM "& db_SmallClass_Table &" where bigclassid="&cstr(rs("bigclassid"))&" order by smallclassorder")%>
																			<%do while not nrs.eof%>
																<tr> 
																	<td width=100% align=center valign="middle"> 
																				<%if not nRs.eof then%>
																		<a class=class href="SmallClass.asp?typeid=<%=typeid%>&BigClassID=<%=nrs("bigclassid")%>&smallclassid=<%=nrs("smallclassid")%>"><%=nRs("smallclassname")%></a>
																	</td>
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
													set Rs=conn.execute("SELECT * FROM "& db_BigClass_Table &" where typeid="&typeid&" order by bigclassorder") 
													do while not rs.eof
														%>
														<table width="100%" border="0" cellpadding="3" cellspacing="0" id="AutoNumber2" style="border-collapse: collapse">
															<tr> 
																<td height=22 colspan="2"> <li></li>
																	<%if not Rs.eof then%>
																		<a class=middle href="BigClass.asp?typeid=<%=typeid%>&bigclassid=<%=Rs("bigclassid")%>"><font color="#000000"><%=Rs("bigclassname")%></font></a>
																</td>
															</tr>
																		<%set nRs=conn.execute("SELECT * FROM "& db_SmallClass_Table &" where bigclassid="&cstr(rs("bigclassid"))&" order by smallclassorder")%>
																		<%do while not nrs.eof%>
															<tr> 
																<td width=50% align=center valign="middle"> 
																			<%if not nRs.eof then%>
																				<a class=class href="SmallClass.asp?typeid=<%=typeid%>&BigClassID=<%=nrs("bigclassid")%>&smallclassid=<%=nrs("smallclassid")%>"><%=nRs("smallclassname")%></a>
																</td>
																				<%nRs.movenext   
																			end if%>
																<td width=50% align=center valign="middle"> 
																			<%if not nRs.eof then%>
																				<a class=class href="SmallClass.asp?typeid=<%=typeid%>&BigClassID=<%=nrs("bigclassid")%>&smallclassid=<%=nrs("smallclassid")%>"><%=nRs("smallclassname")%></a>
																</td>
																				<%nRs.movenext
																			end if%>
															</tr>
																		<%loop
																		nRs.Close
																		set nRs=nothing
																		%>
														</table>
																		<%Rs.movenext   
																	end if
													loop  
													rs.close
													%>
												<%end if%>
											</td>
										</tr>
									</table>
								</td>
							</tr>
						</table>
						<table width="102%" height="25" border="0" cellpadding="0" cellspacing="0" id="AutoNumber4" style="border-collapse: collapse">
							<tr> 
								<td height="25" align=center background="IMAGES/menu-d-pic.gif"><font color="#000000">最新图文</font></td>
							</tr>
							<tr> 
								<td width="100%" height="20" align=center background="IMAGES/menu-d.gif" bgcolor="#FFFFFF"><br> 
									<%
									set rs3=server.CreateObject("ADODB.RecordSet")
									if uselevel=1 then
										if Request.cookies(Forcast_SN)("key")="super" or Request.cookies(Forcast_SN)("key")="typemaster" or Request.cookies(Forcast_SN)("key")="bigmaster" or Request.cookies(Forcast_SN)("key")="smallmaster" or Request.cookies(Forcast_SN)("key")="check" then
											rs3.Source ="select top " & top_img & " * from "& db_News_Table &" where picnews=1 and checkked=1 and typeid="&typeid&" and picname is not null order by NewsID DESC"
										end if
										if Request.cookies(Forcast_SN)("key")="" then
											rs3.Source ="select top " & top_img & " * from "& db_News_Table &" where picnews=1 and checkked=1  and typeid="&typeid&" and picname is not null order by NewsID DESC"
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
											<td style="WORD-WRAP: break-word">
												<a class=middle href="ReadNews.asp?NewsID=<%=rs3("NewsID")%>" target=_blank title="<%=htmlencode4(rs3("title"))%>"> 
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
													<embed src="<%=FileUploadPath & rs3("picname")%>" width="65" height="65" border=0 align="left" pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash'></embed>
													</object>
											<%end if%>
													<%=rs3("title")%>
												</a>
											</td>
										</tr>
									</table>
											<%
											rs3.MoveNext
										wend
									else
										Response.Write "<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"" align=""center"" style=""TABLE-LAYOUT: fixed"" background=""IMAGES/menu-d.gif""><tr><td width=100% align=center height=18 background=""IMAGES/menu-d.gif"">暂无</td></tr></table>"
									end if
									rs3.close
									set rs3=nothing
									%>
								</td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
		</td>
		<td width="6" bgcolor="#FFFFFF"></td>
		<td background="IMAGES/menu-l.gif"> 
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
				<tr> 
					<td height="25" background="IMAGES/menu-l-m.gif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font class=m_tittle>&nbsp;</font>栏目导航&nbsp;&nbsp;<a href="./" class="daohang">网站首页</a>&gt;&gt;<%=typename%></td>
				</tr>
				<tr> 
					<td background="IMAGES/menu-l.gif">
						<div align="center"> 
							<script language=javascript src=./zongg/ad.asp?i=13></script>
						</div>
								<%end if %>
                 </td>
				</tr>
			</table>
							
        
      <!--模版开始-->
      <% select case mode%>
      <% case 1 %>
      <!--图片模版-->
      <%
										if rseof=1 then
											for i=1 to abcount
												BigClassID=ArrayBigClassID(i)
												BigClassName=ArrayBigClassName(i)
												BigClassview=ArrayBigClassview(i)
												if ArrayBigClassView(i)=1 then 
													%>
      <table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber5">
        <tr> 
          <td height="25" background="IMAGES/menu-l-zj.gif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font class=body><%=BigClassName%></font></td>
        </tr>
        <%set rs3=server.CreateObject("ADODB.RecordSet")
													if uselevel=1 then
														if Request.cookies(Forcast_SN)("key")="super" or Request.cookies(Forcast_SN)("key")="typemaster" or Request.cookies(Forcast_SN)("key")="bigmaster" or Request.cookies(Forcast_SN)("key")="smallmaster" or Request.cookies(Forcast_SN)("key")="check" then
															rs3.Source="select top "& top_news &" * from "& db_News_Table &" where (BigClassid="& BigClassid &" and checkked=1) order by newsid DESC"
														end if
														if Request.cookies(Forcast_SN)("key")="" then
															rs3.Source="select top "& top_news &" * from "& db_News_Table &" where (BigClassid="& BigClassid &" and checkked=1 ) order by newsid DESC"
														end if
														if Request.cookies(Forcast_SN)("key")="selfreg" then
															if Request.cookies(Forcast_SN)("reglevel")=3 then
																rs3.Source="select top "& top_news &" * from "& db_News_Table &" where (BigClassid="& BigClassid &" and checkked=1 ) order by newsid DESC"
															end if
															if Request.cookies(Forcast_SN)("reglevel")=2 then
																rs3.Source="select top "& top_news &" * from "& db_News_Table &" where (BigClassid="& BigClassid &" and checkked=1 ) order by newsid DESC"
															end if
															if Request.cookies(Forcast_SN)("reglevel")=1 then
																rs3.Source="select top "& top_news &" * from "& db_News_Table &" where (BigClassid="& BigClassid &" and checkked=1 ) order by newsid DESC"
															end if
														end if
													else
														rs3.Source="select top "& top_news &" * from "& db_News_Table &" where (BigClassid="& BigClassid &" and checkked=1) order by newsid DESC"
													end if
													rs3.Open rs3.Source,conn,1,1
													%>
        <tr> 
          <td><table width="90%" border="0" cellspacing="0" cellpadding="3" align="center">
              <!--图片换行显示1-->
              <%do while not rs3.EOF%>
              <tr> 
                <%
				fileExt=lcase(getFileExtName(rs3("picname")))
				Content=htmlencode4(rs3("Content"))
				content=replace(content,"[---分页---]","")
				%>
                <td width=25% align=center valign="top" style="table-layout:fixed; word-break:break-all"> 
                  <%if not rs3.EOF then%>
                  <div align="center"> 
                    <% content=htmlencode4(rs3("Content"))
					content=replace(content,"[---分页---]","")%>
                    <a class=middle href="ReadNews.asp?NewsID=<%=rs3("NewsID")%>" target=_blank title="<%=CutStr(nohtml(Content),150)%>..."> 
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
                    </a> <br>
                    <br>
                    <a class=middle href="ReadNews.asp?NewsID=<%=rs3("NewsID")%>" target=_blank title="<%=rs3("title")%>"><%=CutStr(rs3("title"),18)%></a> 
                  </div></td>
                <%rs3.movenext   
														end if %>
                <!--图片换行显示2-->
				<%
				fileExt=lcase(getFileExtName(rs3("picname")))
				Content=htmlencode4(rs3("Content"))
				content=replace(content,"[---分页---]","")
				%>
                <td width=25% align=center valign="top" style="table-layout:fixed; word-break:break-all"> 
                  <%if not rs3.EOF then%>
                  <div align="center"> 
                    <%
					content=htmlencode4(rs3("Content"))
					content=replace(content,"[---分页---]","")%>
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
                    </a> <br>
                    <br>
                    <a class=middle href="ReadNews.asp?NewsID=<%=rs3("NewsID")%>" target=_blank title="<%=rs3("title")%>"><%=CutStr(rs3("title"),18)%></a> 
                  </div></td>
                <%rs3.movenext   
														end if %>
                <!--图片换行显示3-->
				<%
				fileExt=lcase(getFileExtName(rs3("picname")))
				Content=htmlencode4(rs3("Content"))
				content=replace(content,"[---分页---]","")
				%>
                <td width=25% align=center valign="top" style="table-layout:fixed; word-break:break-all"> 
                  <%if not rs3.EOF then%>
                  <div align="center"> 
                    <%
					content=htmlencode4(rs3("Content"))
					content=replace(content,"[---分页---]","")%>
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
                    </a> <br>
                    <br>
                    <a class=middle href="ReadNews.asp?NewsID=<%=rs3("NewsID")%>" target=_blank title="<%=rs3("title")%>"><%=CutStr(rs3("title"),18)%></a> 
                  </div></td>
                <%rs3.movenext   
														end if %>
                <!--图片换行显示4-->
				<%
				fileExt=lcase(getFileExtName(rs3("picname")))
				Content=htmlencode4(rs3("Content"))
				content=replace(content,"[---分页---]","")
				%>
                <td width=25% align=center valign="top" style="table-layout:fixed; word-break:break-all"> 
                  <%if not rs3.EOF then%>
                  <div align="center"> 
                    <%
															content=rs3("Content")
															content=replace(content,"[---分页---]","")%>
                    <a class=middle href="ReadNews.asp?NewsID=<%=rs3("NewsID")%>" target=_blank title="<%=CutStr(nohtml(Content),150)%>..."> 
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
                    </a> <br>
                    <br>
                    <a class=middle href="ReadNews.asp?NewsID=<%=rs3("NewsID")%>" target=_blank title="<%=rs3("title")%>"><%=CutStr(rs3("title"),18)%></a> 
                  </div></td>
                <%rs3.movenext   
														end if %>
              </tr>
              <!--图片换行结束-->
              <%loop
													rs3.Close
													set rs3=nothing
													%>
            </table>
            <TABLE cellSpacing=0 cellPadding=3 width="100%" align=center border=0>
              <tr> 
                <td width=100% height="19" align=right><a class=class href='BigClass.asp?typeid=<%=typeid%>&BigClassid=<%=BigClassid%>'><img src="images/more.gif" border="0" alt="更多<%=BigClassName%>"></a>&nbsp;&nbsp;</td>
              </tr>
            </table></tr>
      </table>
      <%
												end if
											next
										else 
											Response.Write "<table width=98% border=0 height=25><tr><td>暂无大类</td></tr></table>"
										end if
										%>
      <% case 2 %>
      <!--新闻模版-->
      <%if rseof=1 then
											for i=1 to abcount
												BigClassID=ArrayBigClassID(i)
												BigClassName=ArrayBigClassName(i)
												BigClassview=ArrayBigClassview(i)
												if ArrayBigClassView(i)=1 then 
													%>
      <table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber5">
        <tr> 
          <td height="25" background="IMAGES/menu-l-zj.gif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font class=body><%=BigClassName%></font></td>
        </tr>
        <%set rs3=server.CreateObject("ADODB.RecordSet")
													if uselevel=1 then
														if Request.cookies(Forcast_SN)("key")="super" or Request.cookies(Forcast_SN)("key")="typemaster" or Request.cookies(Forcast_SN)("key")="bigmaster" or Request.cookies(Forcast_SN)("key")="smallmaster" or Request.cookies(Forcast_SN)("key")="check" then
															rs3.Source="select top " & top_news & " * from "& db_News_Table &" where (BigClassid=" & BigClassid &" and checkked=1) order by newsid DESC"
														end if
														if Request.cookies(Forcast_SN)("key")="" then
															rs3.Source="select top " & top_news & " * from "& db_News_Table &" where (BigClassid=" & BigClassid &" and checkked=1 ) order by newsid DESC"
														end if
														if Request.cookies(Forcast_SN)("key")="selfreg" then
															if Request.cookies(Forcast_SN)("reglevel")=3 then
																rs3.Source="select top " & top_news & " * from "& db_News_Table &" where (BigClassid=" & BigClassid &" and checkked=1 ) order by newsid DESC"
															end if
															if Request.cookies(Forcast_SN)("reglevel")=2 then
																rs3.Source="select top " & top_news & " * from "& db_News_Table &" where (BigClassid=" & BigClassid &" and checkked=1 ) order by newsid DESC"
															end if
															if Request.cookies(Forcast_SN)("reglevel")=1 then
																rs3.Source="select top " & top_news & " * from "& db_News_Table &" where (BigClassid=" & BigClassid &" and checkked=1 ) order by newsid DESC"
															end if
														end if
													else
														rs3.Source="select top " & top_news & " * from "& db_News_Table &" where (BigClassid=" & BigClassid &" and checkked=1) order by newsid DESC"
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
															img="<img src='images/news_img.gif' border='0'>"
														else
															img=""
														end if
														title=trim(rs3("title"))
														title=replace(title,"<br>","")
														%>
        <tr> 
          <td> <div align="center"> 
              <table width="98%" border="0" cellpadding="0" cellspacing="0">
                <tr> 
                  <td width="65%">&nbsp;&nbsp;<img src="IMAGES/006.gif" width="8" height="10"> 
                    <%=img%> <a class=middle href="<%if rs3("titleface")="无" then %><%=newsurl%><% else %> <%=newswwwurl%><%end if%>" title="<%=rs3("title")%>" target="_blank"> 
                    <<%=rs3("titletype")%>><font color="<%=rs3("titlecolor")%>"> 
                    <%=CutStr(title,46)%></font></<%=rs3("titletype")%>></a> 
                    <!--标题后评论提示-->
                    <% if rs3("titlesize")>=1 then %>
                    <A class=middle HREF="<%=path%>review.asp?NewsID=<%=rs3("NewsID")%>" target="_blank" ><font color=red><b>评</b></font></A> 
                    <%end if %>
                    <!--标题后评论提示-->
                  </td>
                  <td width="7%" align="right"> 
                    <%if showauthor="1" then%>
                    <%=rs3("Author")%> 
                    <%end if%>
                  </td>
                  <td width="15%" align="left"> 
                    <%if showtime="1" then%>
                    <%=datetime%> 
                    <%end if%>
                    &nbsp; </td>
                  <td width="5%" align="center"> 
                    <%if showclick="1" then%>
                    <font color=#666666><%=rs3("click")%></font> 
                    <%end if%>
                  </td>
                  <td width="8%" align="left">
                    <%if showtime="1" then%>
                    <%if year(rs3("updatetime"))=year(date()) and month(rs3("updatetime"))=month(date()) and day(rs3("updatetime"))=day(date()) then%>
                    <%end if%>
                    <img src="images/new.gif"> 
                    <%end if%>
                    <%if rs3("goodnews")="1" then%>
                    <img src="images/g.gif" > 
                    <%end if%>
                  </td>
                </tr>
              </table>
            </div></td>
        </tr>
        <%
														rs3.MoveNext
													wend
													%>
        <tr> 
          <td width=100% height="19" align=right><a class=class href='BigClass.asp?typeid=<%=typeid%>&BigClassid=<%=BigClassid%>'><img src="images/more.gif" border="0" alt="更多<%=BigClassName%>"></a>&nbsp;&nbsp;</td>
        </tr>
      </table>
      <%
												end if
											next
											'rs3.close
											'set rs3=nothing
										else 
											Response.Write "<table width=98% border=0 height=25><tr><td>暂无大类</td></tr></table>"
										end if
										%>
      <% case 3 %>
      <!--网址模版-->
      <%
										if rseof=1 then
											for i=1 to abcount
												BigClassID=ArrayBigClassID(i)
												BigClassName=ArrayBigClassName(i)
												BigClassview=ArrayBigClassview(i)
												if ArrayBigClassView(i)=1 then 
													%>
      <table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber5">
        <tr> 
          <td height="25" background="IMAGES/menu-l-zj.gif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font class=body><%=BigClassName%></font></td>
        </tr>
        <%set rs3=server.CreateObject("ADODB.RecordSet")
													if uselevel=1 then
														if Request.cookies(Forcast_SN)("key")="super" or Request.cookies(Forcast_SN)("key")="typemaster" or Request.cookies(Forcast_SN)("key")="bigmaster" or Request.cookies(Forcast_SN)("key")="smallmaster" or Request.cookies(Forcast_SN)("key")="check" then
															rs3.Source="select top " & top_news & " * from "& db_News_Table &" where (BigClassid=" & BigClassid &" and checkked=1) order by newsid DESC"
														end if
														if Request.cookies(Forcast_SN)("key")="" then
															rs3.Source="select top " & top_news & " * from "& db_News_Table &" where (BigClassid=" & BigClassid &" and checkked=1 ) order by newsid DESC"
														end if
														if Request.cookies(Forcast_SN)("key")="selfreg" then
															if Request.cookies(Forcast_SN)("reglevel")=3 then
																rs3.Source="select top " & top_news & " * from "& db_News_Table &" where (BigClassid=" & BigClassid &" and checkked=1 ) order by newsid DESC"
															end if
															if Request.cookies(Forcast_SN)("reglevel")=2 then
																rs3.Source="select top " & top_news & " * from "& db_News_Table &" where (BigClassid=" & BigClassid &" and checkked=1 ) order by newsid DESC"
															end if
															if Request.cookies(Forcast_SN)("reglevel")=1 then
																rs3.Source="select top " & top_news & " * from "& db_News_Table &" where (BigClassid=" & BigClassid &" and checkked=1 ) order by newsid DESC"
															end if
														end if
													else
														rs3.Source="select top " & top_news & " * from "& db_News_Table &" where (BigClassid=" & BigClassid &" and checkked=1) order by newsid DESC"
													end if
													rs3.Open rs3.Source,conn,1,1
													%>
        <tr> 
          <td> <table width="98%" border="1" cellspacing="0" cellpadding="3" bordercolorlight=#cccccc  bordercolordark=#ffffff align="center">
              <!--网址换行显示1-->
              <%do while not rs3.EOF%>
              <tr> 
                <%
				Content=htmlencode4(rs3("Content"))
				%>
                <td width=25% align=center valign="middle"> 
                  <%if not rs3.EOF then%>
                  <div align="center"> <a class=middle href="<%=rs3("Original")%>" target=_blank title="<%=CutStr(nohtml(rs3("Content")),80)%>"><%=CutStr(rs3("title"),30)%></a> 
                  </div></td>
                <%rs3.movenext   
														end if %>
                <!--网址换行显示2-->
                <td width=25% align=center valign="middle"> 
                  <%if not rs3.EOF then%>
                  <div align="center"> <a class=middle href="<%=rs3("Original")%>" target=_blank title="<%=CutStr(nohtml(rs3("Content")),80)%>"><%=CutStr(rs3("title"),30)%></a> 
                  </div></td>
                <%rs3.movenext   
														end if %>
              </tr>
              <!--网址换行结束-->
              <%loop
													rs3.Close
													set rs3=nothing
													%>
            </table></tr>
        <tr> 
          <td> <TABLE cellSpacing=0 cellPadding=3 width="100%" align=center border=0>
              <tr> 
                <td width=100% height="19" align=right><a class=class href='BigClass.asp?typeid=<%=typeid%>&BigClassid=<%=BigClassid%>'><img src="images/more.gif" border="0" alt="更多<%=BigClassName%>"></a>&nbsp;&nbsp;</td>
              </tr>
            </table></tr>
        <%
												end if
											next
										else 
											Response.Write "<table width=98% border=0 height=25><tr><td>暂无大类</td></tr></table>"
										end if
										%>
      </table>
      <% case 4 %>
      <!--软件模版-->
      <%if rseof=1 then
											for i=1 to abcount
												BigClassID=ArrayBigClassID(i)
												BigClassName=ArrayBigClassName(i)
												BigClassview=ArrayBigClassview(i)
												if ArrayBigClassView(i)=1 then 
													%>
      <table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber5">
        <tr> 
          <td height="25" background="IMAGES/menu-l-zj.gif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font class=body><%=BigClassName%></font></td>
        </tr>
        <%set rs3=server.CreateObject("ADODB.RecordSet")
													if uselevel=1 then
														if Request.cookies(Forcast_SN)("key")="super" or Request.cookies(Forcast_SN)("key")="typemaster" or Request.cookies(Forcast_SN)("key")="bigmaster" or Request.cookies(Forcast_SN)("key")="smallmaster" or Request.cookies(Forcast_SN)("key")="check" then
															rs3.Source="select top " & top_news & " * from "& db_News_Table &" where (BigClassid=" & BigClassid &" and checkked=1) order by newsid DESC"
														end if
														if Request.cookies(Forcast_SN)("key")="" then
															rs3.Source="select top " & top_news & " * from "& db_News_Table &" where (BigClassid=" & BigClassid &" and checkked=1 ) order by newsid DESC"
														end if
														if Request.cookies(Forcast_SN)("key")="selfreg" then
															if Request.cookies(Forcast_SN)("reglevel")=3 then
																rs3.Source="select top " & top_news & " * from "& db_News_Table &" where (BigClassid=" & BigClassid &" and checkked=1 ) order by newsid DESC"
															end if
															if Request.cookies(Forcast_SN)("reglevel")=2 then
																rs3.Source="select top " & top_news & " * from "& db_News_Table &" where (BigClassid=" & BigClassid &" and checkked=1 ) order by newsid DESC"
															end if
															if Request.cookies(Forcast_SN)("reglevel")=1 then
																rs3.Source="select top " & top_news & " * from "& db_News_Table &" where (BigClassid=" & BigClassid &" and checkked=1 ) order by newsid DESC"
															end if
														end if
													else
														rs3.Source="select top " & top_news & " * from "& db_News_Table &" where (BigClassid=" & BigClassid &" and checkked=1) order by newsid DESC"
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
															img="<img src='images/news_img.gif' border='0'>"
														else
															img=""
														end if
														title=trim(rs3("title"))
														title=replace(title,"<br>","")
														%>
        <tr> 
		<%
fileExt=lcase(getFileExtName(rs3("picname")))
Content=htmlencode4(rs3("Content"))
content=replace(content,"[---分页---]","")%>
          <td> <div align="center"> 
              <table width="95%" border="0" cellpadding="3" cellspacing="0">
                <tr bgcolor="#EFEFEF"> 
                  <td colspan="2"> &nbsp;<img src="images/news_img.gif" width="9" height="9"> 
                    <a class=middle href="<%if rs3("titleface")="无" then %><%=newsurl%><% else %> <%=newswwwurl%><%end if%>" title="<%=htmlencode4(title)%>" target="_blank"> 
                    <font color="<%=rs3("titlecolor")%>"><strong>                 
                    <%=CutStr(title,40)%> 
                    </strong></font> </a> </td>
                  <td width="22%" align="right" bgcolor="#EFEFEF"> 
                    <%if showtime="1" then%>
                    <%=datetime%> 
                    <%end if%>
                    <%if showclick="1" then%>
                    <font color=#666666>[<%=rs3("click")%>]</font> 
                    <%end if%>
                    &nbsp; </td>
                  <td width="13%" align="center"> 
                    <%if year(rs3("updatetime"))=year(date()) and month(rs3("updatetime"))=month(date()) and day(rs3("updatetime"))=day(date()) then%>
                    <img src="images/new.gif"> 
                    <%end if%>
                    <%if rs3("goodnews")="1" then%>
                    <img src="images/g.gif" > 
                    <%end if%>
                  </td>
                </tr>
                <tr> 
                  <td width="16%"> <a class=middle href="<%if rs3("titleface")="无" then %><%=newsurl%><% else %> <%=newswwwurl%><%end if%>" target=_blank title="<%=htmlencode4(rs3("title"))%>"> 
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
                    </a> </td>
                  <td colspan="3" valign="top"> <a class=middle href="<%if rs3("titleface")="无" then %><%=newsurl%><% else %> <%=newswwwurl%><%end if%>" target="_blank" title="<%=htmlencode4(title)%>"><%=CutStr(nohtml(rs3("Content")),250)%></a> 
                  </td>
                </tr>
              </table>
            </div></td>
        </tr>
        <%
														rs3.MoveNext
													wend
													%>
        <tr> 
          <td width=100% height="19" align=right><a class=class href='BigClass.asp?typeid=<%=typeid%>&BigClassid=<%=BigClassid%>'><img src="images/more.gif" border="0" alt="更多<%=BigClassName%>"></a>&nbsp;&nbsp;</td>
        </tr>
      </table>
      <%
												end if
											next
										else 
											Response.Write "<table width=98% border=0 height=25><tr><td>暂无大类</td></tr></table>"
										end if
										%>
	<%end select
	rs3.close
	set rs3=nothing
	conn.close
	set conn=nothing
	%>
      <!--模版结束-->
    </td>	
        </tr>
			<tr valign="top">
				<td height="20" background="IMAGES/menu-d-t.gif"></td>
				<td height="20" bgcolor="#FFFFFF"></td>
				<td height="20" background="IMAGES/menu-l-t.gif" bgcolor="#FFFFFF" class="daohang"></td>
			</tr>
		</table>
<!--#include file="bottom.asp"-->