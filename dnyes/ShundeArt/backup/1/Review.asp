<%@ Language=VBScript%>
<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="function.asp"-->
<!--#include file="char.inc"-->

<%if reviewable="1" then
	set rs=server.CreateObject("ADODB.RecordSet")
	rs.Source="select * from "& db_Type_Table &" order by typeorder"
	rs.Open rs.Source,conn,1,1

	dim ArraytypeID(10000),ArraytypeName(10000),Arraytypecontent(10000)
	typeCount=rs.RecordCount
	for i=1 to typeCount
		ArraytypeID(i)=rs("typeID")
		ArraytypeName(i)=rs("typeName")
		Arraytypecontent(i)=rs("typecontent")
		rs.MoveNext
	next
	rs.Close

	NewsID=Request.QueryString("NewsID")

	if NewsID="" then
		Response.Write "<script>alert('未指定参数');history.back()</script>"
		response.end
	else
		if not IsNumeric(NewsID) then
			response.write "<script>alert('非法参数');history.back()</script>"
			response.end
		else
			reviewid=Request.QueryString("reviewid")
		end if
	end if
	%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>文章评论_<%=jjgn%></title>
<LINK href=news.css rel=stylesheet>
<script language="javascript">
<!--
function checkdata()
{
if (document.form1.Author.value=="")
	{
	alert("对不起，请输入您的大名！")
	return false
	 }
	if (document.form1.email.value=="")
		{
		alert("对不起，请输入您的EMAIL！")
		 return false
		}
		if (document.form1.email.value.indexOf("@",0)== -1||document.form1.email.value.indexOf(".",0)==-1)
			{
			alert("对不起，您输入的电子邮件有误！")
			return false
			}
			if (document.form1.content.value=="")
			{
				alert("对不起，请输入留言内容！")
				return false
			}
	}
	function changedata() {
        if ( document.form1.none.checked ) {
                document.form1.Author.value = "匿名";
        	} 
	}
	function MM_goToURL() { //v3.0
	var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
	for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
	}

	function MM_openBrWindow(theURL,winName,features) { //v2.0
	window.open(theURL,winName,features);
	}
	//-->
</script>
</head>
<body topmargin="0">
<!--#include file=top.asp -->
<form method="POST" action="AddReview.asp" name=form1 OnSubmit="return checkdata()">
	
  <table width="760" border="0" align="center" cellpadding="0" cellspacing="0" background="IMAGES/menu-guest-l.gif">
    <tr>
      <td background="IMAGES/menu-guest-l.gif"> 
        <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" >
          <tr>
					<td height="3" ><img src="images/kb.gif" width="9" height="3"></td>
				</tr>
				<tr> 
					<td width="100%" height="25" background="images/menu-guestbook.gif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;栏目导航&nbsp;当前位置：<a class="daohang" href="./" >网站首页</a>&gt;&gt;评论： 
		<%set rs1=server.CreateObject("ADODB.RecordSet")
		rs1.Source="select * from "& db_News_Table &" where newsid="&newsid
		rs1.Open rs1.Source,conn,1,1
		if not rs1.EOF then
			title=htmlencode4(rs1("title"))
			titlecolor=rs1("titlecolor")
			Response.Write title
		else
			Response.Write "未知文章"
		end if
		rs1.close
		set rs1=nothing%>
            </td>
          </tr>
          <tr> 
            <td background="IMAGES/menu-guest-l.gif"><CENTER><script language=javascript src=./zongg/ad.asp?i=13></script></script></CENTER>
          <tr> 
            <td height="25" align="center" background="IMAGES/menu-l-guest.gif">
            </td>
          </tr>
        </table>

        <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" background="IMAGES/menu-guest-l.gif">
          <tr>
    <td ><CENTER><font color="<%=titlecolor%>" size="+2" face="楷体_GB2312"><a href="ReadNews.asp?NewsId=<%=newsid%>"><strong><U><%=title%></U></strong></a>原文</font></CENTER><HR></td>
    </tr>
		<%set rs=server.CreateObject("ADODB.RecordSet")
		rs.Source="select * from "& db_Review_Table &" where NewsID=" & NewsID & " and passed=1 order by reviewid desc"
		rs.Open rs.Source,conn,3,1

		PageShowSize = 10            '每页显示多少个页
		MyPageSize   = 10           '每页显示多少条
	
		If Not IsNumeric(Request("page")) Or IsEmpty(Request("page")) Or Request("page") <=0 Then
			MyPage=1
		Else
			MyPage=Int(Abs(Request("page")))
		End if
		if rs.EOF then  NoReview=1
			if NoReview then
				Response.Write "<tr><td background='IMAGES/menu-guest-l.gif'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;该信息当前没有评论！</td></tr>"
			else
				'if not NoReview then
				Rs.PageSize     = MyPageSize
				MaxPages         = Rs.PageCount
				Rs.absolutepage = MyPage
				total            = Rs.RecordCount
				i = 0
				%>
                <tr>
                  <td>&nbsp;&nbsp;<b>共 <%=total%> 条，第 <%=Mypage%>/<%=Maxpages%> 页，每页 <%=MyPageSize%> 条</b></td>
                </tr>
                <tr>
                  
            <td align="right" background="IMAGES/menu-guest-l.gif">（以下留言仅代表网友的个人观点，不代表本站观点。）&nbsp;&nbsp;</td>
                </tr>
                <%do until Rs.Eof or i = Rs.PageSize
		author=server.HTMLEncode(trim(rs("author")))
		email=server.HTMLEncode(trim(rs("email")))
		reviewip=rs("reviewip")
		content=trim(rs("content"))
		%>
		
		</table>
                
        <table width="98%" border="0" align="center" cellpadding="5" cellspacing="0">
          <tr> 
                  <td colspan="2"> 
				  <table width="100%" border="0" cellpadding="5" cellspacing="0" style="table-layout:fixed; word-break:break-all" >
				        <tr> 
                        <td bgcolor="#EFEFEF" >&nbsp;&nbsp;发表人：<%=author%></td>
                        <td bgcolor="#EFEFEF" > 
                          <p align="right"> 
                            <%if Request.cookies(Forcast_SN)("key")="super" or showip="1" then%>
                            IP：<%=reviewip%> 
                            <%end if%>
                        </td>
                      </tr>
                      <tr> 
                        <td colspan="2"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
                            <tr> 
                              <td height="19">&nbsp;&nbsp;发表人邮件：<a href='mailto:<%=email%>'><%=email%></a></td>
                              <td align="right">&nbsp;&nbsp;发表时间：<%=rs("updatetime")%></td>
                            </tr>
                          </table></td>
                      </tr>
                      <tr> 
                        <td colspan="2"> 
			<%	
			DisplayContent=CheckStr(Content)
			Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;" & displaycontent
			%>
                        </td>
                      </tr>
					  
                    </table>
			<%
			Rs.MoveNext
			i = i + 1
			loop
			%>
                    <table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td align="center">共 <%=total%> 条，第 <%=Mypage%>/<%=Maxpages%> 页，每页 <%=MyPageSize%> 条 
			<%
			url="review.asp?NewsID=" & NewsID									
			PageNextSize=int((MyPage-1)/PageShowSize)+1
			Pagetpage=int((total-1)/Rs.PageSize)+1
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
			End If
			%>
                        </td>
                      </tr>
                    </table>
                    <hr size="1" noshade> <br> </td>
                </tr>
                <%if Request.cookies(Forcast_SN)("username")<>"" then
			set rs1=server.CreateObject("ADODB.RecordSet")
			rs1.Source="select * from "& db_User_Table &" where "& db_User_Name &"='"&Request.cookies(Forcast_SN)("username")&"'"
			rs1.Open rs1.Source,ConnUser,1,1
			%>
			 <tr> 
				
            <td width="100%"> 
              <table border="0" cellspacing="0" width="100%" cellpadding="4">
						<input type=hidden name=NewsID value=<%=NewsID%>>
						<input type=hidden name=url value="http://<%= Request.ServerVariables("SERVER_NAME") %><%=Request.ServerVariables("url")%>?<%=Request.ServerVariables("QUERY_STRING")%>">
						<tr> 
							<td width="13%" align="right"><a name="send"><font color="#FF0000">*</font>用&nbsp;户&nbsp;名：</a></td>
							<td width="87%"> <input type="text" name="Author" size="56" value="<%=Request.cookies(Forcast_SN)("username")%>" readonly></td>
						</tr>
						<tr> 
							<td width="13%" align="right"><font color="#FF0000">*</font>电子邮件：</td>
							<td width="87%"><input type="text" name="email" size="56" value="<%=rs1(db_User_Email)%>"></td>
						</tr>
						<tr> 
							<td width="13%" align="right" valign="top"><font color="#FF0000">*</font>评论内容：</td>
							<td width="87%">
								<script language="javascript">
								document.write ('<iframe src="guestbox.asp?action=new" id="message" width="350" height="100" align=left></iframe>')
								frames.message.document.designMode = "On";
								</script>
			<%rs1.close
			set rs1=nothing%>
							</td>
						</tr>
						<tr> 
							
                  <td width="13%"></td>
							<td width="87%" height="50">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
								<input type="hidden" name="editor" value="<%=editor%>"> 
								<input name="passed" type="hidden" value="<%if reviewcheck="1" then%>1<%else%>0<%end if%>"> 
								<input type="submit" value="发表评论" name="Submit" OnClick="document.form1.content.value = frames.message.document.body.innerHTML;"> <input name="reset" type=reset value="重新填写"><input type="hidden" name="content" value="">
								<input type="hidden" name="title" value="评论：<%=title%>">
							</td>
						</tr>
					</table>
				</td>
			</tr>
		<%else%>
			<tr> 
            <td width="100%" > 
              <table border="0" cellspacing="0" width="100%" cellpadding="4">
						<input type=hidden name=NewsID value=<%=NewsID%>>
						<input type=hidden name=url value="http://<%= Request.ServerVariables("SERVER_NAME") %><%=Request.ServerVariables("url")%>?<%=Request.ServerVariables("QUERY_STRING")%>">
						<tr> 
							<td width="13%" align="right"><a name="send"><font color="#FF0000">*</font>用&nbsp;户&nbsp;名：</a></td>
							<td width="87%"> <input type="text" name="Author" size="56">&nbsp;匿名：<INPUT onclick=changedata() type=checkbox value=true  name=none></td>
						</tr>
						<tr> 
							<td width="13%" align="right"><font color="#FF0000">*</font>电子邮件：</td>
							<td width="87%"> <input name="email" type="text" value=""></td>
						</tr>
						<tr> 
							<td width="13%" align="right" valign="top"><font color="#FF0000">*</font>评论内容：</td>
							<td width="87%">
								<script language="javascript">
								document.write ('<iframe src="guestbox.asp?action=new" id="message" width="350" height="100" align=left></iframe>')
								frames.message.document.designMode = "On";
								</script>
							</td>
						</tr>
						<tr> 
							<td width="13%"></td>
							<td width="87%" height="50">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
								<input type="hidden" name="editor" value="<%=editor%>"> 
								<input name="passed" type="hidden" value="<%if reviewcheck="1" then%>1<%else%>0<%end if%>"> 
								<input type="submit" value="发表评论" name="Submit" OnClick="document.form1.content.value = frames.message.document.body.innerHTML;"> <input name="reset" type=reset value="重新填写"><input type="hidden" name="content" value="">
								<input type="hidden" name="title" value="评论：<%=title%>">
							</td>
						</tr>
					</table>
              <%end if%>
            </td>
			</tr>
</table></td>
</tr>
  </table>
  <table width="760" height="19" border="0" align="center" cellpadding="0" cellspacing="0" background="IMAGES/menu-guest-t.gif">
    <tr>
    <td>&nbsp;</td>
  </tr>
</table>

  <!--#include file=bottom.asp -->
</form>
</body>
</html><%else%><script language=javascript>
alert("对不起，评论功能已被管理员关闭！")
</script>
<body  onload="javascript:window.close()">
</body>
<%end if%>
