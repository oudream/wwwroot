<%@ Language=VBScript%>
<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="function.asp"-->
<%
request_BigClassid=checkstr(Request("BigClassid"))
keyword=trim(checkstr(Request("keyword")))


PageShowSize = 10	'每页显示多少个页
MyPageSize   = 20	'每页显示多少条
	
If Not IsNumeric(Request("page")) Or IsEmpty(Request("page")) Or Request("page") <=0 Then
	MyPage=1
Else
	MyPage=Int(Abs(Request("page")))
End if

if keyword="" or keyword="关键字" then
	%>
	<script language=javascript>
		alert("请输入查询关键字！")
		history.back()
	</script>
	<body onload=javascript:window.close()></body> 
	<%
	Response.End
end if
if request("action")="" then
	findword="title like '%"&keyword&"%' or content like '%"&keyword&"%' or author like '%"&keyword&"%' or editor like '%"&keyword&"%' or about like '%"&keyword&"%'"
elseif request("action")="title" then
	findword="title like '%"&keyword&"%' "
elseif request("action")="content" then
	findword="content like '%"&keyword&"%' "
elseif request("action")="editor" then
	findword="editor like '%"&keyword&"%' or author like '%"&keyword&"%' "
elseif request("action")="about" then
	findword="about like '%"&keyword&"%' "
elseif request("action")="UpdateTime" then
	findword="UpdateTime like '%"&keyword&"%' "
end if 

set rs=server.CreateObject("ADODB.RecordSet")
sql1="select * from "& db_Type_Table &" order by typeorder"
rs.Open sql1,conn,1,1

dim ArraytypeID(10000),ArraytypeName(10000),Arraytypecontent(10000)
typeCount=rs.RecordCount
for i=1 to typeCount
	ArraytypeID(i)=rs("typeID")
	ArraytypeName(i)=rs("typeName")
	Arraytypecontent(i)=rs("typecontent")
	rs.MoveNext
next
rs.Close
%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>关键字：<%=keyword%>__在[<%=Request("action")%>]__搜索__<%=jjgn%></title>
<LINK href=news.css rel=stylesheet>
</head>
<BODY ONLOAD="newCalendar()" OnUnload="window.returnValue = document.all.ret.value;" topmargin="0"> 
<!--#include file="top.asp"-->
<table width="760" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
	<tr valign="top"> 
		<td width="180" background="IMAGES/menu-d.gif"><div align="center">
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#EAEAF4">
  <tr>
    <td><div align="center"><IMG SRC="IMAGES/left1-m9.gif"></div></td>
  </tr>
  <tr> 

<td><div align="center">
<STYLE TYPE="text/css"> 
.today {font-weight:bold;BACKGROUND: #6699cc} 
.satday{color:green} 
.sunday{color:red} 
.days {font-weight:bold} 
</STYLE>
<script language="JavaScript"> 
//中文月份,如果想显示英文月份，修改下面的注释 
/*var months = new Array("January?, "February?, "March", 
"April", "May", "June", "July", "August", "September", 
"October", "November", "December");*/ 
var months = new Array("一月", "二月", "三月", 
"四月", "五月", "六月", "七月", "八月", "九月", 
"十月", "十一月", "十二月"); 
var daysInMonth = new Array(31, 28, 31, 30, 31, 30, 31, 31, 
30, 31, 30, 31); 
//中文周 如果想显示 英文的，修改下面的注释 
/*var days = new Array("Sunday", "Monday", "Tuesday", 
"Wednesday", "Thursday", "Friday", "Saturday");*/ 
var days = new Array("日","一", "二", "三", 
"四", "五", "六"); 
function getDays(month, year) { 
//下面的这段代码是判断当前是否是闰年的 
if (1 == month) 
return ((0 == year % 4) && (0 != (year % 100))) || 
(0 == year % 400) ? 29 : 28; 
else 
return daysInMonth[month]; 
} 

function getToday() { 
//得到今天的年,月,日 
this.now = new Date(); 
this.year = this.now.getFullYear(); 
this.month = this.now.getMonth(); 
this.day = this.now.getDate(); 
} 


today = new getToday(); 

function newCalendar() { 

today = new getToday(); 
var parseYear = parseInt(document.all.year 
[document.all.year.selectedIndex].text); 
var newCal = new Date(parseYear, 
document.all.month.selectedIndex, 1); 
var day = -1; 
var startDay = newCal.getDay(); 
var daily = 0; 
if ((today.year == newCal.getFullYear()) &&(today.month == newCal.getMonth())) 
day = today.day; 
var tableCal = document.all.calendar.tBodies.dayList; 
var intDaysInMonth =getDays(newCal.getMonth(), newCal.getFullYear()); 
for (var intWeek = 0; intWeek < tableCal.rows.length;intWeek++) 
for (var intDay = 0;intDay < tableCal.rows[intWeek].cells.length;intDay++) 
{ 
var cell = tableCal.rows[intWeek].cells[intDay]; 
if ((intDay == startDay) && (0 == daily)) 
daily = 1; 
if(day==daily) 
//今天，调用今天的Class 
cell.className = "today"; 
else if(intDay==6) 
//周六 
cell.className = "sunday"; 
else if (intDay==0) 
//周日 
cell.className ="satday"; 
else 
//平常 
cell.className="normal"; 

if ((daily > 0) && (daily <= intDaysInMonth)) 
{ 
cell.innerText = daily; 
daily++; 
} 
else 
cell.innerText = ""; 
} 
} 

function getDate() { 
var sDate; 
//这段代码处理鼠标点击的情况 
if ("TD" == event.srcElement.tagName) 
if ("" != event.srcElement.innerText) 
{ 
sDate = document.all.year.value + "年" + document.all.month.value + "月" + event.srcElement.innerText + "日"; 
//alert(sDate); 
window.open ("Result.asp?keyword="+document.all.year.value+"-"+document.all.month.value+"-"+event.srcElement.innerText+"&action=UpdateTime") 
} 
} 
          </script> <input type="hidden" name="ret"> <table id="calendar" border="0" cellpadding="3" cellspacing="0"  bgcolor="#EAEAF4">

                        <thead>
                          <tr> 
<TD COLSPAN=7 ALIGN=CENTER> 
<SELECT ID="month" ONCHANGE="newCalendar()"> 
<SCRIPT LANGUAGE="JavaScript"> 
for (var intLoop = 0; intLoop < months.length; 
intLoop++) 
document.write("<OPTION VALUE= " + (intLoop + 1) + " " + 
(today.month == intLoop ? 
"Selected" : "") + ">" + 
months[intLoop]); 
</SCRIPT> 
</SELECT> 

<SELECT ID="year" ONCHANGE="newCalendar()"> 
<SCRIPT LANGUAGE="JavaScript"> 
for (var intLoop = today.year-50; intLoop < (today.year + 10); 
intLoop++) 
document.write("<OPTION VALUE= " + intLoop + " " + 
(today.year == intLoop ? 
"Selected" : "") + ">" + 
intLoop); 
</script>
                              </select>
                              <br>
                              <br> </td>
                          </tr>
                          <tr class="days"> 
                            <script language="JavaScript"> 

document.write("<TD class=satday>" + days[0] + "</TD>"); 
for (var intLoop = 1; intLoop < days.length-1; 
intLoop++) 
document.write("<TD>" + days[intLoop] + "</TD>"); 
document.write("<TD class=sunday>" + days[intLoop] + "</TD>"); 
</script>
                          </tr>
                        </thead>
                        <tbody border=1 cellspacing="0" cellpadding="0" id="dayList"align=CENTER onClick="getDate()">
                          <script language="JavaScript"> 
for (var intWeeks = 0; intWeeks < 6; intWeeks++) { 
document.write("<TR style='cursor:hand'>"); 
for (var intDays = 0; intDays < days.length; 
intDays++) 
document.write("<TD></TD>"); 
document.write("</TR>"); 
} 
</script>
</tbody>
</table>
                     
					 </td></div>
                  </tr>
  
</table>


		</div></td>
		<td width="6"></td>
		<td background="IMAGES/menu-l.gif">
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td height="25" background="IMAGES/menu-l-m.gif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font class=m_tittle>&nbsp;</font>栏目导航</td>
				</tr>
				<tr>
					<td  background="IMAGES/menu-l.gif" ><br>&nbsp;&nbsp;<a class="daohang" href="./" >网站首页</a>&gt;&gt;搜索>>关键字：<%=keyword%></b><br><br></td>
				</tr>
				<tr> 
					<td width=100% height="25" valign="middle" background="IMAGES/menu-l-zj.gif"></td>
				</tr>  
			</table>
			<table width="100%"  border="0" align="center" cellpadding="0" cellspacing="0" id="AutoNumber3" style="border-collapse: collapse">
				<tr> 
					<td background="IMAGES/menu-l.gif"></td>
				</tr>
				<tr> 
					<td background="IMAGES/menu-l.gif" >
						<table border="0" cellpadding="3" cellspacing="0" width="100%">
<%set rs=server.CreateObject("ADODB.RecordSet")
if request_BigClassid<>"" then
	if Request.cookies(Forcast_SN)("key")="" then
		key_word="select * from "& db_News_Table &" where BigClassid=" & request_BigClassid &" and " & findword & " and checkked=1 and newslevel=0 order by NewsID DESC"
	end if
	if Request.cookies(Forcast_SN)("key")="super" or Request.cookies(Forcast_SN)("key")="typemaster" or Request.cookies(Forcast_SN)("key")="bigmaster" or Request.cookies(Forcast_SN)("key")="smallmaster" or Request.cookies(Forcast_SN)("key")="check" then
		key_word="select * from "& db_News_Table &" where BigClassid=" & request_BigClassid &" and " & findword & " and checkked=1 order by NewsID DESC"
	end if
	if Request.cookies(Forcast_SN)("key")="selfreg" then
		if Request.cookies(Forcast_SN)("reglevel")=3 then
			key_word="select * from "& db_News_Table &" where BigClassid=" & request_BigClassid &" and " & findword & " and checkked=1  order by NewsID DESC"
		end if
		if Request.cookies(Forcast_SN)("reglevel")=2 then
			key_word="select * from "& db_News_Table &" where BigClassid=" & request_BigClassid &" and " & findword & " and checkked=1  order by NewsID DESC"
		end if
		if Request.cookies(Forcast_SN)("reglevel")=1 then
			key_word="select * from "& db_News_Table &" where BigClassid=" & request_BigClassid &" and " & findword & " and checkked=1  order by NewsID DESC"
		end if
	end if
else
	if Request.cookies(Forcast_SN)("key")="" then
		key_word="select * from "& db_News_Table &" where " & findword & " and checkked=1 and newslevel=0 order by NewsID DESC"
	end if
	if Request.cookies(Forcast_SN)("key")="super" or Request.cookies(Forcast_SN)("key")="typemaster" or Request.cookies(Forcast_SN)("key")="bigmaster" or Request.cookies(Forcast_SN)("key")="smallmaster" or Request.cookies(Forcast_SN)("key")="check" then
		key_word="select * from "& db_News_Table &" where " & findword & " and checkked=1 order by NewsID DESC"
	end if
	if Request.cookies(Forcast_SN)("key")="selfreg" then
		if Request.cookies(Forcast_SN)("reglevel")=3 then
			key_word="select * from "& db_News_Table &" where " & findword & " and checkked=1 order by NewsID DESC"
		end if
		if Request.cookies(Forcast_SN)("reglevel")=2 then
			key_word="select * from "& db_News_Table &" where " & findword & " and checkked=1 order by NewsID DESC"
		end if
		if Request.cookies(Forcast_SN)("reglevel")=1 then
			key_word="select * from "& db_News_Table &" where " & findword & " and checkked=1 order by NewsID DESC"
		end if
	end if
end if
rs.Open key_word,conn,1,1
rs.PageSize=20
rs.CacheSize = RS.PageSize

for i=1 to rs.PageSize *( page-1)
	if not rs.EOF then
		rs.MoveNext
	end if
next
Response.Write "<tr><td  width=100% height=25 align=center >&nbsp;"
if rs.EOF then
	Response.Write "<img src='images/Image24.gif' ></img><br><br>"
	Response.Write "<font color=red>抱歉，没有搜索到相关的资料！</font>"
else
	rs.PageSize     = MyPageSize
	MaxPages         = rs.PageCount
	rs.absolutepage = MyPage
	total            = rs.RecordCount
	Response.Write "共搜索到" & total & "条相关资料，当前第"& myPage &"/"& MaxPages &"页，每页"& rs.PageSize &"条"
end if
Response.Write "</td></tr>"
If Not rs.eof then
	i = 0
	do until rs.Eof or i = rs.PageSize
		if rs("picname")<>"" then
			img="<font color=red>[图]</font>"
		else
			img=""
		end if
		%>
							<tr> 
		<%
		title=htmlencode4(trim(rs("title")))
		dim bigclassid
		bigclassid=rs("bigclassid")
		set rs11=server.CreateObject("ADODB.RecordSet")
		rs11.Source="select * from "& db_BigClass_Table &" where bigclassid="&bigclassid&""
		rs11.Open rs11.Source,conn,1,1
		typeid=rs11("typeid")
		rs11.Close
		dim mode
		set rs12=server.CreateObject("ADODB.RecordSet")
		rs12.Source="select * from "& db_Type_Table &" where typeid="&typeid&""
		rs12.Open rs12.Source,conn,1,1
		mode=rs12("mode")
		rs12.Close
		%>
								<td width="100%" height="20">&nbsp;<img src="images/006.gif" width="8" height="10" border="0">
				<%if mode="2" then%>
			    <%=img%>
		         <%end if%>
				<a class=middle href="ReadNews.asp?NewsID=<%=rs("NewsID")%>" target=_blank title="<%=title%>">
		<font color="<%=rs("titlecolor")%>">
	    <%=CutStr(title,46)%>
		</font></a> <font class=middle>(<%=rs("UpdateTime")%>)[<font color="#ff0000"><%=rs("click")%></font>]</font>
								</td>
							</tr>
		<%
		rs.MoveNext
		i = i + 1
	loop
	%>
							<tr> 
								<td width="100%" align=center>第 <%=Mypage%>/<%=Maxpages%> 页，每页 <%=MyPageSize%> 条 
	<%
	url="Result.asp?action="&request("action")&"&keyword=" & keyword
	PageNextSize=int((MyPage-1)/PageShowSize)+1
	Pagetpage=int((total-1)/rs.PageSize)+1

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
	end if
	rs.close					
	%>
								</td>
							</tr>
							<tr>
							</tr>
						</table>
					</td>
				</tr>
				<tr> 
					<td valign=top></td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td height="20" background="IMAGES/menu-d-t.gif"></td>
		<td></td>
		<td height="20" background="IMAGES/menu-l-t.gif"></td>
	</tr>
</table>
<!--#include file="bottom.asp"-->
</body>
</html>
