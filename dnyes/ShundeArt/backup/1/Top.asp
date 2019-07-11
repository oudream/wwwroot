<!--#include file="ConnUser.asp"-->
<!--#include file="inc/config.asp"-->
<!--#include file="char.inc"-->


<%if TransShow="on" then%>
<META content=revealTrans(Transition=23,Duration=1.0) http-equiv=Page-Enter>
<META content=revealTrans(Transition=23,Duration=1.0) http-equiv=Page-Exit>
	<SCRIPT language=JavaScript>
	<!--
	//页面随机切换效果
	function transDemo(n) {
	if (document.all && navigator.userAgent.indexOf("Mac")==-1) {
	t=document.all.transmeta;
	t.style.width=document.body.clientWidth;
	t.style.height=document.body.offsetHeight;
	t.style.top=document.body.scrollTop;
	t.style.backgroundColor="#003333";
	t.style.visibility="visible";
	t.filters[0].transition=n;
	setTimeout("transShow()"); // separated to force screen paint
	} else {
	alert("You can view transitions only on Windows IE 4.0 and later.");
	} 
	}
	
	function transShow() {
	t.filters[0].Apply();
	t.style.visibility="hidden";
	t.filters[0].Play();
	}
	//-->
	</SCRIPT>
<%end if%>
<%		'获取当前 URL
dim ViewUrl
if Request.ServerVariables("QUERY_STRING")<>"" then
	ViewUrl="http://"& Request.ServerVariables("SERVER_NAME") &""& Request.ServerVariables("url") &"?"& Request.ServerVariables("QUERY_STRING")&""
else
	ViewUrl="http://"& Request.ServerVariables("SERVER_NAME") &""& Request.ServerVariables("url") &""
end if
response.cookies(Forcast_SN)("ViewUrl")=ViewUrl%>

<div id=menuDiv style='Z-INDEX: 2; VISIBILITY: hidden; WIDTH: 1px; POSITION: absolute; HEIGHT: 1px; BACKGROUND-COLOR: #9cc5f8'></div>
<%if topbg=1 then %>
	<script src="clearevents.js"></script>
	<noscript><iframe src=*></iframe></noscript>
<%end if%>
<div align="center">
  <table width="760" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
    <tr>
      <td height="6"><img src="pic/1.gif" width="1" height="1"></td>
    </tr>
  </table>
  <table width="760" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td><img src="pic/logo2.gif" width="760" height="151"></td>
    </tr>
  </table>
  <table width="760" border="0" cellpadding="0" cellspacing="0" bgcolor="#99CDFF">
	</table>
</div>
<table width="760" border="0" align="center" cellpadding="2" cellspacing="0">
  <tr> 
    <td width="30%" background="images/top5.gif">&nbsp;&nbsp;http://www.ShundeArt.com/</td>
    <form method="post" name="myform" action="Result.asp" target="newwindow">
      <td height="22" align="right" background="images/top5.gif"> 
        <%if showsearch=1 then%>
        <%if search="1" then%>
        <select name="action" style="border:1 solid #CCD5F9; font-size:9pt; background-color:#EAEAF4"　size="1">
          <option selected="" value="">全部内容</option>
          <option value="title">按标题</option>
          <option value="content">按内容</option>
          <option value="editor">按作者</option>
          <option value="about">按关键字</option>
        </select> <input type="text" name="keyword" style="font-size:9pt; background-color:#EAEAF4" size="10" value="关键字" onFocus="if (value =='关键字'){value =''}" onBlur="if (value ==''){value='关键字'}" maxlength="50" /> 
        <input type="submit" name="Submit" value="搜索" /　style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'" > 
        <%else%>
        <%
				dim count
				set rs7=server.createobject("adodb.recordset")
				rs7.Source= "select * from "& db_SmallClass_Table &" order by SmallClassID asc"
				rs7.open rs7.Source,conn,1,1
				%>
        <script language="JavaScript">
var onecount;
onecount=0;
subcat = new Array();
        <%
        count = 0		
        do while not rs7.eof 
        %>
subcat[<%=count%>] = new Array("<%=rs7("SmallClassName")%>","<%=rs7("BigClassid")%>","<%=rs7("SmallClassID")%>");
        <%
        count = count + 1
        rs7.movenext
        loop
        rs7.close
		set rs7=nothing
        %>
onecount=<%=count%>;

function changelocation(locationid)
    {
    document.myform.SmallClassID.length = 0; 
    document.myform.SmallClassID.options[document.myform.SmallClassID.length] = new Option("全部小类", "");
    var locationid=locationid;
    var i;
    for (i=0;i < onecount; i++)
        {
            if (subcat[i][1] == locationid)
            { 
                document.myform.SmallClassID.options[document.myform.SmallClassID.length] = new Option(subcat[i][0], subcat[i][2]);
            }        
        }
        
    }    
</script> <select name="select" style="border:1 solid #CCD5F9; font-size:9pt; background-color:#EAEAF4" size="1">
          <option selected="" value="">全部内容</option>
          <option value="title">按标题</option>
          <option value="content">按内容</option>
          <option value="editor">按作者</option>
          <option value="about">按关键字</option>
        </select> <select name="BigClassid" style="border:1 solid #CCD5F9; font-size:9pt; background-color:#EAEAF4" onChange="changelocation(document.myform.BigClassid.options[document.myform.BigClassid.selectedIndex].value)" size="1">
          <option selected="" value="">全部大类</option>
          <%         
					set rs8=server.CreateObject("ADODB.RecordSet")
					rs8.Source="select * from " & db_BigClass_Table & " order by BigClassID"
					rs8.Open rs8.Source,conn,1,1
					do while not rs8.eof
						%>
          <option value="<%=rs8("BigClassid")%>"><%=rs8("BigclassName")%></option>
          <%
						rs8.movenext
					loop
					rs8.close
					set rs8=nothing
				        %>
        </select> <select name="SmallClassID" style="border:1 solid #CCD5F9; font-size:9pt; background-color:#EAEAF4">
          <option selected="" value="">全部小类</option>
        </select> <input type="text" name="keyword" style="font-size:9pt; background-color:#EAEAF4" size="10" value="关键字" onFocus="if (value =='关键字'){value =''}" onBlur="if (value ==''){value='关键字'}" maxlength="50" /> 
        <input type="submit" name="Submit" value="搜索" / style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'" > 
        <%end if%>
        <%end if%>
      </td>
    </form>
  </tr>
  </table>
  <table width="760" border="0" align="center" cellpadding="2" cellspacing="0">
    <tr> 
    <td width="30%" background="pic/4x4.gif"> <font color="#FFFFFF"> 
      &nbsp;&nbsp;<%Response.Write FormatDateTime(Date(),1) & "  星期"& Mid("日一二三四五六",WeekDay(Date),1) %>
      </font> </td>
    <td height="22" align="right" background="pic/4x4.gif"> <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td align="right" class="top"> | <a href="Index.asp" class="top1">本站首页</a> 
            | <a href="About.asp" class="top1">中心介绍</a> | <a href="ZuoPin.asp" class="top1">名家作品</a> 
            | <a href="Note.asp" class="top1">中心记事</a> | <a href="TiZhi.asp" class="top1">名家题字</a> 
            | <a href="FeedBack.asp" class="top1">请您留言</a> | <a href="Contact.asp" class="top1">联系我们</a> 
            |</td>
        </tr>
      </table></td>
  </tr>
  <tr align="right"> 
    <td height="22" colspan="2" bgcolor="#535396"> <table border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td class="top"> 
            <!--#include file="Menu.asp"-->
          </td>
        </tr>
      </table></td>
  </tr>
</table>
