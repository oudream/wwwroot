<%@ Language=VBScript%>
<!--#include file=conn.asp -->
<!--#include file=config.asp -->
<!--#include file="function.asp"-->
<!--#include file="char.inc"-->
<%
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
PageShowSize = 10            '每页显示多少个页
MyPageSize   = 20           '每页显示多少条新闻
	
If Not IsNumeric(Request("page")) Or IsEmpty(Request("page")) Or Request("page") <=0 Then
MyPage=1
Else
MyPage=Int(Abs(Request("page")))
End if
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>顺德国际艺术交流中心_记事</title>
<LINK href=news.css rel=stylesheet>
</head>

<body topmargin="0">
<!--#include file=top.asp -->
<table width="760" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <tr valign="top"> 
    <td width="180" background="IMAGES/menu-d.gif"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td valign="top"><table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="100%" id="AutoNumber2" height="23">
              <tr> 
                <td width="100%" height="25" valign="middle" background="IMAGES/menu-d-fl.gif"> 
                  <p align="center"><font color="#000000">文章排行</font></td>
              </tr>
              <tr> 
                <td width="100%" background="IMAGES/menu-d.gif"><div align="center"> 
                    <table width="99%" border="0" cellspacing="0" cellpadding="3">
                      <tr> 
                        <td> 
                          <%
set rs=server.CreateObject("ADODB.RecordSet")
if uselevel=1 then
if Request.cookies(Forcast_SN)("key")="super" or Request.cookies(Forcast_SN)("key")="typemaster" or Request.cookies(Forcast_SN)("key")="bigmaster" or Request.cookies(Forcast_SN)("key")="smallmaster" or Request.cookies(Forcast_SN)("key")="check" then
rs.Source="select top "& top_txt &" * from "& db_News_Table &" where  checkked=1 order by click DESC,newsid desc"
end if
if Request.cookies(Forcast_SN)("key")="" then
rs.Source="select top "& top_txt &" * from "& db_News_Table &" where (checkked=1 and newslevel=0) order by click DESC,newsid desc"
end if
if Request.cookies(Forcast_SN)("key")="selfreg" then
	if Request.cookies(Forcast_SN)("reglevel")=3 then
rs.Source="select top "& top_txt &" * from "& db_News_Table &" where (checkked=1 ) order by click DESC,newsid desc"
	end if
	if Request.cookies(Forcast_SN)("reglevel")=2 then
rs.Source="select top "& top_txt &" * from "& db_News_Table &" where (checkked=1 ) order by click DESC,newsid desc"
	end if
	if Request.cookies(Forcast_SN)("reglevel")=1 then
rs.Source="select top "& top_txt &" * from "& db_News_Table &" where (checkked=1 ) order by click DESC,newsid desc"
	end if
end if
else
rs.Source="select top "& top_txt &" * from "& db_News_Table &" where  checkked=1 order by click DESC,newsid desc"
end if
rs.Open rs.Source,conn,1,1
while not rs.EOF
title=htmlencode4(trim(rs("title")))

%>
                          &nbsp;· <a class=middle href="ReadNews.asp?NewsID=<%=rs("NewsID")%>" target="_blank" title="<%=title%>"><font color="<%=rs("titlecolor")%>">                           
                          <%=CutStr(title,18)%> 
                          </font></a>&nbsp;<font color=red><%=rs("click")%></font><br> 
                          <%
rs.MoveNext
wend
rs.close
%>
                        </td>
                      </tr>
                    </table>
                  </div></td>
              </tr>
            </table></td>
        </tr>
        <tr bordercolor="#999999"> 
          <td height="25" align="center" bordercolor="#999999" background="IMAGES/menu-d-hot.gif"><FONT color=#000000>本月热门</FONT></td>
        </tr>
        <tr bordercolor="#999999"> 
          <td height="36" align="center" valign=top background="IMAGES/menu-d.gif"> 
            <br> <table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="93%" id="AutoNumber5">
              <% 
 dim ii
ii = 0
if uselevel=1 then
if Request.cookies(Forcast_SN)("key")="" then
rs.Source="select top " & top_txt & " * from "& db_News_Table &" where year(updatetime)=year(now()) and month(updatetime)=month(now()) and checkked=1 and newslevel=0 order by click DESC"   '选择本月
end if
if Request.cookies(Forcast_SN)("key")="selfreg" then
	if Request.cookies(Forcast_SN)("reglevel")="3" then
	rs.Source="select top " & top_txt & " * from "& db_News_Table &" where year(updatetime)=year(now()) and month(updatetime)=month(now()) and checkked=1 order by click DESC"   '选择本月
	end if
	if Request.cookies(Forcast_SN)("reglevel")="2" then
	rs.Source="select top " & top_txt & " * from "& db_News_Table &" where year(updatetime)=year(now()) and month(updatetime)=month(now()) and checkked=1 order by click DESC"   '选择本月
	end if
	if Request.cookies(Forcast_SN)("reglevel")="1" then
	rs.Source="select top " & top_txt & " * from "& db_News_Table &" where year(updatetime)=year(now()) and month(updatetime)=month(now()) and checkked=1 order by click DESC"   '选择本月
	end if
end if
if Request.cookies(Forcast_SN)("key")="super" or Request.cookies(Forcast_SN)("key")="typemaster" or Request.cookies(Forcast_SN)("key")="bigmaster" or Request.cookies(Forcast_SN)("key")="smallmaster" or Request.cookies(Forcast_SN)("key")="check" then
rs.Source="select top " & top_txt & " * from "& db_News_Table &" where year(updatetime)=year(now()) and month(updatetime)=month(now()) and checkked=1 order by click DESC"   '选择本月
end if
else
rs.Source="select top " & top_txt & " * from "& db_News_Table &" where year(updatetime)=year(now()) and month(updatetime)=month(now()) and checkked=1 order by click DESC"   '选择本月
end if
rs.Open rs.Source,conn,1,1
	if rs.bof and rs.eof then 
		response.write "<td align=center>无</td>" 
	else 
	do while not rs.eof 
%>
              <tr> 
                <td height=18> · 
                  <%if rs("picnews")=1 then%>
                  <img src="images/news_img.gif"> 
                  <%end if%>
                  <a class=middle href="ReadNews.asp?NewsID=<%=rs("NewsID")%>" title="<%=htmlencode4(rs("title"))%>" target="_blank" > 
                  <%=htmlencode4(rs("title"))%></a><font color="#666666">[<%=rs("click")%>]</font> 
                </td>
              </tr>
              <%  ii = ii + 1
    if ii>50 then exit do
	rs.movenext     
	loop
	end if  
	rs.close   
	set rs=nothing
%>
            </table></td>
        </tr>
        <tr> 
          
        </tr>
      </table>
    </td>
    <td width="6" bgcolor="#FFFFFF">　</td>
    <td background="IMAGES/menu-l.gif"> 
      <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td height="25" background="IMAGES/menu-l-m.gif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font class=m_tittle>&nbsp;</font>栏目导航&nbsp;<b><a class="daohang" href="./" > 
            留言</a></b></td>
        </tr>
        <tr> 
          <td background="IMAGES/menu-l.gif">&nbsp;</td>
        </tr>
        <tr> 
          
        </tr>
      </table>
      <TABLE height=38 cellSpacing=0 cellPadding=0 width=100% border=0>
        <TBODY>
          <TR> 
            <TD vAlign=bottom align=right width=8></TD>
            <TD vAlign=bottom align=middle> <TABLE class=pt9 cellSpacing=0 cellPadding=0 width=510 
                  align=center border=0>
                <TBODY>
                  <SCRIPT language=javascript>
function checkform(){
if (frm_back.name.value==""){
	alert("姓名不能为空!");
	frm_back.name.focus();
	return false;
}

if (frm_back.tel.value==""){
	alert("电话不能为空!");
	frm_back.tel.focus();
	return false;
}

if (frm_back.e_mail.value==""){
	alert("E-mail不能为空!");
	frm_back.e_mail.focus();
	return false;
}

if (frm_back.message.value==""){
	alert("留言不能为空!");
	frm_back.message.focus();
	return false;
}

}</SCRIPT>
                <FORM name=frm_back onsubmit="return checkform();" 
                    action=feedback.asp
                    method=post>
                  <INPUT type=hidden 
                    value=support@giantpacific.com.cn name=sendemail>
                </form>
                  <TD colSpan=2 >&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;感谢您一直以来对我们的关心和支持。希望您对我们的产品，服务，管理，<br> 
                  <br>
                    宣传或其他方面提出宝贵的意见。您的资料将有助于我们向您联络或解答问题，<font color="#FF0000">并将得到<br>
                  <br>
                  完善的保密</font>。 <br> <br> <br> &nbsp;&nbsp;注：下面有“ <font color="#FF0000">* 
                  </font>标记的项目为必须填写项目。<br> <br> </TD>
                <TR> 
                  <TD  width=394 height="40"> <DIV align=center>您的姓名：</DIV></TD>
                  <TD  align=left width=394 height=30><INPUT 
                        id=name name=name> <font color="#FF0000">*</font></TD>
                </TR>
                <TR> 
                  <TD  width=394 height="40"> <DIV align=center>公司名称：</DIV></TD>
                  <TD  align=left width=394 height=30><INPUT 
                        id=compname size=52 name=compname> </TD>
                </TR>
                <TR> 
                  <TD  width=394 height="40"> <DIV align=center>联系地址：</DIV></TD>
                  <TD  align=left width=394 height=30><INPUT 
                        id=compaddres size=52 name=compaddres> </TD>
                </TR>
                <TR> 
                  <TD  width=394 height="40"> <DIV align=center>联系电话：</DIV></TD>
                  <TD  align=left width=394 height=30><INPUT 
                        id=telnumber name=tel> <font color="#FF0000">*</font> 
                  </TD>
                </TR>
                <TR> 
                  <TD  width=394 height="40"> <DIV align=center>传　　真：</DIV></TD>
                  <TD  align=left width=394 height=30><INPUT 
                        id=fax name=fax> </TD>
                </TR>
                <TR> 
                  <TD  width=394 height="40"> <DIV align=center>电子信箱：</DIV></TD>
                  <TD  align=left width=394 height=30><INPUT 
                        id=e_mail size=52 name=e_mail> <font color="#FF0000">*</font></TD>
                </TR>
                <TR> 
                  <TD  width=394 height="40"> <DIV align=center>信息类型：</DIV></TD>
                  <TD  align=left width=394 height=30><select name="infotype" size="1" >
                      <option value="建　　议" selected>建　　议</option>
                      <option value="投　　诉">投　　诉</option>
                      <option value="订     货">订　　货</option>
                      <option value="普通联络">普通联络</option>
                      <option value="紧急联络">紧急联络</option>
                    </select> <font color="#FF0000">*</font> </TD>
                </TR>
                <TR> 
                  <TD  width=394>&nbsp;</TD>
                  <TD  align=left width=394>&nbsp;</TD>
                </TR>
                <TR> 
                  <TD width=394 height="40"  vAlign=top> <DIV align=center>具体内容：</DIV></TD>
                  <TD  align=left width=394><TEXTAREA id=message name=message rows=10 cols=47></TEXTAREA> 
                    <font color="#FF0000">*</font> </TD>
                </TR>
                <TR> 
                  <TD colSpan=2> <DIV align=center> 
                      <TABLE cellSpacing=0 cellPadding=0 width="100%" 
border=0>
                        <TBODY>
                          <TR> 
                            <TD height=10></TD>
                          </TR>
                        </TBODY>
                      </TABLE>
                      <INPUT 
                        type=hidden name=tHtml>
                      <input type="hidden" name="sendready" id="sendready" value="ok">
                      <INPUT id=button1 type=submit value=马上发送 name=button1>
                      <INPUT type=reset value=全部重写 name=B1>
                    </DIV></TD>
                </TR>
              </TABLE></TD>
            <TD width=10>&nbsp;</TD>
          </TR>
      </TABLE></td>
  </tr>
  <tr valign="top"> 
    <td height="20" background="IMAGES/menu-d-t.gif">&nbsp;</td>
    <td height="20" bgcolor="#FFFFFF">&nbsp;</td>
    <td height="20" background="IMAGES/menu-l-t.gif" class="daohang">&nbsp;</td>
  </tr>
</table>
<!--#include file=bottom.asp -->

</body>

</html>