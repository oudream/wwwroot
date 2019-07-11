<%@ LANGUAGE="VBSCRIPT" %>
<!--#include file=config.asp -->

<script language=JavaScript>
<!--

function opw(url,name, width, height) { //v2.0
window.open(url,name,''+'width='+width+',height='+height+',scrollbars=yes'+'');
}
//-->
</script>
<!--#include file="top.asp"-->

<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber2">
   <form method=post action=list.asp?type=search>

 <tr>
      <td width="100%" align="right"><font color="#808080">快速搜索=&gt;&gt;
      <select size="1" name="adorder" style="color: #808080">
<option value="id">广告ID</option>
<option value="name">名称关键字</option>
</select> <input type="text" name="nr" size="20">
<input type="submit" value="查询" name="B1" style="color: #808080"></font><hr color="#808080" size="1"></td>
    </tr></form>
  </table>
  </center>
</div>




          <table border=0 width=100% cellspacing=3 cellpadding=3>
            <tr> 
              <td align=center> 
                <%dim MaxPerPage,adssql,adsrs,totalPut,CurrentPage,TotalPages,i,advlistact
                  dim px
                  if request("px")="" then
                  px="desc"
                  else
                  px=""
                  end if
                  
                   Select Case request.querystring("type")
                   
                          Case "img"
                           adssql="select * from ads where act=1 and (xslei='gif' or xslei='swf') order by regtime "&px
                %>
                <font size="3" color=red><b>正常播放的图片类广告条列表</b></font>  <a href=?type=<%=request.querystring("type")%>&px=x>升</a>  <a href=?type=<%=request.querystring("type")%>>降</a>
                <%        Case "txt"
                           adssql="select * from ads where act=1 and xslei='txt' order by regtime "&px
                %>
                <font size="3" color=red><b>正常播放的纯文本广告条列表</b></font>  <a href=?type=<%=request.querystring("type")%>&px=x>升</a>  <a href=?type=<%=request.querystring("type")%>>降</a>
                <%
                          Case "close"
                           adssql="select * from ads where act=0 order by regtime "&px

                %>
                <font size="3" color=red><b>处于暂停而未失效的广告条列表</b></font>  <a href=?type=<%=request.querystring("type")%>&px=x>升</a>  <a href=?type=<%=request.querystring("type")%>>降</a>
                <%
                          Case "lose"
                           adssql="select * from ads where act=2 order by regtime "&px
                %>
                <font size="3" color=red><b>已经失效的的广告条列表</b></font>  <a href=?type=<%=request.querystring("type")%>&px=x>升</a>  <a href=?type=<%=request.querystring("type")%>>降</a> 
                <%
                          Case "click"
                           adssql="select * from ads where act<>2 order by click "&px
                %>
                <font size="3" color=red><b>按点击次数<%if px="desc" then: response.write "降序":else:response.write "升续":end if%>排列未失效广告条</b></font>  <a href=?type=<%=request.querystring("type")%>&px=x>升</a>  <a href=?type=<%=request.querystring("type")%>>降</a>
               <%
                          Case "show"
                           adssql="select * from ads where act<>2 order by show "&px
                %>
                <font size="3" color=red><b>按显示次数<%if px="desc" then: response.write "降序":else:response.write "升续":end if%>排列未失效广告条</b></font>  <a href=?type=<%=request.querystring("type")%>&px=x>升</a>  <a href=?type=<%=request.querystring("type")%>>降</a>
               <%
                          Case "place"
                          
                          if isnumeric(request("place"))=true then
                           adssql="select * from ads where act=1 and place="&trim(request("place"))&" order by regtime "&px
                          
                %>
                <font size="3" color=red><b>ID为<%=request("place")%>的广告位中正常播放的广告条</b></font>  <a href=?type=<%=request.querystring("type")%>&place=<%=request("place")%>&px=x>升</a>  <a href=?type=<%=request.querystring("type")%>&place=<%=request("place")%>>降</a>
                <%else
                  adssql="select * from ads where act=1 order by regtime "&px
                %>
                <font size="3" color=red><b>所有正常播放的广告条列表</b></font>  <a href=?type=<%=request.querystring("type")%>&px=x>升</a>  <a href=?type=<%=request.querystring("type")%>>降</a>
                        
                <%end if%>
               <%
                          Case "search"
                          if request("adorder")="id" and isnumeric(request("nr"))=true then
                           adssql="select * from ads where id="&trim(request("nr"))
                          
                %>
                <font size="3" color=red><b>查询 ID为<%=request("nr")%> 的广告条信息</b></font>
                <%        else
                  adssql="select * from ads where sitename like '%"&request("nr")&"%' order by regtime "&px
                %>
                <font size="3" color=red><b>查询名称含有关键字“<%=request("nr")%>”广告条</b></font>  <a href=?type=<%=request.querystring("type")%>&px=x>升</a>  <a href=?type=<%=request.querystring("type")%>>降</a>
                        
                <%end if%>

                <%       
                          Case else
                          adssql="select * from ads where act=1 order by regtime "&px
                %>
                <font size="3" color=red><b>所有正常播放的广告条列表</b></font>  <a href=?type=<%=request.querystring("type")%>&px=x>升</a>  <a href=?type=<%=request.querystring("type")%>>降</a>
                <%
                    
                     
                    end Select
                %>
              </td>
            </tr>
          </table>
            <%

adsconn.Open adsdata
if isnumeric(request("page")) then
currentPage=cint(request("page"))
else
currentPage=1
end if
set adsrs=server.createobject("adodb.recordset")

adsrs.open adssql,adsconn,1,1
if adsrs.eof and adsrs.bof then
response.write "<tr><td bgcolor=#ffffff align=center><BR><BR>没有任何相关记录<BR><BR><BR><BR>"
else
adsrs.pagesize=advertlistnumber'每页显示的记录数
totalPut=adsrs.recordcount '记录总数
totalPage=adsrs.pagecount
MaxPerPage=adsrs.pagesize
if currentpage<1 then
currentpage=1
end if
if currentpage>totalPage then
currentpage=totalPage
end if
if currentPage=1 then
showContent
showpages
else
if (currentPage-1)*MaxPerPage<totalPut then
adsrs.move  (currentPage-1)*MaxPerPage
dim bookmark
bookmark=adsrs.bookmark '移动到开始显示的记录位置
showContent
showpages
end if
end if
adsrs.close
set adsrs=nothing
end if
adsconn.close
set adsconn=nothing

%>
 
          <%
sub showContent
i=0
do while not (adsrs.eof or err)
%>
           <div align="center">
    <center>
    <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#C0C0C0" width="100%" id="AutoNumber1">
      <tr>
        <td width="20%" bgcolor="#F5F5F5"><font color="#FF0000">&nbsp;广告ID：<%=adsrs("id")%> </font></td>
        <td bgcolor="#F5F5F5">&nbsp;名称：<input type="text" name="h" size="23" value="<%=adsrs("sitename")%>"></td>
        <td bgcolor="#F5F5F5" width="300">
       &nbsp;URL： 
       <input type="text" name="h" size="30" value="<%=adsrs("url")%>"></td>
        <td bgcolor="#F5F5F5" width="100" align="center">
        <%if adsrs("xslei")="txt" then%>
           <a href=javascript:opw('option.asp?id=<%=adsrs("id")%>&job=yulan','banner<%=adsrs("id")%>',400,200)>预览广告</a>
        <%else
        
        %>
            <a href=javascript:opw('option.asp?id=<%=adsrs("id")%>&job=yulan','banner<%=adsrs("id")%>',800,600)>预览广告</a>
       <%end if%>
　</td>
      </tr>
      <tr>
        <td width="20%" height="60">&nbsp;打开：<%= Ggdklx(adsrs("window"))%><br>&nbsp;显示：<%= Ggxslx(adsrs("xslei"))%><br>
        &nbsp;类型：<%= Ggwlx(adsrs("place"))%></td>
        <td height="60">&nbsp;加入时间：<font color=red><%=adsrs("regtime")%></font><br>&nbsp;最新显示：<font color=red><%=adsrs("time")%></font><br>
        &nbsp;最新点击：<font color=red><%=adsrs("lasttime")%></font></td>
        <td height="60" width="300">&nbsp;显示次数：<%call  Xscs()%><br>&nbsp;点击次数：<%call  Djcs()%><br>
        &nbsp;广 告 位：<%= Ggwm(adsrs("place"))%>  ID=<font color=red><%=adsrs("place")%></font></td>
        <td height="60" width="100" align="center">              <%
if adsrs("act")=1 then
%>
              <a href=javascript:opw('option.asp?id=<%=adsrs("id")%>&job=close','close<%=adsrs("id")%>',300,140)>暂停</a> 
              <%
else
%>
              <a href=javascript:opw('option.asp?id=<%=adsrs("id")%>&job=open','open<%=adsrs("id")%>',300,140)>激活</a> 
              <%end if%>
              　<a href=javascript:opw('option.asp?id=<%=adsrs("id")%>&job=delit','del<%=adsrs("id")%>',300,140)>删除</a></td>
      </tr>
      <tr>
        <td colspan="3" height="20">&nbsp;失效条件：<%call  Sxtj()%></td>
        <td height="20" width="100" align="center"><a href=javascript:opw('edit.asp?id=<%=adsrs("id")%>','<%=adsrs("id")%>',700,520)>修改详细设置</a></td>
      </tr>
      </table>
    </center>
</div>
          <%
i=i+1
if i>=MaxPerPage then exit do '循环时如果到尾部则先退出，如果记录达到页最大显示数，也退出
adsrs.movenext
loop
end sub 

sub showpages()
dim n
n=totalPage
%>
    
        <table border=0 width=100% cellpadding=2>
          <form method=post action=list.asp?type=<%=request.querystring("type")%>>
            <tr bgcolor=#ffffff> 
              <td align=right colspan=4>
 共<font color=red><%=totalput%></font>条、<font color=red><%=totalPage%></font>页,<font color=red><%=advertlistnumber%></font>条/页,第<font color=red><%=currentPage%></font>页
 
 <a href="?type=<%=request.querystring("type")%>&page=<%=currentPage-1%>">上一页</a> <a href="?type=<%=request.querystring("type")%>&page=<%=currentPage+1%>">下一页</a>

<%
response.write "<br> 转到：<select name='page' size=1>"
for i=1 to n
response.write "<option value="& i
if currentpage=i then
response.write " selected"
end if
response.write ">"& i &"</option>"
next
response.write "</select>&nbsp;<input type='submit'  value=' go '>"
%>
              </td>
            </tr>
          </form>
        </table>
     
<%
end sub

Sub Sxtj()

%>
<%
if adsrs("class")=1 then
%>
              点击<font color=red><%=adsrs("clicks")%></font>次 
              <%
elseif adsrs("class")=2 then
%>
              显示<font color=red><%=adsrs("shows")%></font>次 
              <%
elseif adsrs("class")=3 then
%>
              截止期<font color=red><%=adsrs("lasttime")%></font> 
              <%
elseif adsrs("class")=4 then
%>
              点击<font color=red><%=adsrs("clicks")%></font>次，显示<font color=red><%=adsrs("shows")%></font>次 
              <%
elseif adsrs("class")=5 then
%>
              点击<font color=red><%=adsrs("clicks")%></font>次，截止期<font color=red><%=adsrs("lasttime")%></font> 
              <%
elseif adsrs("class")=6 then
%>
              显示<font color=red><%=adsrs("shows")%></font>次，截止期<font color=red><%=adsrs("lasttime")%></font> 
              <%
elseif adsrs("class")=7 then
%>
              点击<font color=red><%=adsrs("clicks")%></font>次，显示<font color=red><%=adsrs("shows")%></font>次，截止期<font color=red><%=adsrs("lasttime")%></font> 
              <%
else
%>
              无限制条件 
<%
end if
%><%end sub

Sub Xscs()%>
 <font color=red><%=adsrs("show")%></font> (<a href=javascript:opw('listip.asp?id=<%=adsrs("id")%>&ip=sip','s<%=adsrs("id")%>',400,500)>显示记录</a>)
<%end sub

Sub Djcs()%>
 <font color=red><%=adsrs("click")%></font> (<a href=javascript:opw('listip.asp?id=<%=adsrs("id")%>&ip=cip','c<%=adsrs("id")%>',400,500)>点击记录</a>)
<%end sub

%><!--#include file="boot.asp"-->