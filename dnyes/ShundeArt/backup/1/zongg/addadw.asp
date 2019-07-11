<%@ LANGUAGE="VBSCRIPT" %>
<!--#include file=config.asp -->


<script language="javascript">
<!--
function isok(theform)
{
    if (theform.placename.value=="")
    {
        alert("请填写广告位标识！");
        theform.placename.focus();
        return (false);
    }

}
-->
</script>
<script language=JavaScript>
<!--

function opw(url,name, width, height) { //v2.0
window.open(url,name,''+'width='+width+',height='+height+',scrollbars=yes'+'');
}
//-->
</script>
<!--#include file="top.asp"-->
<center>
  <table border="0" align="center"  width=100%  cellpadding="0" cellspacing="0" height="1">
    <tr> 
           <td valign="top">
        <table border=0 width=98% cellspacing=0 cellpadding=0>
          <tr bgcolor=#ffffff align=center valign=top> 
            <td> 
              <table border=0 width=100% cellspacing=0 cellpadding=2 style="border-collapse: collapse" bordercolor="#111111">
				<form method="POST"  action=?job=add onSubmit="return isok(this)">

              <%
if request.querystring("job")="add" then

adsconn.open adsdata
 
set adsrs=server.createobject("adodb.recordset")

if  request("place")="0" then
adssql="select * from place "
adsrs.open adssql,adsconn,1,3
adsrs.AddNew
else
adssql="select * from place where place="&trim(request("place"))
adsrs.open adssql,adsconn,1,3
end if
adsrs(1) = trim(request("placename"))
adsrs(2)= trim(request("placelei"))
adsrs(3)= trim(request("placehei"))
adsrs(4)= trim(request("placewid"))
adsrs.update
adsrs.close
set adsrs=nothing
adsconn.close
set adsconn=nothing
response.redirect "?"

end if

if request.querystring("job")="del" then
if  isnumeric(request("place"))=true then
adsconn.open adsdata
set adsrs=server.createobject("adodb.recordset")
adssql="select * from place where place="&trim(request("place"))
adsrs.open adssql,adsconn,3,3
adsrs.delete
adsrs.close
set adsrs=nothing
adsconn.close
set adsconn=nothing
response.redirect "?"
end if
end if
%>
              <tr bgcolor=#ffffff> 
                <td>　</td>
                <td> <input type=hidden name=place value="0" >
                  <p align="center">广告位标识&nbsp;<input type=text name=placename size=30 maxlength=30>
                  <font color="#FF0000">15字以内 </font>高度<input type=text name=placehei value="60" size=3 maxlength=30> 
                  宽度<input type=text name=placewid value="468" size=3 maxlength=30>&#12288;类型 
 <%Call Ggwlei(1)%>&nbsp; 
                  <input type=submit value=新增广告位 name=B1></p>
                <p align="center">
                  <font color="#808080">*** 请尽量不要使用相同的广告位标识，高度、宽度主要应用于弹出窗口大小、滚动区域设置，请适当设置，不可为空</font></td>
              </tr>
             
            </form>
          </table>
        </td>
      </tr>
    </table>
  </td>
    </tr>
  </table>
  
  <hr color="#808080" size="1">
<p> <font color="#FF0000">广告插入方法：</font>将下表内容放到预定广告位置，并将其中的<font color="#FF0000">广告位ID</font>对应正确 
<font color="#808080">请在广告位列表中查看</font><font color="#FF0000">广告位ID</font><br><input type="text" name="T1" size="100" value="<script language=javascript src=<%=DqUrl()%>/ad.asp?i=广告位ID></script>"></p>
<p align="center"> <font color=red><b>已有广告位列表</b></font></p>
<p align=left><font color="#808080">*** 请在高度、宽度中输入合适的<b>数字或百分比或为空自动</b></font></p>
<div align="center">
  <center>
  <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#C0C0C0" width="100%" id="AutoNumber1">
    <tr>
      <td width="80" align="center" height="20" bgcolor="#F5F5F5">
      <font color="#FF0000">广告位ID</font></td>
      <td align="center" height="20" bgcolor="#F5F5F5">广告位标识</td>
      <td align="center" height="20" bgcolor="#F5F5F5">高度</td>
      <td align="center" height="20" bgcolor="#F5F5F5">宽度</td>
      <td width="20%" align="center" height="20" bgcolor="#F5F5F5">广告位显示方式</td>
      <td width="30%" align="center" height="20" bgcolor="#F5F5F5">操 作</td>
    </tr>
    
   <%
adsconn.open adsdata
set adsrs=server.createobject("adodb.recordset")
adssql="select * from place"
adsrs.open adssql,adsconn,1,1
do while not adsrs.eof 

%><form method="POST" action="?job=add"  onSubmit="return isok(this)">
    <tr>
     
      <td width="80" align="center"><font color=red><%=adsrs(0)%></font><input type=hidden name=place value="<%=adsrs(0)%>" >　</td>
      <td align="center"><input type=text name=placename value="<%=adsrs(1)%>" size=30 maxlength=30></td>
      <td align="center">
      <input type=text name=placehei value="<%=adsrs(3)%>" size=3 maxlength=30></td>
      <td align="center">
      <input type=text name=placewid value="<%=adsrs(4)%>" size=3 maxlength=30></td>
      <td width="20%" align="center">
     <%Call Ggwlei(adsrs("placelei"))%>&nbsp; 
</td>
      <td width="30%" align="center"><input type="submit" value="修改" name="B1">&nbsp;&nbsp; 
      <a href=?job=del&place=<%=adsrs(0)%>>删除</a>&nbsp; <a href=list.asp?type=place&place=<%=adsrs(0)%>>已有广告条</a> <a href=javascript:opw('option.asp?id=<%=adsrs(0)%>&job=yulanggw','',800,600)>预览广告</a></td>
    </tr>
    </form>
    <%adsrs.movenext
      loop
      adsrs.close
      set adsrs=nothing
adsconn.close
set adsconn=nothing
      %>
  </table>
  </center>
</div>
　
  <p align="left">
  <p align="left"><hr color="#808080" size="1">
<p align="left"><font color="#FF0000"><a name="说明">广告位显示方式说明</a>：</font></p>
<center>
  </p>
  <ul>
    <li><p align="left">
    页内嵌入循环：就是将广告位直接置入某页面一固定位置，并在同一位置循环显示广告位中的所有正常广告条，这样，每刷新一次就会更替显示一个新的广告条</p>
    </li>
    <li><p align="left">上下排列置入：从上到下竖排广告位中的所有正常广告条</p>
    </li>
    <li><p align="left">左右排列置入：从左到右横排广告位中的所有正常广告条</p>
    </li>
    <li><p align="left">向上滚动置入：向上滚动显示广告位中的所有正常广告条</p>
    </li>
    <li><p align="left">向左滚动置入：向左滚动显示广告位中的所有正常广告条</p>
    </li>
    <li><p align="left">弹出多个窗口：页面打开时同时弹出多个窗口，每个窗口内显示一个广告条，弹出数量跟该广告位中的正常广告条数一致</p>
    </li>
    <li><p align="left">
    循环弹出窗口：页面打开时同时弹出一个窗口，在同一窗口内循环显示广告位中的正常广告，这样，每刷新一次就会在弹出窗口中更替显示一个新的广告条</p>
    </li>
  </ul>
  <!--#include file="boot.asp"-->
<p><br><br>
       

        　</p>