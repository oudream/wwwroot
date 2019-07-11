<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<title><%=Forum_info(0)%>--管理页面</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta NAME=GENERATOR Content="Microsoft FrontPage 3.0" CHARSET=GB2312>
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0" marginheight="0" marginwidth="0" bgcolor="#DDEEFF">
<%
	if not master or session("flag")="" then
		Errmsg=Errmsg+"<br>"+"<li>本页面为管理员专用，请<a href=admin_index.asp target=_top>登陆</a>后进入。<br><li>您没有管理本页面的权限。"
		call dvbbs_Error()
	else
		dim body
		dim dadadas
		dim country,city,num
		dim ip1,ip2,oldip1,oldip2
		dim esip,str1,sip,str2,str3,str4
		dim s1,s21,s2,s31,s3,s4
		set rs= server.createobject ("adodb.recordset")
		call main()
		set rs=nothing
		conn.close
		set conn=nothing
	end if

sub main()
%>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th width="100%" colspan=5>IP数据源管理－－<a href="?action=add"><font color=#FFFFFF>添加ＩＰ地址</font></a></th>
</tr>
<tr> 
<td width="20%" class=forumrow>注意事项</td>
<td width="80%" class=forumrow colspan=4>如果需要添加IP数据来源请直接添加，如果添加的来源在数据库中已经存在，将提示您是否进行修改，数据库中尚没有的记录将直接添加，您也可以直接对现有的数据进行管理操作。</td>
</tr>
<%
	if request("action") = "add" then 
		call addip()
	elseif request("action")="edit" then
		call editip()
	elseif request("action")="savenew" then
		call savenew()
	elseif request("action")="savedit" then
		call savedit()
	elseif request("action")="del" then
		call del()
	elseif request("action")="query" then
		call ipinfo()
        elseif request("action")="upip" then
                call upip()
        elseif request("action")="saveip" then
                call saveip()
	else
		call ipquery()
	end if
	response.write body
%>
</table>
<%
end sub

sub addip()
%>
<form action="?action=savenew" method=post>
<tr> 
    <th colspan="5" align=left>修改 IP 地址</th>
</tr>
<tr> 
<td width="20%" class=forumrow>起始 I P</td>
<td width="80%" class=forumrow colspan=4><input type="text" name="ip1" size="30"></td>
</tr>
<tr> 
<td width="20%" class=forumrow>结尾 I P</td>
<td width="80%" class=forumrow colspan=4><input type="text" name="ip2" size="30"></td>
</tr>
<tr> 
<td width="20%" class=forumrow>来源国家</td>
<td width="80%" class=forumrow colspan=4><input type="text" name="country" size="30"></td>
</tr>
<tr> 
<td width="20%" class=forumrow>来源城市</td>
<td width="80%" class=forumrow colspan=4><input type="text" name="city" size="30"></td>
</tr>
<tr> 
<td width="20%" class=forumrow></td>
<td width="80%" class=forumrow colspan=4><input type="submit" name="Submit" value="提 交"></td>
</tr>
</form>
<%
end sub

sub editip()
ip1=request("ip1")
ip2=request("ip2")
sql="select ip1,ip2,country,city from address where ip1="&ip1&" and ip2="&ip2&""
rs.open sql,conn,1,1
%>
<form action="?action=savedit" method=post>
<tr> 
    <th colspan="5" align=left>修改 IP 地址</th>
</tr>
<tr> 
<td width="20%" class=forumrow>起始 I P</td>
<td width="80%" class=forumrow colspan=4><input type="text" name="ip1" size="30" value="<%=deaddr(ip1)%>"></td>
</tr>
<tr> 
<td width="20%" class=forumrow>结尾 I P</td>
<td width="80%" class=forumrow colspan=4><input type="text" name="ip2" size="30" value="<%=deaddr(ip2)%>"></td>
</tr>
<tr> 
<td width="20%" class=forumrow>来源国家</td>
<td width="80%" class=forumrow colspan=4><input type="text" name="country" size="30" value="<%=rs("country")%>"></td>
</tr>
<tr> 
<td width="20%" class=forumrow>来源城市</td>
<td width="80%" class=forumrow colspan=4><input type="text" name="city" size="30" value="<%=rs("city")%>"></td>
</tr>
<input type="hidden" name="oldip1" value="<%=request("ip1")%>">
<input type="hidden" name="oldip2" value="<%=request("ip2")%>">
<tr> 
<td width="20%" class=forumrow></td>
<td width="80%" class=forumrow colspan=4><input type="submit" name="Submit" value="提 交"></td>
</tr>
</form>
<%
end sub

sub savenew()
if request.form("ip1")="" then
body="<tr><td colspan=5 class=forumrow>请填写IP地址！</td></tr>"
exit sub
end if
if request.form("ip2")="" then
body="<tr><td colspan=5 class=forumrow>请填写IP地址！</td></tr>"
exit sub
end if
if request.form("country")="" or request.form("city")="" then
body="<tr><td colspan=5 class=forumrow>国家和城市必须填写其一！</td></tr>"
exit sub
end if
ip1=enaddr(request.form("ip1"))
ip2=enaddr(request.form("ip2"))
country=request.form("country")
city=request.form("city")
sql="select ip1,ip2,country,city from address where ip1<="&ip1&" and ip2>="&ip2&""
rs.open sql,conn,1,3
if rs.eof and rs.bof then
rs.AddNew 
rs("ip1")=ip1
rs("ip2")=ip2
rs("country")=country
rs("city")=city
rs.update
body="<tr><td colspan=5 class=forumrow>添加成功，<a href=?>请继续其他操作</a>。</td></tr>"
else
body="<tr><td colspan=5 class=forumrow>添加失败，数据已存在，请到搜索中搜索该ip地址并进行修改。</td></tr>"
end if
rs.close
end sub

sub savedit()
oldip1=int(request.form("oldip1"))
oldip2=int(request.form("oldip2"))
ip1=enaddr(request.form("ip1"))
ip2=enaddr(request.form("ip2"))
country=request.form("country")
city=request.form("city")
sql="select ip1,ip2,country,city from address where ip1="&oldip1&" and ip2="&oldip2&""
rs.open sql,conn,1,3
if not(rs.eof and rs.bof) then
rs("ip1")=ip1
rs("ip2")=ip2
rs("country")=country
rs("city")=city
rs.update
body="<tr><td colspan=5 class=forumrow>IP地址修改成功！</td></tr>"
else
body="<tr><td colspan=5 class=forumrow>IP地址修改失败！</td></tr>"
end if
rs.close
end sub

sub del()
ip1=request("ip1")
ip2=request("ip2")
sql="delete from address where ip1="&ip1&" and ip2="&ip2&""
conn.Execute(sql)
body="<tr><td colspan=5 class=forumrow>删除成功，<a href=?>请继续其他操作</a>。</td></tr>"
end sub

sub ipinfo()
	dim currentpage,page_count,Pcount
	dim totalrec,endpage
	currentPage=request("page")
	if currentpage="" or not isInteger(currentpage) then
		currentpage=1
	else
		currentpage=clng(currentpage)
		if err then
			currentpage=1
			err.clear
		end if
	end if
	country=trim(request("country"))
	city=trim(request("city"))
	if request("ip")="" and country="" and city="" then
	sql="select ip1,ip2,country,city from address order by ip1 desc,ip2 desc"
	else
	sql="select ip1,ip2,country,city from address where "
	if request("ip")<>"" then
		num=enaddr(request("ip"))
		sql = sql&" ip1 <="&num&" and ip2 >="&num&" order by ip1 desc,ip2 desc"
	else
		if country<>"" then sql=sql&" country like '%"&country&"%'"
		if city<>"" then sql=sql&" city like '%"&city&"%'"
		sql=sql&" order by ip1 desc,ip2 desc"
	end if
	end if
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
	call addip()
	else
%>
<tr> 
    <th colspan="5" align=left>IP 地址库搜索结果</th>
</tr>
  <tr align="center"> 
    <td width="20%" class="forumHeaderBackgroundAlternate"><B>起始 IP</B></td>
    <td width="20%" class="forumHeaderBackgroundAlternate"><B>结尾 IP</B></td>
    <td width="18%" class="forumHeaderBackgroundAlternate"><B>国 家</B></td>
    <td width="30%" class="forumHeaderBackgroundAlternate"><B>城 市</B></td>
    <td width="12%" class="forumHeaderBackgroundAlternate"><B>操 作</B></td>
  </tr>
<%
		rs.PageSize = Cint(Forum_Setting(11))
		rs.AbsolutePage=currentpage
		page_count=0
		totalrec=rs.recordcount
		while (not rs.eof) and (not page_count = Cint(Forum_Setting(11)))
%>
    <tr>
    <td width="20%" class=forumrow><%=deaddr(rs("ip1"))%></td>
    <td width="20%" class=forumrow><%=deaddr(rs("ip2"))%></td>
    <td width="18%" class=forumrow><%=rs("country")%></td>
    <td width="30%" class=forumrow><%=rs("city")%></td>
    <td width="12%" align="center" class=forumrow><a href="?action=edit&ip1=<%=rs("ip1")%>&ip2=<%=rs("ip2")%>">编辑</a>|<a href="?action=del&ip1=<%=rs("ip1")%>&ip2=<%=rs("ip2")%>">删除</a></td>
    </tr>
<%
		page_count = page_count + 1
		rs.movenext
		wend
%>
<tr><td colspan=5 class=forumrow align=center>分页：
<%Pcount=rs.PageCount
	if currentpage > 4 then
	response.write "<a href=""?page=1&action=query&country="&country&"&city="&city&"&ip="&request("ip")&""">[1]</a> ..."
	end if
	if Pcount>currentpage+3 then
	endpage=currentpage+3
	else
	endpage=Pcount
	end if
	for i=currentpage-3 to endpage
	if not i<1 then
		if i = clng(currentpage) then
        response.write " <font color=red>["&i&"]</font>"
		else
        response.write " <a href=""?page="&i&"&action=query&country="&country&"&city="&city&"&ip="&request("ip")&""">["&i&"]</a>"
		end if
	end if
	next
	if currentpage+3 < Pcount then 
	response.write "... <a href=""?page="&Pcount&"&action=query&country="&country&"&city="&city&"&ip="&request("ip")&""">["&Pcount&"]</a>"
	end if
%>
</td>
</tr>
<%
	end if
	rs.close
end sub

sub ipquery()
%>
<form action="?action=query" method = post>
<tr> 
    <th colspan="5" align=left>IP 地址库搜索</th>
</tr>
  <tr> 
    <td width="20%" class=forumrow>IP 地址：</td>
    <td width="80%" class=forumrow colspan=4> 
      <input type="text" name="ip" size=40>
    </td>
  </tr>
  <tr> 
    <td width="20%" class=forumrow>国 家：</td>
    <td width="80%" class=forumrow colspan=4> 
      <input type="text" name="country" size=40>
    </td>
  </tr>
  <tr> 
    <td width="20%" class=forumrow>城 市：</td>
    <td width="80%" class=forumrow colspan=4> 
      <input type="text" name="city" size=40>
    </td>
  </tr>
  <tr> 
    <td width="20%" class=forumrow></td>
    <td width="80%" class=forumrow colspan=4> 
       <input type="submit" name="Submit" value="搜 索" class="button">
    </td>
  </tr>
</form>

<%end sub

'*************************************
'更换IP库
'SUNWIN  2002-9-23
'*************************************

sub upip()
%>

    <tr  id=tdbg>
      <th colspan="3" align=left><a href=?action=upip><font color=#FFFFFF>IP库导入功能</font></a></th>
    </tr>
  <tr >
    <td width="100%" height="1" colspan="3" align="center" class=forumrow>
      <p align="center">注意:请认真按要求填写以下内容!</td>
  </tr>
 <form action="?action=saveip" method=post>
  <tr >
    <td width="20%" height="1" align="right" valign="middle" class=forumrow>导入数据库名：</td>
    <td width="60%" height="1" class=forumrow><input type="text" name="ipname" size="30">如：IP.mdb　或　IP/IP.mdb</td>
    <td width="20%" height="3" rowspan="2" class=forumrow>更换IP库前，请查看新ＩＰ数据库的名称与记录ＩＰ的表名，并填写到左边表单中。</td>
  </tr>
  <tr >
    <td  class=forumrow align="right" valign="middle">导入IP库表名：</td>
    <td class=forumrow><input type="text" name="iptable" size="30" value="address">如：address</td>
  </tr>
  <tr >
    <td class=forumrow align="right" valign="middle"></td>
    <td class=forumrow colspan="2"><input type="submit" name="Submit" value="更新"></td>
  </tr>
</form>
    <tr >
       <th colspan="3" align=left><a href=?action=upip ><font color=#FFFFFF>IP库导入功能</font></a></th>
    </tr>
<% end sub

sub saveip()
%>  <tr>
    <td width="100%" height="1" colspan="5" align="left" class=forumrow>
<%
dim ipname,iptable,TCONN

  	ipname=Checkstr(trim(request("ipname")))
   	iptable=Checkstr(trim(request.form("iptable")))
if ipname="" then
     		ErrMsg=ErrMsg+"<Br>"+"<li>请填写导入的数据库名"
   		response.write Errmsg
            exit sub
end if
if iptable="" then
     		ErrMsg=ErrMsg+"<Br>"+"<li>请填写导入的IP库表名"
   		response.write Errmsg
            exit sub
end if
'删除旧的IP数据
	conn.execute("delete from address")

'*************************************
'取出新的数据进行更新
'*************************************
	'on error resume next
       	Set tconn = Server.CreateObject("ADODB.Connection")
	tconn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(ipname)
        dim trs,nrs
		set trs=tconn.execute(" select * from "&iptable&" ")
		do while not trs.eof
     conn.execute("insert into address (ip1,ip2,country,city) values ('"&trs("ip1")&"','"&trs("ip2")&"','"&trs("country")&"','"&trs("city")&"')")

		trs.movenext
		loop
     trs.close
     set trs=nothing
	if err.number<>0 then
		ErrMsg=ErrMsg+"<Br>"+"<li>数据库操作失败，请以后再试，"&err.Description
		err.clear
		response.write Errmsg
		exit sub
	else
		response.write "数据导入成功！<br>"
	end if

%></td>
  </tr>
<%
end sub%>
<%

function enaddr(sip)
 esip=cstr(sip)
 str1=left(sip,cint(instr(sip,".")-1))
 sip=mid(sip,cint(instr(sip,"."))+1)
 str2=left(sip,cint(instr(sip,"."))-1)
 sip=mid(sip,cint(instr(sip,"."))+1)
 str3=left(sip,cint(instr(sip,"."))-1)
 str4=mid(sip,cint(instr(sip,"."))+1)
 enaddr=cint(str1)*256*256*256+cint(str2)*256*256+cint(str3)*256+cint(str4)-1
end function
function deaddr(sip)
sip=sip+1
s1=int(sip/256/256/256)
s21=s1*256*256*256
s2=int((sip-s21)/256/256)
s31=s2*256*256+s21
s3=int((sip-s31)/256)
s4=sip-s3*256-s31
deaddr=cstr(s1)+"."+cstr(s2)+"."+cstr(s3)+"."+cstr(s4)
end function
%>
<script language="JavaScript">
<!--
function CheckAll(form)  {
  for (var i=0;i<form.elements.length;i++)    {
    var e = form.elements[i];
    if (e.name != 'chkall')       e.checked = form.chkall.checked; 
   }
  }
//-->
</script>
