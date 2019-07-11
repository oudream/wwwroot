<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<%
'=========================================================
' File: look_ip.asp
' Version:5.0
' Date: 2002-9-28
' Script Written by satan
'=========================================================
' Copyright (C) 2001,2002 AspSky.Net. All rights reserved.
' Web: http://www.aspsky.net,http://www.dvbbs.net
' Email: info@aspsky.net,eway@aspsky.net
'=========================================================

dim canlookip,canlockip
dim lockid
canlookip=false
canlockip=false
if (master or superboardmaster or boardmaster) and Cint(GroupSetting(30))=1 then
	canlookip=true
else
	canlookip=false
end if
if UserGroupID>3 and Cint(GroupSetting(30))=1 then
	canlookip=true
end if
if FoundUserPer and Cint(GroupSetting(30))=1 then
	canlookip=true
elseif FoundUserPer and Cint(GroupSetting(30))=0 then
	canlookip=false
end if

if (master or superboardmaster or boardmaster) and Cint(GroupSetting(31))=1 then
	canlockip=true
else
	canlockip=false
end if
if UserGroupID>3 and Cint(GroupSetting(31))=1 then
	canlockip=true
end if
if FoundUserPer and Cint(GroupSetting(31))=1 then
	canlockip=true
elseif FoundUserPer and Cint(GroupSetting(31))=0 then
	canlockip=false
end if
stats="查看用户IP信息"
call nav()
call head_var(1)


if founderr then
	call dvbbs_error()
else
	if request("action")="setlockip" then
		call Setlockip()
	elseif request("action")="unlock" then
		call unlock()
	else
		call lookip()
	end if
	if founderr then call dvbbs_error()
	call activeonline()
end if
call footer()

sub lookip()
if not canlookip then
	Errmsg=Errmsg+"<br><li>您没有执行此操作的权限。"
	founderr=true
	exit sub
end if
%>
<table class=tableborder1 cellspacing="1" cellpadding="3" align="center">
  <tr align=center> 
    <th height=25>查看 <%=request("ip")%> 的来源</th>
  </tr>
  <tr><td height=25 class=tablebody1><blockquote><%=lookaddress(replace(request("ip"),"'",""))%></blockquote></td></tr>
  <%if canlockip then%>
  <tr><td height=25 class=tablebody2 align=center><B>管理操作</B>：
  <%if GetLockIP(replace(request("ip"),"'","")) then%>
  <a href="?action=unlock&boardid=<%=request("boardid")%>&id=<%=lockid%>">该用户IP已被锁定，解除锁定
  <%else%>
  <a href="?action=setlockip&ip=<%=request("ip")%>&boardid=<%=request("boardid")%>">限制该IP不允许访问</a>
  <%end if%>
  </td></tr>
  <%end if%>
</table>
<%
end sub

sub Setlockip()
if not canlockip then
	Errmsg=Errmsg+"<br><li>您没有执行此操作的权限。"
	founderr=true
	exit sub
end if
if request("reaction")="yes" then
	dim sip,str1,str2,str3,str4,num_1,num_2
	sip=cstr(request.form("ip1"))
	str1=left(sip,cint(instr(sip,".")-1))
	sip=mid(sip,cint(instr(sip,"."))+1)
	str2=left(sip,cint(instr(sip,"."))-1)
	sip=mid(sip,cint(instr(sip,"."))+1)
	str3=left(sip,cint(instr(sip,"."))-1)
	str4=mid(sip,cint(instr(sip,"."))+1)
	num_1=cint(str1)*256*256*256+cint(str2)*256*256+cint(str3)*256+cint(str4)-1

	sip=cstr(request.form("ip2"))
	str1=left(sip,instr(sip,".")-1)
	sip=mid(sip,instr(sip,".")+1)
	str2=left(sip,instr(sip,".")-1)
	sip=mid(sip,instr(sip,".")+1)
	str3=left(sip,instr(sip,".")-1)
	str4=mid(sip,instr(sip,".")+1)
	num_2=cint(str1)*256*256*256+cint(str2)*256*256+cint(str3)*256+cint(str4)-1

	set rs = server.CreateObject ("adodb.recordset")
	sql="select * from LockIP"
	rs.open sql,conn,1,3
	rs.addnew
	rs("ip1")=num_1
	rs("ip2")=num_2
	rs("sip1")=request.form("ip1")
	rs("sip2")=request.form("ip2")
	rs.update
	rs.close
	set rs=nothing
	sucmsg="<li>您选择IP已经被禁止访问论坛。"
	call dvbbs_suc()
else
	dim userip,ips,GetIp1,GetIp2
	if request("ip")<>"" then
		userip=request("ip")
		ips=Split(userIP,".")
		GetIp1=ips(0)&"."&ips(1)&"."&ips(2)&".1"
		GetIp2=ips(0)&"."&ips(1)&"."&ips(2)&".255"
	else 
		userip=""
		GetIp1=""
		GetIp2=""
	end if
%>
<table class=tableborder1 cellspacing="1" cellpadding="3" align="center">
  <tr align=center> 
    <th height=25>锁定 <%=request("ip")%> 的来源</th>
  </tr>
  <tr><td height=25 class=tablebody1><blockquote><%=lookaddress(replace(request("ip"),"'",""))%></blockquote></td></tr>
<FORM METHOD=POST ACTION="?action=setlockip&boardid=<%=request("boardid")%>">
<input type=hidden name="reaction" value="yes">
    <tr><td height=40 class=tablebody1>
  <B>起始I&nbsp;P</B>：<input type="text" name="ip1" size="30" value="<%=GetIp1%>">&nbsp;&nbsp;<B>结尾I&nbsp;P</B>：<input type="text" name="ip2" size="30" value="<%=GetIp2%>">&nbsp;&nbsp;<input type="submit" name="Submit" value="提 交">
  </td></tr>
  </FORM>
</table>
<%
end if
end sub

sub unlock()
if not canlockip then
	Errmsg=Errmsg+"<br><li>您没有执行此操作的权限。"
	founderr=true
	exit sub
end if
if not isnumeric(request("id")) then
	Errmsg=Errmsg+"<br><li>错误的IP参数。"
	founderr=true
	exit sub
end if
conn.execute("delete from lockip where id="&request("id"))
sucmsg="<li>成功解除了该IP的限定"
call dvbbs_suc()
end sub

function lookaddress(sip)
dim str1,str2,str3,str4
dim num
dim irs
if isnumeric(left(sip,2)) then
	if sip="127.0.0.1" then sip="192.168.0.1"
	str1=left(sip,instr(sip,".")-1)
	sip=mid(sip,instr(sip,".")+1)
	str2=left(sip,instr(sip,".")-1)
	sip=mid(sip,instr(sip,".")+1)
	str3=left(sip,instr(sip,".")-1)
	str4=mid(sip,instr(sip,".")+1)
	if isNumeric(str1)=0 or isNumeric(str2)=0 or isNumeric(str3)=0 or isNumeric(str4)=0 then

	else
		num=cint(str1)*256*256*256+cint(str2)*256*256+cint(str3)*256+cint(str4)-1
		sql="select country,city from address where ip1 <="&num&" and ip2 >="&num
		set irs=conn.execute(sql)
		if irs.eof and irs.bof then 
			lookaddress="数据库未保存该IP信息"
		else
			do while not irs.eof
			lookaddress=lookaddress & "<br>" &irs(0) & irs(1)
			irs.movenext
			loop
		end if
		irs.close
		set irs=nothing
	end if
else
	lookaddress="数据库未保存该IP信息"
end if
end function

function getLockIP(sip)
	dim str1,str2,str3,str4
	dim num
	getLockIP=false
	if isnumeric(left(sip,2)) then
		str1=left(sip,instr(sip,".")-1)
		sip=mid(sip,instr(sip,".")+1)
		str2=left(sip,instr(sip,".")-1)
		sip=mid(sip,instr(sip,".")+1)
		str3=left(sip,instr(sip,".")-1)
		str4=mid(sip,instr(sip,".")+1)
		if isNumeric(str1)=0 or isNumeric(str2)=0 or isNumeric(str3)=0 or isNumeric(str4)=0 then
	
		else
			num=cint(str1)*256*256*256+cint(str2)*256*256+cint(str3)*256+cint(str4)-1
			sql="select * from LockIP where ip1 <="&num&" and ip2 >="&num&""
			set rs=conn.execute(sql)
			if not (rs.eof and rs.bof) then
				getLockIP=true
				lockid=rs("id")
			end if
			set rs=nothing
		end if
	end if
end function
%>