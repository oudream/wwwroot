<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="md5.asp"-->
<!--#include file="ChkURL.asp"-->

<%if request.cookies(Forcast_SN)("key")<>"" then%>
<!--#include file="ChkUser.asp"-->
<%end if
author=trim(request("guestname"))
reviewid=request("reviewid")
password=md5(trim(request("password")))
if Instr(request("guestname"),"=")>0 or Instr(request("guestname"),"%")>0 or Instr(request("guestname"),chr(32))>0 or Instr(request("guestname"),"?")>0 or Instr(request("guestname"),"&")>0 or Instr(request("guestname"),";")>0 or instr(request("guestname"),",")>0 or Instr(request("guestname"),"'")>0 or Instr(request("guestname"),",")>0 or Instr(request("guestname"),chr(34))>0 or Instr(request("guestname"),chr(9))>0 or Instr(request("guestname"),"��")>0 or Instr(request("guestname"),"$")>0 then
	Show_Err("�Բ�����������û����а����зǷ��ַ���<br><br><a href='javascript:history.back()'>����</a>")
	response.end
end if
Set rs=Server.CreateObject("Adodb.RecordSet")
rs.Open "Select * from "& db_User_Table &" where "& db_User_Name &"='"&author&"'",ConnUser
if not rs.EOF then
	if password<>rs(db_User_Password) then			''��֤�޸�������
		if request.cookies(Forcast_SN)("key")<>"super" then 	''�ж��Ƿ��ǳ����û����޸�,��������ʾ����
			Show_Err("�Բ��������������<br><br><a href='javascript:history.back()'>����</a>")
			Response.End 
		end if
	end if
	user=true
end if
rs.close
set rs=nothing

shengfen=checkstr(request.form("shengfen"))
from=checkstr(request.form("from"))
homepage=checkstr(request.form("homepage"))
reviewip=checkstr(Request.serverVariables("REMOTE_ADDR"))
oicq=checkstr(request.form("oicq"))
face=checkstr(request.form("face"))
content=trim(htmlencode1((request.form("content"))))
content=replace(content,"<p> ","")
content=replace(content,"<P> ","")
qqh=(request.form("qqh"))

dim byte1
byte1=split(byteType,"|")

for i=0 to ubound(byte1)
	content=replace(content,trim(byte1(i)),"***")
next
Response.cookies(Forcast_SN)("content")=content

if Instr(request("content"),"'")>0 or Instr(request("content"),"script")>0 or Instr(request("content"),"onClick")>0  or Instr(request("content"),"onload")>0 then
	Show_Err("�Բ�����������������ݰ����зǷ��ַ���<br><br><a href='javascript:history.back()'>����</a>")
	Response.End 
end if

if content="" then
	Show_Err("�������ݲ���Ϊ�գ��������������ݣ�<br><br><a href='javascript:history.back()'>����</a>")
	Response.End
end if

if user=true then
	set rs=server.createobject("adodb.recordset")
	sql="select * from "& db_User_Table &" where "& db_User_Name &"='"&author&"'"
	rs.open sql,ConnUser,1,3
	email=rs(db_User_Email)
	if UserTableType="FeiTium" then
		face=rs(db_User_Face)
	else
		if UserTableType="Dvbbs" then			''�ж��Ƿ�������̳
			face=(BbsPath+rs(db_User_Face))		''ע���û�ͷ��λ����̳·��
		end if
	end if
	rs.close
	set rs=nothing
else
	email=trim(request.form("email"))
end if

if reviewid="" then
	set rs=server.createobject("adodb.recordset")
	sql="select * from "& db_Review_Table &"" 
	rs.open sql,conn,1,3
	rs.addnew
	rs("author")=author
	rs("content")=content
	rs("shengfen")=shengfen
	rs("[from]")=from			'����������
	rs("homepage")=homepage
	rs("reviewip")=reviewip
	rs("email")=email
	rs("oicq")=oicq
	rs("face")=face
	rs("updatetime")=now()
	rs("NewsID")=qqh
	if reviewcheck=1 then		'��������Ĭ�����״̬�����������״̬������
		rs("passed")=1
	else
		rs("passed")=0
	end if
	rs.update
	rs.close
	set rs=nothing
else
	set rs=server.createobject("adodb.recordset")
	sql="select * from "& db_Review_Table &" where reviewid="&reviewid 
	rs.open sql,conn,1,3
	if not rs.eof then
		rs("author")=author
		rs("content")=content
		rs("shengfen")=shengfen
		rs("[from]")=from	'����������
		rs("homepage")=homepage
		rs("reviewip")=reviewip
		rs("email")=email
		rs("oicq")=oicq
		if user=true then			''�ж����޸������ʺ��Ƿ���ע���û�,��������ȡͷ��λ��
			rs("face")=face
		end if
		rs("updatetime")=now()
		
		rs("NewsID")=qqh
		rs.update
	end if
	rs.close
	set rs=nothing
end if
%>
<html>
<head>
<title>���Գɹ�__<%=jjgn%></title>
<meta http-equiv=refresh content="1;url=guestbook.asp">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="news.css" rel=stylesheet type=text/css>
</head>
<body bgcolor="#FFFFFF" text="#000000" topmargin="0">
<!--#include file="top.asp"-->
<table width="760" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
	<tr valign="top"> 
		<td>
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td height="3" ><img src="images/kb.gif" width="9" height="3"></td>
				</tr>
				<tr bgcolor="CDCCCC">
					<td height="25" background="images/menu-guestbook.gif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��Ŀ����&nbsp;&nbsp;<a class=daohang href="./" >��վ��ҳ</a>&gt;&gt;<a href="guestbook.asp" class="daohang">�������</a>&gt;&gt;���Գɹ�</td>
				</tr>
				<tr bgcolor="CDCCCC">
					<td background="IMAGES/menu-guest-l.gif" bgcolor="#FFFFFF">
						<div align="center">
							<a class=daohang href="./" >
								<script language=javascript src=./zongg/ad.asp?i=13></script></script>
							</a>
							<br>
						</div>
					</td>
				</tr>
				<tr>
					<td height="25" valign="top" background="IMAGES/menu-l-guest.gif" bgcolor="cdcccC"><div align="center"></div></td>
				</tr>
			</table>
		</td>
	</tr>
	<tr valign="top"> 
		<td background="IMAGES/menu-guest-l.gif">
			<table width="98%" height="100" border="0" align="center" cellpadding="0" cellspacing="0">
				<tr>
					<td align="center">
						<table border="0" cellspacing="0" width="30%" bordercolorlight="#FFFFFF" bordercolordark="#FFFFFF" cellpadding="0">
							<tr> 
								<td width="100%" height="20"><p align="center"><font color=red><b>���Գɹ�&nbsp;&nbsp;&nbsp;&nbsp;</b></font><a class=daohang href="guestbook.asp">����</a>
									</p>
								</td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr valign="top">
		<td height="19" background="IMAGES/menu-guest-t.gif"></td>
	</tr>
</table>
<!--#include file="bottom.asp"-->
<%Response.cookies(Forcast_SN)("content")=""
'response.redirect "guestbook.asp"%>