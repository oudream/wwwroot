<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!--#include file="inc/usercp.asp"-->
<%Response.Buffer = true%>
<%
'=========================================================
' File: usermanager.asp
' Version:5.0
' Date: 2002-9-11
' Script Written by satan
'=========================================================
' Copyright (C) 2001,2002 AspSky.Net. All rights reserved.
' Web: http://www.aspsky.net,http://www.dvbbs.net
' Email: info@aspsky.net,eway@aspsky.net
'=========================================================

stats="���������ҳ"
call nav()
if not founduser then
  	errmsg=errmsg+"<br>"+"<li>��û��<a href=login.asp target=_blank>��¼</a>��"
	call head_var(1)
	call dvbbs_error()
else
	call cp_var()
	call main()
	call activeonline()
end if
call footer()
sub main
	dim userclass
	set rs=server.createobject("adodb.recordset")
	sql="select * from [User] where userid="&userid
	rs.open sql,conn,1,1
	userclass=rs("userclass")


response.write "<table cellpadding=0 cellspacing=6 width="&Forum_body(12)&" align=center >"&_
               "<tr  align=center ><td  width=28% valign=top>"&_
               "<table align=center style=""width:100%"" height=100% cellspacing=1 cellpadding=6 class=tableborder1>"&_
               "<tr><th height=25>�û�ͷ��</th></tr>"&_
               "<tr align=center><td class=tablebody1>"&_
               "<img src="""&htmlencode(rs("face"))&""" width="&rs("width")&" height="&rs("height")&" align=absmiddle >"&_
               "</td></tr>"&_
               "<tr><th height=25>������Ϣ</th></tr>"&_
               "<tr><td align=left class=tablebody1 valign=top>" 
if trim(rs("title"))<>"" then
response.write "�û�ͷ�Σ� "&htmlencode(RS("title"))&"<br>"
end if

response.write "�û��ȼ��� "& rs("userclass") &"<br>"

if rs("UserGroup")<>"" then
response.write "�û����ɣ� "&rs("UserGroup")&"<br>"
end if

response.write "�û��Ƹ��� "&rs("userwealth")&"<br>"&_
               "�û����飺 "&rs("userEP")&"<br>"&_
               "�û������� "&rs("userCP")&"<br>"&_
               "���������� "&rs("userisbest")&"<br>"&_
               "���������� "&rs("article")&"<br>"&_
               "ע��ʱ�䣺 "&rs("addDAte")&"<br>"&_
               "��½������ "&rs("logins")&"<br>"
        rs.close
        set rs=nothing

response.write"</td></tr></table></td>"&_
              "<td valign=top>"&_
              "<table cellpadding=3 cellspacing=1 style=""width:100%"" height=29 align=center  class=tableborder1>"&_
              "<tr>"&_
              "<th height=25 align=left>-=> ��̳����Ϣ"&_
              "</td></tr><tr><td class=tablebody1>"
if newincept() >0 then
        response.write"Ŀǰ����<font color="""&Forum_body(8)&"""><b> ["&newincept()&"] </b></font>���Ķ���Ϣ��"
        else
        response.write"Ŀǰ��û���µĶ���Ϣ,"
        end if
response.write"<a href=usersms.asp?action=inbox><font color="""&Forum_body(8)&""">�ռ���</font></a>�й��� <b>["&cint(incept())&"]</b> ����Ϣ��<a href='usersms.asp?action=issend'><font color="""&Forum_body(8)&""">������</font></a>�й��� <b>["&allsend()&"]</b> ����Ϣ�Է�δ���ġ�<br>"&_
              "</td></tr></table><br>"&_
              "<table cellpadding=3 cellspacing=1 style=""width:100%"" align=center class=tableborder1>"&_
              "<tr>"&_
              "<th colspan=5 height=25 align=left>-=> �����յ��Ķ�Ѷ</th></tr>"&_
            "<tr>"&_
                "<td align=center valign=middle width=30 class=tabletitle2><b>״̬</b></td>"&_
                "<td align=center valign=middle width=100 class=tabletitle2><b>������</b></td>"&_
                "<td align=center valign=middle width=300 class=tabletitle2><b>����</b></td>"&_
                "<td align=center valign=middle width=150 class=tabletitle2><b>����</b></td>"&_
                "<td align=center valign=middle width=50 class=tabletitle2><b>��С</b></td>"&_
            "</tr>"

	set rs=server.createobject("adodb.recordset")
	sql="select top 5 * from message where incept='"&trim(membername)&"' and issend=1 and delR=0 order by flag,sendtime desc"
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then

response.write"<tr><td class=tablebody1 align=center valign=middle colspan=6>�����ռ�����û���κ����ݡ�</td></tr>"
        else
		dim tablebody
do while not rs.eof
response.write "<tr>"
if rs("flag")=0 then
tablebody="tablebody2"
else
tablebody="tablebody1"
end if
response.write"<td align=center valign=middle class="&tablebody&">"

if rs("flag")=0 then
response.write"<img src=pic/m_news.gif>"
else
response.write"<img src="&Forum_info(7)&"m_olds.gif>"
end if
response.write"</td>"&_
              "<td align=center valign=middle class="&tablebody&">"
if rs("flag")=0 then
response.write"<b>"
end if
response.write"<a href=dispuser.asp?name="&htmlencode(rs("sender"))&" target=_blank>"&htmlencode(rs("sender"))&"</a></td>"&_
              "<td align=left class="&tablebody&"><a href=usersms.asp?action=read&id="&rs("id")&"&sender="&rs("sender")&" >"
if rs("flag")=0 then
response.write"<b>"
end if
response.write""&htmlencode(rs("title"))&"</a></td>"&_
               "<td class="&tablebody&">"
if rs("flag")=0 then
response.write"<b>"
end if
response.write""&rs("sendtime")&"</font></td>"&_
               "<td class="&tablebody&">"
if rs("flag")=0 then
response.write"<b>"
end if
response.write""&len(rs("content"))&"Byte</td></tr>"

	rs.movenext
	loop
	end if
	rs.close
	set rs=nothing

    response.write"</td></tr></table><br>"&_
                  "<table cellpadding=3 cellspacing=1 style=""width:100%"" align=center class=tableborder1>"&_
                  "<tr>"&_
                  "<th height=25 align=left>-=> ������������</th></tr>"

dim topic,srs
set srs=conn.execute("select top 5 rootid,boardid,dateandtime,topic,expression,body,announceid from "&NowUseBBS&" where PostUserID="&userid&" and not locktopic=2 order by announceid desc")
do while not srs.eof
topic=replace(srs(3)," ","")
          if topic<>"" then
          topic=topic
          else
          topic=srs(5)
          topic=replace(topic,chr(13),"")
          topic=replace(topic,chr(10),"")
          end if
          if len(topic)>30 then
          topic=left(topic,30)&"..."
          end if
          topic=server.htmlencode(topic)

 response.write" <tr>"&_ 
               "<td align=left class=tablebody1>&nbsp;����&nbsp;<a href=dispbbs.asp?boardid="&srs(1)&"&id="&srs(0)&"&replyid="&srs(6)&"#"&srs(6)&">"&htmlencode(topic)&"</a>&nbsp;--&nbsp;"&srs(2)&"</td>"&_
               "</tr>"&_

srs.movenext
loop
srs.close
set srs=nothing

response.write"</table><br></td></tr></table>"
end sub

REM ͳ���ѷ����¶���
function allsend()
        rs=conn.execute("Select Count(id) From Message Where flag=0 and issend=1 And sender='"& membername &"'")
        allsend=rs(0)
	set rs=nothing
	if isnull(allsend) then allsend=0
end function
REM ͳ���ռ����еĶ���
function incept()
        incept=0
        rs=conn.execute("Select Count(id) From Message Where  issend=1 and delR=0 And incept='"& membername &"'")
        incept=rs(0)
	set rs=nothing
	if isnull(incept) then incept=0
end function
%>