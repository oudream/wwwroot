<!--#include file="conn.asp"-->
<!--#include file="function.asp"-->
<!--#include file="char.inc"-->
<%
dim shu,order,path,ss,oo,tt
shu=request.querystring("shu")
order=request("order")
title=request.querystring("title")'
Path = "./"
if shu<>"" then   '��ʾ��Ŀ������Ϊ6
	ss=shu
else
	ss=6
end if
if order="reviewid" then  '����ʽ��clickΪ�������
	oo="UpdateTime"
else
	oo="reviewid"
end if

if title<>"" then   '��ʾ������Ŀ������Ϊ10
	nn=title
else
	nn=10
end if
set rs=conn.execute("select top "&ss&" * from "& db_Review_Table &" where passed=1 order by "&oo&" desc")
if rs.eof and rs.bof then %>
	<div align=center><br>û ��</div>
<%else
	do while not rs.eof
		title=nohtml(CheckStr(trim(rs("Content"))))'��ȡ����������ֶ����ݲ��滻��ʽ��
		if rs("NewsID")>0 and rs("passed")=1 then%>
			<LINK href=news.css rel=stylesheet><br><font class=class>����<A class=class href="<%=path%>ReadNews.asp?NewsID=<%=rs("NewsID")%>" title="<%=CutStr(title,30)%>" target="_blank"><%=CutStr(title,nn)%></A><%if tt=1 then%><%end if%><%if date()-rs("updatetime")<2 then%><FONT color="#ff0000">&nbsp;��!<%end if%></font>
		<%else%>
			<%if rs("NewsID")<0 and rs("passed")=1 then%>
				<LINK href=news.css rel=stylesheet><br><font class=class>�ԡ�<A class=class href="<%=path%>GuestBook.asp?ReViewID=<%=rs("ReViewID")%>" title="������Ա�����Ļ����ǹ���Ա���ܲ鿴!" target="_blank"><font color=green>&nbsp;���Ļ�!</font></A><%if date()-rs("updatetime")<2 then%><FONT color="#ff0000">&nbsp;��!</FONT><%end if%></font>
			<%else%>
				<LINK href=news.css rel=stylesheet><br><font class=class>�ԡ�<A class=class href="<%=path%>GuestBook.asp?ReViewID=<%=rs("ReViewID")%>" title="<%=CutStr(title,30)%>" target="_blank"><%=CutStr(title,nn)%></A><%if date()-rs("updatetime")<2 then%><FONT color="#ff0000">&nbsp;��!</FONT><%end if%></font>
			<%end if%>
		<%end if
		rs.movenext 
	loop 
end if
Rs.Close
set Rs=nothing
%>