<%@ Language=VBScript%>
<!--#include file=conn.asp -->
<!--#include file="ConnUser.asp"-->
<!--#include file=config.asp -->
<!--#include file="char.inc"-->
<!--#include file="function.asp"-->
<%
username1=checkstr(trim(Request("user")))

PageShowSize = 10            'ÿҳ��ʾ���ٸ�ҳ
MyPageSize   = 20           'ÿҳ��ʾ������
	
If Not IsNumeric(Request("page")) Or IsEmpty(Request("page")) Or Request("page") <=0 Then
	MyPage=1
	Else
	MyPage=Int(Abs(Request("page")))
End if

if username1="" then
	Response.Write "<script>alert('δָ������');history.back()</script>"
	Response.End
end if

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
set rs=nothing
%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��Ա��<%=username1%>__ȫ����Ա__<%=jjgn%></title>
<LINK href=news.css rel=stylesheet>
</head>
<body topmargin="0">
<!--#include file=top.asp -->
<table width="760" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <tr valign="top"> 
    <td width="180" background="IMAGES/menu-d.gif">
	<table width="180" border="0" align="center" cellpadding="0" cellspacing="0" id="AutoNumber2" style="border-collapse: collapse">
        <tr> 
          <td width="100%" height="25" valign="middle" background="IMAGES/menu-d-user.gif"> 
            <p align="center"><font color="#000000">��Ա����</font></td>
        </tr>
        <%
	set rs2=server.CreateObject("ADODB.RecordSet")
	rs2.Source="select * from "& db_User_Table &" where "& db_User_Name &"='"&username1&"'"
	rs2.Open rs2.Source,ConnUser,1,1
	if not rs2.eof then
	%>
        <tr> 
          <td width="100%" height="24" align=center valign="middle" background="IMAGES/menu-d.gif"> 
            <br> 
	<%if rs2(db_User_Face)<>"" then%>
		<%if UserTableType = "FeiTium" then%>
			<img src="<%=rs2(db_User_Face)%>" border="0" width="120"> 
		<%else%>
			<%if UserTableType="Dvbbs" then%>
				<img src="<%=BbsPath%><%=rs2(db_User_Face)%>" border="0" width="<%=rs2(db_User_FaceWidth)%>" height="<%=rs2(db_User_FaceHeight)%>"> 
				<%''��ʾ�û�ͷ�񣬼�'bbs/'ǰ׺·��,ʹͼ��ϵͳֱ����ʾ������̳ͷ��%><br>
				<a href="<%=BbsPath%>dispuser.asp?name=<%=username1%>" target="_blank"><font color=blue>��̳�еĸ�������</font></a> 
				<br><br>
			<%end if%>
		<%end if%>
	<%else%>
		<img src="images/nopic.gif" border="0"> 
	<%end if%> 
            
          </td>
        </tr>
        <tr> 
          <td width="100%" height="24" valign="middle" background="IMAGES/menu-d.gif"><div align="center">
            <table width="90%" border="0" cellpadding="3" cellspacing="0">
              <tr>
                <td style="LEFT: 0px; WIDTH: 165px; WORD-WRAP: break-word">
                <div>��ʵ������<%=rs2("fullname")%><br>
		<%if db_Birthday_Select = "FeiTium" then%>
		    ��λ��<%=rs2("depname")%> <br>���ţ�<%=rs2("deptype")%><br>
			���գ�<%=rs2(db_User_Birthyear)%>��<%=rs2(db_User_Birthmonth)%>��<%=rs2(db_User_Birthday)%>��<br>
		<%else%>
			<%if db_Birthday_Select = "Text" then%>
				���գ�<%=year(rs2(db_User_birthday))%>��<%=month(rs2(db_User_birthday))%>��<%=day(rs2(db_User_birthday))%>��<br>
			<%end if%>
		<%end if%>
		<%if db_Sex_Select = "FeiTium" then%>
			�Ա�<%=rs2(db_User_sex)%><br>
		<%else%>
			<%if db_Sex_Select = "Number" then%>
				�Ա�<%if rs2(db_User_Sex)="0" then%>Ů<%else%>��<%end if%><br>
			<%end if%>
		<%end if%>
		EMAIL��<%=rs2(db_User_Email)%><br>
		ע��ʱ�䣺<%=year(rs2(db_User_RegDate))%>��<%=month(rs2(db_User_RegDate))%>��<%=day(rs2(db_User_RegDate))%>��<br>
		���ҽ��ܣ� <%=rs2("content")%>
		</div></td>
              </tr>
            </table>
          </div></td>
        </tr>
        <%rs2.close
	set rs2=nothing%>
        
        
      </table>
    </td>
    <td width="6">&nbsp;</td>
    <td width="574" background="IMAGES/menu-l.gif"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td height="25" background="IMAGES/menu-l-m.gif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font class=m_tittle>&nbsp;</font>��Ŀ����<b>&nbsp;&nbsp; 
            ��ǰλ�ã�<a class="daohang" href="./" >��վ��ҳ</a>&gt;&gt;<a class=daohang href="alluser.asp">ȫ����Ա</a>>>��Ա��<font color=red><%=username1%></font></b></td>
        </tr>
        <tr> 
          <td height="25" background="IMAGES/menu-l.gif"><div align="center"><b> 
              <font color=red> <a class=daohang href="./" > 
              <script language=javascript src=zongg/ad.asp?i=14></script> 
              </a> </font></div></td>
        </tr>
        <tr> 
          <td height="25" background="IMAGES/menu-l-zj.gif" class="daohang">&nbsp;</td>
        </tr>
        <tr> 
          <td height="25" background="IMAGES/menu-l.gif"><div align="center"> 
              <table width="95%"  border="0" align="center" cellpadding="0" cellspacing="0" id="AutoNumber3" style="border-collapse: collapse">
                <tr> 
                  <td valign=top>&nbsp;</td>
                </tr>
                <tr> 
                  <td valign=top><table border="0" cellpadding="3" cellspacing="0" width="100%">
                      <%
			set rs=server.CreateObject("ADODB.RecordSet")
			if session("key")="" then
				rs.Source="select * from "& db_News_Table &" where editor='" & username1 & "' and checkked=1 and newslevel=0 order by NewsID DESC"
			end if
			if session("key")="super" or session("key")="typemaster" or session("key")="bigmaster" or session("key")="smallmaster" or session("key")="check" then
				rs.Source="select * from "& db_News_Table &" where editor='" & username1 & "' and checkked=1 order by NewsID DESC"
			end if
			if session("key")="selfreg" then
				if session("reglevel")=3 then
					rs.Source="select * from "& db_News_Table &" where editor='" & username1 & "' and checkked=1 order by NewsID DESC"
				end if
				if session("reglevel")=2 then
					rs.Source="select * from "& db_News_Table &" where editor='" & username1 & "' and checkked=1 order by NewsID DESC"
				end if
				if session("reglevel")=1 then
					rs.Source="select * from "& db_News_Table &" where editor='" & username1 & "' and checkked=1 order by NewsID DESC"
				end if
			end if

			rs.Open rs.Source,conn,1,1
			rs.PageSize=20
			rs.CacheSize = RS.PageSize
			for i=1 to rs.PageSize *(page-1)
				if not rs.EOF then
					rs.MoveNext
				end if
			next

			Response.Write "<tr><td width=100% height=28 align=center>&nbsp;"
			if rs.EOF then
				Response.Write "<font color=red>�û�Ա��û�����£�</font>"
			else
				rs.PageSize     = MyPageSize
				MaxPages         = rs.PageCount
				rs.absolutepage = MyPage
				total            = rs.RecordCount
				Response.Write "�û�Ա����" & total & "ƪ���£���ǰ��"& myPage &"/"& MaxPages &"ҳ��ÿҳ"& rs.PageSize &"��"
			end if
			Response.Write "</td></tr>"

			If Not rs.eof then
				i = 0
				do until rs.Eof or i = rs.PageSize
				if rs("picname")<>"" then
					img="<font color=red>[ͼ]</font>"
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
                        <td width="100%" height="20">&nbsp;<img src="images/006.gif" width="8" height="10" border="0">&nbsp;<a class=middle href="ReadNews.asp?NewsID=<%=rs("NewsID")%>" target=_blank title="<%=title%>"> 
                          <%if mode="2" then%>
                          <%=img%> 
                          <%end if%>
                          <font color="<%=rs("titlecolor")%>">                
                          <%=CutStr(title,50)%> 
                          </font></a><font class=middle>(<%=rs("UpdateTime")%>)[<font color="#ff0000"><%=rs("click")%></font>]</font></td>
                      </tr>
			<%
			rs.MoveNext
			i = i + 1
			loop
			%>
                      <tr> 
                        <td width="100%" align=center>�� <%=Mypage%>/<%=Maxpages%>ҳ��ÿҳ <%=MyPageSize%> �� 
			<%
			url="user.asp?user=" & username1
			PageNextSize=int((MyPage-1)/PageShowSize)+1
			Pagetpage=int((total-1)/rs.PageSize)+1
			if PageNextSize >1 then
				PagePrev=PageShowSize*(PageNextSize-1)
				Response.write "<a class=black href='" & Url & "&page=" & PagePrev & "' title='��" & PageShowSize & "ҳ'>��һ��ҳ</a> "
				Response.write "<a class=black href='" & Url & "&page=1' title='��1ҳ'>ҳ��</a> "
			end if
			if MyPage-1 > 0 then
				Prev_Page = MyPage - 1
				Response.write "<a class=black href='" & Url & "&page=" & Prev_Page & "' title='��" & Prev_Page & "ҳ'>��һҳ</a> "
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
					Response.write "<a class=black href='" & Url & "&page=" & Next_Page & "' title='��" & Next_Page & "ҳ'>��һҳ</A>"
				end if

				if MaxPages > PageShowSize*PageNextSize then
					PageNext = PageShowSize * PageNextSize + 1
					Response.write " <A class=black href='" & Url & "&page=" & Pagetpage & "' title='��"& Pagetpage &"ҳ'>ҳβ</A>"
					Response.write " <a class=black href='" & Url & "&page=" & PageNext & "' title='��" & PageShowSize & "ҳ'>��һ��ҳ</a>"
				End if
			end if
			rs.close					
			%>
                        </td>
                      </tr>
                    </table></td>
                </tr>
                <tr> 
                  <td valign=top>&nbsp;</td>
                </tr>
              </table>
            </div></td>
        </tr>
      </table>
    </td>
  </tr>
  <tr valign="top">
    <td height="20" background="IMAGES/menu-d-t.gif">&nbsp;</td>
    <td>&nbsp;</td>
    <td height="20" background="IMAGES/menu-l-t.gif" class="daohang">&nbsp;</td>
  </tr>
</table>
<!--#include file=bottom.asp -->
</body>
</html>
<%
else

Response.Write "<script>alert('�޴��û���');history.back()</script>"
Response.End

end if
%>