<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<%
	'on error resume next
	dim currentpage,page_count,Pcount
	dim totalrec,endpage
	dim stype,pSearch,nSearch,keyword
	dim searchboard,ordername
	dim searchdate,searchDateLimit,searchday
	dim stable
	stats="�������"
	stype=request("stype")
	pSearch=request("pSearch")
	nSearch=request("nSearch")
	keyword=trim(checkStr(request("keyword")))
	stable=checkstr(request("stable"))
	if stable="" then stable=Nowusebbs
	if request("page")="" or not isInteger(request("page"))  then
		currentPage=1
	else
		currentPage=cint(request("page"))
	end if
	if stype<3 then
	if keyword="" then
		Errmsg=Errmsg+"<br>"+"<li>���������ѯ�ؼ��֡�"
		founderr=true
	end if
	'����������������
	searchDateLimit=180
	if request("SearchDate")="ALL" then
	searchday=" "
	else
		if not isInteger(request("SearchDate"))  then
			Errmsg=Errmsg+"<br>"+"<li>���������������������"
			founderr=true
		else
			searchday=" datediff('d',dateandtime,Now()) < "&request("SearchDate")&" and "
		end if
	end if
	end if


	call nav()
	if cint(boardid)<1 then
	call head_var(1)
	searchboard=""
	else
	call head_var(2)
	searchboard=" boardid="&boardid&" and "
	end if

	if Cint(GroupSetting(14))=0 then
		Errmsg=Errmsg+"<br>"+"<li>��û���ڱ���̳������Ȩ�ޣ���<a href=login.asp>��½</a>����ͬ����Ա��ϵ��"
		founderr=true
	end if
	if founderr then
		call dvbbs_error()
	else
		call search()
		if founderr then call dvbbs_error()
	end if
	call footer()

	sub search()
	sql=""
	set rs=server.createobject("adodb.recordset")
	select case stype
	case 1
		select case nSearch
		'����������������
		case 1
		sql="select top 1000 locktopic,boardid,rootid,announceid,body,Expression,topic,username,dateandtime,postuserid from "&stable&" where username='"&keyword&"' and "&searchboard&" "&searchday&" parentid=0 and not locktopic=2 ORDER BY announceID desc"
		ordername="����������������"
		'�����ظ���������
		case 2
		sql="select top 1000 locktopic,boardid,rootid,announceid,body,Expression,topic,username,dateandtime,postuserid from "&stable&" where username='"&keyword&"' and "&searchboard&" "&searchday&" parentid>0 and not locktopic=2 ORDER BY announceID desc"
		ordername="�����ظ���������"
		'���߶�����
		case 3
		sql="select top 1000  locktopic,boardid,rootid,announceid,body,Expression,topic,username,dateandtime,postuserid from "&stable&" where  "&searchboard&" "&searchday&" username='"&keyword&"' and not locktopic=2 ORDER BY announceID desc"
		ordername="��������ͻظ���������"
		end select
	case 2
		select case pSearch
		'��������ؼ���
		case 1
		sql="select top 1000 locktopic,boardid,rootid,announceid,body,Expression,topic,username,dateandtime,postuserid from "&stable&" where  "&searchboard&" "&searchday&" (" & translate(keyword,"topic") & ") and not locktopic=2 ORDER BY announceID desc"
		'�������ݹؼ���
		ordername="��������"
		case 2
		sql="select top 1000  locktopic,boardid,rootid,announceid,body,Expression,topic,username,dateandtime,postuserid from "&stable&" where   "&searchboard&" "&searchday&" (" & translate(keyword,"body") & ") and not locktopic=2 ORDER BY announceID desc"
		'���߶�����
		ordername="��������"
		case 3
		sql="select top 1000  locktopic,boardid,rootid,announceid,body,Expression,topic,username,dateandtime,postuserid from "&stable&" where   "&searchboard&" "&searchday&" (" & translate(keyword,"topic") & " or " & translate(keyword,"body") & ") and not locktopic=2 ORDER BY announceID desc"
		ordername="�������������"
		end select
	case 3
		sql="select top 50  locktopic,boardid,rootid,announceid,body,Expression,topic,username,dateandtime,postuserid from "&stable&" where not locktopic=2 ORDER BY announceID desc"
		ordername="����50��"
	end select
	
	if sql="" then
		Errmsg=Errmsg+"<br>"+"<li>��ָ����ѯ������"
		founderr=true
		exit sub
	end if
	rs.open sql,conn,1,1

	if rs.eof and rs.bof then
		Errmsg=Errmsg+"<br>"+"<li>û���ҵ���Ҫ��ѯ�����ݡ�"
		founderr=true
		exit sub
	else
		rs.PageSize = Forum_Setting(11)
		rs.AbsolutePage=currentpage
		page_count=0
		totalrec=rs.recordcount
       		call searchinfo()
		call listPages3()
	end if

	rs.close
	set rs=nothing
	call activeonline()
	end sub

	sub searchinfo()
%>
            <table cellpadding=0 cellspacing=0 border=0 width="<%=Forum_body(12)%>" align=center>
            <tr><td>��ѯ<%if request("SearchDate")="ALL" then%>��������<%else%><%=request("SearchDate")%>����<%end if%>�����ӣ�<%=ordername%>
<%if totalrec>=1000 then%>
���ֻ�ܲ�ѯ��<font color=<%=Forum_body(8)%>>1000</font>����¼������ľͲ���ʾ��
<%else%>����ѯ��
<font color=<%=Forum_body(8)%>><%=totalrec%></font>�����
<%end if%>
		</td>
            </tr>
            </table>
<TABLE cellPadding=3 cellSpacing=1 class=tableborder1 align=center>
<TR valign=middle>
<Th height=25 width=32>״̬</Th>
<Th width=*>�� ��</Th>
<Th width=80>�� ��</Th>
<Th width=195>������ | �ظ���</Th>
</TR>
<%
       while (not rs.eof) and (not page_count = rs.PageSize)
%>
  <TR> 
    <TD align=middle class=tablebody2 width=32>
<%if rs("locktopic")=1 then%><img src=<%=Forum_info(7)%>lockfolder.gif alt="������������"><%else%><img src=<%=Forum_info(7)%>folder.gif><%end if%>
    </TD> 
    <TD  class=tablebody1 width=*><a href='dispbbs.asp?boardID=<%=rs("boardID")%>&replyID=<%=rs("announceID")%>&ID=<%=rs("RootID")%>&skin=1' target=_blank><img src='<%=Forum_info(8)%><%if instr(rs("Expression"),Forum_Setting(44))>0 then%><%=rs("Expression")%><%else%>face1.gif<%end if%>' border=0 alt="���´������������"></a> <a href='dispbbs.asp?boardID=<%=rs("boardID")%>&replyID=<%=rs("announceID")%>&ID=<%=rs("RootID")%>&skin=1'><%if rs("topic")="" then%><%=left(htmlencode(replace(rs("body"),chr(10)," ")),26)%>...<%else%><%=htmlencode(rs("topic"))%><%end if%></a>    </TD> 
    <TD align=middle  class=tablebody2  width=80><a href="dispuser.asp?id=<%=htmlencode(rs("postuserid"))%>"><%=htmlencode(rs("username"))%></a></TD> 
    <TD  class=tablebody1 width=195>&nbsp;
<%=FormatDateTime(rs("dateandtime"),2)%>&nbsp;<%=FormatDateTime(rs("dateandtime"),4)%>
&nbsp;<font color="<%=Forum_body(8)%>">|</font>&nbsp;
<a href="dispuser.asp?id=<%=rs("postuserid")%>" target=_blank><%=htmlencode(rs("username"))%></a>
</TD>
</TR> 
<%
	  page_count = page_count + 1
          rs.movenext
        wend
%>
            </table>
<%
	end sub

	sub listPages3()

	Pcount=rs.PageCount
	response.write "<table border=0 cellpadding=0 cellspacing=3 width="""&Forum_body(12)&""" align=center>"&_
			"<tr><td valign=middle nowrap>"&_
			"<font color="&Forum_body(13)&">ҳ�Σ�<b>"&currentpage&"</b>/<b>"&Pcount&"</b>ҳ"&_
			"ÿҳ<b>"&Forum_Setting(11)&"</b> ������<b>"&totalrec&"</b></font></td>"&_
			"<td valign=middle nowrap><font color="&Forum_body(13)&"><div align=right><p>��ҳ�� "

	if currentpage > 4 then
	response.write "<a href=""?page=1&stype="&stype&"&pSearch="&pSearch&"&nSearch="&nSearch&"&keyword="&keyword&"&SearchDate="&request("SearchDate")&"&boardid="&boardid&"&stable="&stable&""">[1]</a> ..."
	end if
	if Pcount>currentpage+3 then
	endpage=currentpage+3
	else
	endpage=Pcount
	end if
	for i=currentpage-3 to endpage
	if not i<1 then
		if i = clng(currentpage) then
        response.write " <font color="&Forum_body(8)&">["&i&"]</font>"
		else
        response.write " <a href=""?page="&i&"&stype="&stype&"&pSearch="&pSearch&"&nSearch="&nSearch&"&keyword="&keyword&"&SearchDate="&request("SearchDate")&"&boardid="&boardid&"&stable="&stable&""">["&i&"]</a>"
		end if
	end if
	next
	if currentpage+3 < Pcount then 
	response.write "... <a href=""?page="&Pcount&"&stype="&stype&"&pSearch="&pSearch&"&nSearch="&nSearch&"&keyword="&keyword&"&SearchDate="&request("SearchDate")&"&boardid="&boardid&"&stable="&stable&""">["&Pcount&"]</a>"
	end if
	response.write "</p></div></font></td></tr></table>"

	end sub

public function translate(sourceStr,fieldStr)
rem �����߼����ʽ��ת������
  dim  sourceList
  dim resultStr
  dim i,j
  if instr(sourceStr," ")>0 then 
     dim isOperator
     isOperator = true
     sourceList=split(sourceStr)
     '--------------------------------------------------------
     rem Response.Write "num:" & cstr(ubound(sourceList)) & "<br>"
     for i = 0 to ubound(sourceList)
        rem Response.Write i 
    Select Case ucase(sourceList(i))
    Case "AND","&","��","��"
        resultStr=resultStr & " and "
        isOperator = true
    Case "OR","|","��"
        resultStr=resultStr & " or "
        isOperator = true
    Case "NOT","!","��","��","��"
        resultStr=resultStr & " not "
        isOperator = true
    Case "(","��","��"
        resultStr=resultStr & " ( "
        isOperator = true
    Case ")","��","��"
        resultStr=resultStr & " ) "
        isOperator = true
    Case Else
        if sourceList(i)<>"" then
            if not isOperator then resultStr=resultStr & " and "
            if inStr(sourceList(i),"%") > 0 then
                resultStr=resultStr&" "&fieldStr& " like '" & replace(sourceList(i),"'","''") & "' "
            else
                resultStr=resultStr&" "&fieldStr& " like '%" & replace(sourceList(i),"'","''") & "%' "
            end if
                isOperator=false
        End if    
    End Select
        rem Response.write resultStr+"<br>"
     next 
     translate=resultStr
  else '������
     if inStr(sourcestr,"%") > 0 then
         translate=" " & fieldStr & " like '" & replace(sourceStr,"'","''") &"' "
     else
    translate=" " & fieldStr & " like '%" & replace(sourceStr,"'","''") &"%' "
     End if
     rem ǰ�����һ���ո������sqlʱ���˼ӣ�������
  end if  
end function
%>