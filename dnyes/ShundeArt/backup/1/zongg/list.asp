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
      <td width="100%" align="right"><font color="#808080">��������=&gt;&gt;
      <select size="1" name="adorder" style="color: #808080">
<option value="id">���ID</option>
<option value="name">���ƹؼ���</option>
</select> <input type="text" name="nr" size="20">
<input type="submit" value="��ѯ" name="B1" style="color: #808080"></font><hr color="#808080" size="1"></td>
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
                <font size="3" color=red><b>�������ŵ�ͼƬ�������б�</b></font>  <a href=?type=<%=request.querystring("type")%>&px=x>��</a>  <a href=?type=<%=request.querystring("type")%>>��</a>
                <%        Case "txt"
                           adssql="select * from ads where act=1 and xslei='txt' order by regtime "&px
                %>
                <font size="3" color=red><b>�������ŵĴ��ı�������б�</b></font>  <a href=?type=<%=request.querystring("type")%>&px=x>��</a>  <a href=?type=<%=request.querystring("type")%>>��</a>
                <%
                          Case "close"
                           adssql="select * from ads where act=0 order by regtime "&px

                %>
                <font size="3" color=red><b>������ͣ��δʧЧ�Ĺ�����б�</b></font>  <a href=?type=<%=request.querystring("type")%>&px=x>��</a>  <a href=?type=<%=request.querystring("type")%>>��</a>
                <%
                          Case "lose"
                           adssql="select * from ads where act=2 order by regtime "&px
                %>
                <font size="3" color=red><b>�Ѿ�ʧЧ�ĵĹ�����б�</b></font>  <a href=?type=<%=request.querystring("type")%>&px=x>��</a>  <a href=?type=<%=request.querystring("type")%>>��</a> 
                <%
                          Case "click"
                           adssql="select * from ads where act<>2 order by click "&px
                %>
                <font size="3" color=red><b>���������<%if px="desc" then: response.write "����":else:response.write "����":end if%>����δʧЧ�����</b></font>  <a href=?type=<%=request.querystring("type")%>&px=x>��</a>  <a href=?type=<%=request.querystring("type")%>>��</a>
               <%
                          Case "show"
                           adssql="select * from ads where act<>2 order by show "&px
                %>
                <font size="3" color=red><b>����ʾ����<%if px="desc" then: response.write "����":else:response.write "����":end if%>����δʧЧ�����</b></font>  <a href=?type=<%=request.querystring("type")%>&px=x>��</a>  <a href=?type=<%=request.querystring("type")%>>��</a>
               <%
                          Case "place"
                          
                          if isnumeric(request("place"))=true then
                           adssql="select * from ads where act=1 and place="&trim(request("place"))&" order by regtime "&px
                          
                %>
                <font size="3" color=red><b>IDΪ<%=request("place")%>�Ĺ��λ���������ŵĹ����</b></font>  <a href=?type=<%=request.querystring("type")%>&place=<%=request("place")%>&px=x>��</a>  <a href=?type=<%=request.querystring("type")%>&place=<%=request("place")%>>��</a>
                <%else
                  adssql="select * from ads where act=1 order by regtime "&px
                %>
                <font size="3" color=red><b>�����������ŵĹ�����б�</b></font>  <a href=?type=<%=request.querystring("type")%>&px=x>��</a>  <a href=?type=<%=request.querystring("type")%>>��</a>
                        
                <%end if%>
               <%
                          Case "search"
                          if request("adorder")="id" and isnumeric(request("nr"))=true then
                           adssql="select * from ads where id="&trim(request("nr"))
                          
                %>
                <font size="3" color=red><b>��ѯ IDΪ<%=request("nr")%> �Ĺ������Ϣ</b></font>
                <%        else
                  adssql="select * from ads where sitename like '%"&request("nr")&"%' order by regtime "&px
                %>
                <font size="3" color=red><b>��ѯ���ƺ��йؼ��֡�<%=request("nr")%>�������</b></font>  <a href=?type=<%=request.querystring("type")%>&px=x>��</a>  <a href=?type=<%=request.querystring("type")%>>��</a>
                        
                <%end if%>

                <%       
                          Case else
                          adssql="select * from ads where act=1 order by regtime "&px
                %>
                <font size="3" color=red><b>�����������ŵĹ�����б�</b></font>  <a href=?type=<%=request.querystring("type")%>&px=x>��</a>  <a href=?type=<%=request.querystring("type")%>>��</a>
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
response.write "<tr><td bgcolor=#ffffff align=center><BR><BR>û���κ���ؼ�¼<BR><BR><BR><BR>"
else
adsrs.pagesize=advertlistnumber'ÿҳ��ʾ�ļ�¼��
totalPut=adsrs.recordcount '��¼����
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
bookmark=adsrs.bookmark '�ƶ�����ʼ��ʾ�ļ�¼λ��
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
        <td width="20%" bgcolor="#F5F5F5"><font color="#FF0000">&nbsp;���ID��<%=adsrs("id")%> </font></td>
        <td bgcolor="#F5F5F5">&nbsp;���ƣ�<input type="text" name="h" size="23" value="<%=adsrs("sitename")%>"></td>
        <td bgcolor="#F5F5F5" width="300">
       &nbsp;URL�� 
       <input type="text" name="h" size="30" value="<%=adsrs("url")%>"></td>
        <td bgcolor="#F5F5F5" width="100" align="center">
        <%if adsrs("xslei")="txt" then%>
           <a href=javascript:opw('option.asp?id=<%=adsrs("id")%>&job=yulan','banner<%=adsrs("id")%>',400,200)>Ԥ�����</a>
        <%else
        
        %>
            <a href=javascript:opw('option.asp?id=<%=adsrs("id")%>&job=yulan','banner<%=adsrs("id")%>',800,600)>Ԥ�����</a>
       <%end if%>
��</td>
      </tr>
      <tr>
        <td width="20%" height="60">&nbsp;�򿪣�<%= Ggdklx(adsrs("window"))%><br>&nbsp;��ʾ��<%= Ggxslx(adsrs("xslei"))%><br>
        &nbsp;���ͣ�<%= Ggwlx(adsrs("place"))%></td>
        <td height="60">&nbsp;����ʱ�䣺<font color=red><%=adsrs("regtime")%></font><br>&nbsp;������ʾ��<font color=red><%=adsrs("time")%></font><br>
        &nbsp;���µ����<font color=red><%=adsrs("lasttime")%></font></td>
        <td height="60" width="300">&nbsp;��ʾ������<%call  Xscs()%><br>&nbsp;���������<%call  Djcs()%><br>
        &nbsp;�� �� λ��<%= Ggwm(adsrs("place"))%>  ID=<font color=red><%=adsrs("place")%></font></td>
        <td height="60" width="100" align="center">              <%
if adsrs("act")=1 then
%>
              <a href=javascript:opw('option.asp?id=<%=adsrs("id")%>&job=close','close<%=adsrs("id")%>',300,140)>��ͣ</a> 
              <%
else
%>
              <a href=javascript:opw('option.asp?id=<%=adsrs("id")%>&job=open','open<%=adsrs("id")%>',300,140)>����</a> 
              <%end if%>
              ��<a href=javascript:opw('option.asp?id=<%=adsrs("id")%>&job=delit','del<%=adsrs("id")%>',300,140)>ɾ��</a></td>
      </tr>
      <tr>
        <td colspan="3" height="20">&nbsp;ʧЧ������<%call  Sxtj()%></td>
        <td height="20" width="100" align="center"><a href=javascript:opw('edit.asp?id=<%=adsrs("id")%>','<%=adsrs("id")%>',700,520)>�޸���ϸ����</a></td>
      </tr>
      </table>
    </center>
</div>
          <%
i=i+1
if i>=MaxPerPage then exit do 'ѭ��ʱ�����β�������˳��������¼�ﵽҳ�����ʾ����Ҳ�˳�
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
 ��<font color=red><%=totalput%></font>����<font color=red><%=totalPage%></font>ҳ,<font color=red><%=advertlistnumber%></font>��/ҳ,��<font color=red><%=currentPage%></font>ҳ
 
 <a href="?type=<%=request.querystring("type")%>&page=<%=currentPage-1%>">��һҳ</a> <a href="?type=<%=request.querystring("type")%>&page=<%=currentPage+1%>">��һҳ</a>

<%
response.write "<br> ת����<select name='page' size=1>"
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
              ���<font color=red><%=adsrs("clicks")%></font>�� 
              <%
elseif adsrs("class")=2 then
%>
              ��ʾ<font color=red><%=adsrs("shows")%></font>�� 
              <%
elseif adsrs("class")=3 then
%>
              ��ֹ��<font color=red><%=adsrs("lasttime")%></font> 
              <%
elseif adsrs("class")=4 then
%>
              ���<font color=red><%=adsrs("clicks")%></font>�Σ���ʾ<font color=red><%=adsrs("shows")%></font>�� 
              <%
elseif adsrs("class")=5 then
%>
              ���<font color=red><%=adsrs("clicks")%></font>�Σ���ֹ��<font color=red><%=adsrs("lasttime")%></font> 
              <%
elseif adsrs("class")=6 then
%>
              ��ʾ<font color=red><%=adsrs("shows")%></font>�Σ���ֹ��<font color=red><%=adsrs("lasttime")%></font> 
              <%
elseif adsrs("class")=7 then
%>
              ���<font color=red><%=adsrs("clicks")%></font>�Σ���ʾ<font color=red><%=adsrs("shows")%></font>�Σ���ֹ��<font color=red><%=adsrs("lasttime")%></font> 
              <%
else
%>
              ���������� 
<%
end if
%><%end sub

Sub Xscs()%>
 <font color=red><%=adsrs("show")%></font> (<a href=javascript:opw('listip.asp?id=<%=adsrs("id")%>&ip=sip','s<%=adsrs("id")%>',400,500)>��ʾ��¼</a>)
<%end sub

Sub Djcs()%>
 <font color=red><%=adsrs("click")%></font> (<a href=javascript:opw('listip.asp?id=<%=adsrs("id")%>&ip=cip','c<%=adsrs("id")%>',400,500)>�����¼</a>)
<%end sub

%><!--#include file="boot.asp"-->