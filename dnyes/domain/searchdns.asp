
<% 
 ' ---------------------------------------------------------------------------
 '  �˳����ɷ��� http://www.fxsw.net ����ṩ�������������xmlhttp��ѯ������������������
 '������ַ��ͨ����ַ���������ȵ������Ĳ�ѯ����.
 '
 '��������������ʾ http://test.fxsw.net ��̨ http://test.fxsw.net/admin �û��� /���� admin
 '
 '----------------------------------------------------------------------------

%> 

<html>
<head>

<%

on error resume next
domain1=request("domain")
ext=request("ext")
Information=false

if ext="" or ext=empty then 
domain=domain1
else
domain=domain1&"."&ext
end if

function bytes2bstr(vin)
strreturn = ""
for i = 1 to lenb(vin)
thischarcode = ascb(midb(vin,i,1))
if thischarcode < &h80 then
strreturn = strreturn & chr(thischarcode)
else
nextcharcode = ascb(midb(vin,i+1,1))
strreturn = strreturn & chr(clng(thischarcode) * &h100 + cint(nextcharcode))
i = i + 1
end if
next
bytes2bstr = strreturn
end function

Function GetURL(url,ext)
    Set Retrieval = Server.CreateObject("Microsoft.XMLHTTP")
          With Retrieval
          .Open "GET", url, False, "", ""
          .Send
		  if instr(ext,"cn")<>0 then
          GetURL = .responseText
		  else
		  GetURL = .responsebody
		  GetURL= bytes2bstr(GetURL)
		  GetURL=trim(GetURL)
		end if
          End With
    Set Retrieval = Nothing
End Function

 if ext<>"" then
   if ext="com" or ext="net" or ext="org" or ext="biz" or ext="info" then
   			TakenHTML = GetURL("http://www-whois.internic.net/cgi/whois?type=domain&whois_nic=" & Domain,ext)
            pageafter = "<pre>"
			pagebefore = "</pre>"
   end if
   
   if instr(ext,"cn")<>0 then
        TakenHTML = GetURL("http://ewhois.cnnic.net.cn/whois?entity=domain&value="&Domain,ext)
		pageafter = "<table border=1 cellspacing=0 cellpadding=2>"
		pagebefore = "</table>"
     end if
	 
  if ext="�й�" or ext="��˾" or ext="����" then
        TakenHTML = GetURL("http://cwhois.cnnic.net.cn/whois.jsp?entityname=domain&queryinfo=" &Domain,ext)
        pageafter = "<table border=1 cellspacing=0 cellpadding=2>"
        pagebefore = "</table>"
    end if
	
 else
 
   TakenHTML = GetURL("http://seal.cnnic.net.cn/queryserver?nameInfo=" & Domain & "&entityName=keyword","ͨ����ַ")
   pageafter = "<table border=1 cellspacing=0 cellpadding=2>"
   pagebefore = "</table>"
   
end if

  
   tempcontent=TakenHTML
   if TakenHTML="" then
   TakenHTML="��ѯ�д�"
   else
    if instr(TakenHTML,"û�б�ע��")<>0 then
	  TakenHTML="������û�б�ע��"
	  else
       if instr(1,TakenHTML,pageafter,1)<>0 then
          pagestart=InStr(1,tempcontent, pageafter,1)+ Len(pageafter) + 1 
	      pageend=instr(pagestart,tempcontent,pagebefore,1)
	      pageend = pageend - pagestart 
	      TakenHTML = Mid(tempcontent, pagestart, pageend)
		  TakenHTML=pageafter&TakenHTML&pagebefore
      end if
	  
    end if
  end if



%>

<title><%=page_title%>������ѯ</title>
<link href="css.css" rel="stylesheet" type="text/css">
<link href="css.css" rel="stylesheet" type="text/css">
<link href="css.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312"></head>

<body bgcolor="#FFFFFF" leftmargin="2" topmargin="2">
<p align="center">&nbsp;</p>
<p align="center"><font color="#000000" size="4"><strong>����whios��ѯ����</strong></font></p>
<hr width="90%" size="1">
<p align="center">&nbsp;</p>
<table width="768" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="768"> <p> 
        <%if InStr(TakenHTML,"No entries") > 1 then%>
      </p>
      <table width="530" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr> 
          <td><span class="bk"><font color="#FF0000">�Բ�������ѯ������<font color="#000000"> 
            <%=Domain%></font> ���ִ���!</font></span><font color="#FF0000"><br>
            <br>
            ������ʵ�����²�ѯ��лл��</font></td>
        </tr>
      </table>
      <%end if%>
      <%if InStr(TakenHTML,"No match") > 1 or InStr(TakenHTML,"Register system prompt")>1 or InStr(TakenHTML,"not been registered yet")>1 or InStr(TakenHTML,"û���ҵ������Ŀ")>1 or InStr(TakenHTML,"û�в鵽���������ļ�¼")>1 or InStr(TakenHTML,"��Ϣ������")>1 or InStr(TakenHTML,"Not Exist")>1 or InStr(TakenHTML,"NOT FOUND")>1 or  InStr(TakenHTML,"������ѯ����Ϣ������")>1 or  InStr(TakenHTML,"������û�б�ע��")=1 then %>
      <table width="530" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr> 
          <td height="54" valign="top"><font color="#FF0000">��ϲ��������ѯ������<font color="#000000"> 
            <% = Domain%>
            </font> ��û�б�ע��</font><br> <br>
            ������<a href="http://www.fxsw.net/dns_shop.asp?domain=<%=request("domain")%>&ext=<%=request("ext")%>"> 
            <strong><font color="#3366FF">�ύע���� &gt;&gt;</font></strong></a></td>
        </tr>
      </table>
      <%else%>
      <table width="472" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr> 
          <td valign="top"> <p><strong><font color="#006600">�Բ���������ѯ������ <font color="#FF0000"> 
              <% = Domain%>
              </font> �ѱ�ע�ᣬ�������²�ѯ</font></strong></p>
            <p> 
              <%end if%>
            </p>
            <table border="0" cellspacing="1" cellpadding="0" align="center" width="390" bgcolor="#285278">
              <tr bgcolor="#FFFFFF"> 
                <td> <table width="388" border="0" cellspacing="0" cellpadding="0">
                    <form action="searchdns.asp" method="post" name="checkdomain">
                      <tr bgcolor="#B5C9D7"> 
                        <td height="25" valign="middle" bgcolor="B2D1F8"> <div align="center">Ӣ��������ѯ��<font color="#FF6600"><b>WWW</b></font>. 
                            <input type="text" name="domain" id=Text1 size="17" class="form">
                            <input type="hidden" value="1" name="show">
                            <input type="image" border="0" name="imageField" src="img/sousuo.gif" width="40" height="18" align="middle" >
                          </div></td>
                      </tr>
                      <tr> 
                        <td height="25"> <input name="ext" type="radio" class="unnamed1" value="com" checked>
                          .com 
                          <input name="ext" type="radio" class="unnamed1" value="net">
                          .net 
                          <input name="ext" type="radio" class="unnamed1" value="org">
                          .org 
                          <input name="ext" type="radio" class="unnamed1" value="cc">
                          .cc 
                          <input name="ext" type="radio" class="unnamed1" value="tv">
                          .tv 
                          <input name="ext" type="radio" class="unnamed1" value="info">
                          .info 
                          <input name="ext" type="radio" class="unnamed1" value="biz">
                          .biz </td>
                      </tr>
                      <tr> 
                        <td height="25"> <input name="ext" type="radio" class="unnamed1" value="com.cn">
                          .com.cn 
                          <input name="ext" type="radio" class="unnamed1" value="net.cn">
                          .net.cn 
                          <input name="ext" type="radio" class="unnamed1" value="org.cn">
                          .org.cn 
                          <input name="ext" type="radio" class="unnamed1" value="cn"> 
                          <strong><font color="#FF0000">.cn </font></strong></td>
                      </tr>
                    </form>
                  </table></td>
              </tr>
            </table>
            <br> <table border="0" cellspacing="1" cellpadding="0" align="center" width="390" bgcolor="#285278">
              <tr bgcolor="#FFFFFF"> 
                <td> <table width="388" border="0" cellspacing="0" cellpadding="0">
                    <form action="searchdns.asp" method="post" name="checkdomain">
                      <tr bgcolor="#B5C9D7"> 
                        <td height="25" valign="middle" bgcolor="B2D1F8"> <div align="center">����������ѯ��<font color="#FF6600"><b>WWW</b></font>. 
                            <input type="text" name="domain" id=Text1 size="17" class="form">
                            . 
                            <select name=ext>
                              <option selected value=�й�>�й�</option>
                              <option value=����>����</option>
                              <option value=��˾>��˾</option>
                            </select>
                            <input type="hidden" value="2" name="show">
                            <input type="image" border="0" name="imageField" src="img/sousuo.gif" width="40" height="18" align="middle">
                          </div></td>
                      </tr>
                    </form>
                  </table></td>
              </tr>
            </table>
            <br> <table border="0" cellspacing="1" cellpadding="0" align="center" width="390" bgcolor="#285278">
              <tr bgcolor="#FFFFFF"> 
                <td> <table width="388" border="0" cellspacing="0" cellpadding="0">
                    <form action="searchdns.asp" method="post" name="checkdomain">
                      <tr bgcolor="#B5C9D7"> 
                        <td height="25" valign="bottom" bgcolor="B2D1F8"> <div align="center">ͨ����ַ��ѯ�� 
                            <input type="text" name="domain" id=Text1 size="27" class="form">
                            <input type="hidden" value="3" name="show">
                            <input type="image" border="0" name="imageField" src="img/sousuo.gif" width="40" height="18" align="middle">
                          </div></td>
                      </tr>
                    </form>
                  </table></td>
              </tr>
            </table>
            <p><br>
              <br>
              <%if InStr(TakenHTML,"No match") > 1 or InStr(TakenHTML,"Register system prompt")>1 or InStr(TakenHTML,"not been registered yet")>1 or InStr(TakenHTML,"û���ҵ������Ŀ")>1 or InStr(TakenHTML,"������û�б�ע��")>1 or InStr(TakenHTML,"û�в鵽���������ļ�¼")>1 or InStr(TakenHTML,"��Ϣ������")>1 or InStr(TakenHTML,"Not Exist")>1 or InStr(TakenHTML,"NOT FOUND")>1 or InStr(TakenHTML,"������ѯ����Ϣ������")>1 then%>
              ���������Ծ�ע���ˡ��� 
              <% 
			  else %>
              <br>
              ������ע����Ϣ<br>
              <%=TakenHTML%></p>
            <p> 
              <% end if %>
            </p></td>
        </tr>
      </table></td>
  </tr>
</table>
<hr width="90%" size="1">
<p align="center">�������ɷ��������ṩ <a href="http://www.fxsw.net%20">http://www.fxsw.net 
  </a>���κ��������������ϵ tel: 3574338822 QQ: 11752290</p>
<p align="center">200M asp�Ŀռ�+100M��ҵ�ʾ�+��������(com/.net/.org)/֧�ֶ�����̳ 350Ԫһ��</p>
<p align="center">�˳�������ṩ,����Ҫʹ�õ��뱣�����ǵİ�Ȩ</p>
<p>&nbsp;</p>
</body>
</html>
