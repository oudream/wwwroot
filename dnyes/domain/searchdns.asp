
<% 
 ' ---------------------------------------------------------------------------
 '  此程序由飞翔 http://www.fxsw.net 免费提供，这是无组件用xmlhttp查询国际域名，国内域名
 '中文网址，通用网址。。。。等等域名的查询程序.
 '
 '飞翔主机程序演示 http://test.fxsw.net 后台 http://test.fxsw.net/admin 用户名 /密码 admin
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
	 
  if ext="中国" or ext="公司" or ext="网络" then
        TakenHTML = GetURL("http://cwhois.cnnic.net.cn/whois.jsp?entityname=domain&queryinfo=" &Domain,ext)
        pageafter = "<table border=1 cellspacing=0 cellpadding=2>"
        pagebefore = "</table>"
    end if
	
 else
 
   TakenHTML = GetURL("http://seal.cnnic.net.cn/queryserver?nameInfo=" & Domain & "&entityName=keyword","通用网址")
   pageafter = "<table border=1 cellspacing=0 cellpadding=2>"
   pagebefore = "</table>"
   
end if

  
   tempcontent=TakenHTML
   if TakenHTML="" then
   TakenHTML="查询有错"
   else
    if instr(TakenHTML,"没有被注册")<>0 then
	  TakenHTML="此域名没有被注册"
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

<title><%=page_title%>域名查询</title>
<link href="css.css" rel="stylesheet" type="text/css">
<link href="css.css" rel="stylesheet" type="text/css">
<link href="css.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312"></head>

<body bgcolor="#FFFFFF" leftmargin="2" topmargin="2">
<p align="center">&nbsp;</p>
<p align="center"><font color="#000000" size="4"><strong>域名whios查询程序</strong></font></p>
<hr width="90%" size="1">
<p align="center">&nbsp;</p>
<table width="768" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="768"> <p> 
        <%if InStr(TakenHTML,"No entries") > 1 then%>
      </p>
      <table width="530" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr> 
          <td><span class="bk"><font color="#FF0000">对不起，您查询的域名<font color="#000000"> 
            <%=Domain%></font> 出现错误!</font></span><font color="#FF0000"><br>
            <br>
            请您核实后重新查询，谢谢！</font></td>
        </tr>
      </table>
      <%end if%>
      <%if InStr(TakenHTML,"No match") > 1 or InStr(TakenHTML,"Register system prompt")>1 or InStr(TakenHTML,"not been registered yet")>1 or InStr(TakenHTML,"没有找到这个条目")>1 or InStr(TakenHTML,"没有查到符合条件的记录")>1 or InStr(TakenHTML,"信息不存在")>1 or InStr(TakenHTML,"Not Exist")>1 or InStr(TakenHTML,"NOT FOUND")>1 or  InStr(TakenHTML,"你所查询的信息不存在")>1 or  InStr(TakenHTML,"此域名没有被注册")=1 then %>
      <table width="530" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr> 
          <td height="54" valign="top"><font color="#FF0000">恭喜，您所查询的域名<font color="#000000"> 
            <% = Domain%>
            </font> 还没有被注册</font><br> <br>
            　　　<a href="http://www.fxsw.net/dns_shop.asp?domain=<%=request("domain")%>&ext=<%=request("ext")%>"> 
            <strong><font color="#3366FF">提交注册申 &gt;&gt;</font></strong></a></td>
        </tr>
      </table>
      <%else%>
      <table width="472" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr> 
          <td valign="top"> <p><strong><font color="#006600">对不起，您所查询的域名 <font color="#FF0000"> 
              <% = Domain%>
              </font> 已被注册，请您重新查询</font></strong></p>
            <p> 
              <%end if%>
            </p>
            <table border="0" cellspacing="1" cellpadding="0" align="center" width="390" bgcolor="#285278">
              <tr bgcolor="#FFFFFF"> 
                <td> <table width="388" border="0" cellspacing="0" cellpadding="0">
                    <form action="searchdns.asp" method="post" name="checkdomain">
                      <tr bgcolor="#B5C9D7"> 
                        <td height="25" valign="middle" bgcolor="B2D1F8"> <div align="center">英文域名查询：<font color="#FF6600"><b>WWW</b></font>. 
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
                        <td height="25" valign="middle" bgcolor="B2D1F8"> <div align="center">中文域名查询：<font color="#FF6600"><b>WWW</b></font>. 
                            <input type="text" name="domain" id=Text1 size="17" class="form">
                            . 
                            <select name=ext>
                              <option selected value=中国>中国</option>
                              <option value=网络>网络</option>
                              <option value=公司>公司</option>
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
                        <td height="25" valign="bottom" bgcolor="B2D1F8"> <div align="center">通用网址查询： 
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
              <%if InStr(TakenHTML,"No match") > 1 or InStr(TakenHTML,"Register system prompt")>1 or InStr(TakenHTML,"not been registered yet")>1 or InStr(TakenHTML,"没有找到这个条目")>1 or InStr(TakenHTML,"此域名没有被注册")>1 or InStr(TakenHTML,"没有查到符合条件的记录")>1 or InStr(TakenHTML,"信息不存在")>1 or InStr(TakenHTML,"Not Exist")>1 or InStr(TakenHTML,"NOT FOUND")>1 or InStr(TakenHTML,"你所查询的信息不存在")>1 then%>
              以下域名以经注册了。。 
              <% 
			  else %>
              <br>
              以下是注册信息<br>
              <%=TakenHTML%></p>
            <p> 
              <% end if %>
            </p></td>
        </tr>
      </table></td>
  </tr>
</table>
<hr width="90%" size="1">
<p align="center">本程序由飞翔商务提供 <a href="http://www.fxsw.net%20">http://www.fxsw.net 
  </a>有任何问题请和我们联系 tel: 3574338822 QQ: 11752290</p>
<p align="center">200M asp的空间+100M企业邮局+国际域名(com/.net/.org)/支持动网论坛 350元一年</p>
<p align="center">此程序免费提供,有需要使用的请保留我们的版权</p>
<p>&nbsp;</p>
</body>
</html>
