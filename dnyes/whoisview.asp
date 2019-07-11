<%@ language=vbscript%>
<!--#include file="whoisclass.asp"-->
<%
on error resume next 
domain=request.QueryString("domain")
after=request.QueryString("after")

     set list=new whoisclass
       list.get_domain=domain&trim(after)
	   list.get_after=trim(after)
       result=list.exsit
	   domaindetail=list.gettakenhtml
     set list=nothing
  %>
<table width="96%" border="0" align="center" cellpadding="3" cellspacing="1" class="pt9">
  <tr> 
    <td height="33" align="center">
      <p align="left">以下是查找到的关于www.<%=domain%><%=after%>信息</p>
    </td>
  </tr>
  
  <tr>
    <td height="140" valign="top" align="left" bgcolor="#FFFFFF"><%=domaindetail%></td>
  </tr>
  
</table>