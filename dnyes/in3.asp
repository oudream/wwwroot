<!--#include file="whoisclass.asp"-->
<% on error resume next 
  dim after,domain,result
  after=request.form("after")
  domain=trim(request.form("domain"))
   
if domain<>"" then
     set list=new whoisclass
       list.get_domain=domain&trim(after)
	   list.get_after=trim(after)
       result=list.exsit
	   exploration=list.gettakenhtml
     set list=nothing
end if	 
%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
	<tr> 
      <td colspan="5"> 
<% 

if result=0 then
resultstr=domain&after&"  ��û��ע��"
resultlink="�������йظ������Ͳ�ѯ������һЩ���"
else
resultstr=domain&after&"  �ѱ�ע��"
resultlink="�����ǹ��ڸ���������ϸ˵��"
end if

if domain<>"" then
resultlinkstr=resultlink
logonstr=resultstr
else
resultlinkstr=""
logonstr="��׼ȷ������Ҫ��ѯ������"
end if

%>
	  </td>
	</tr>
	<tr><td colspan="5" align="center"><%=request("domaintypestr")%> </td></tr>
	<tr><td colspan="5" align="center"><%=logonstr%> </td></tr>
	<tr>
    <td colspan="5" align="center"><%=resultlinkstr%></td></tr>	
	<tr>
    <td colspan="5"><% if domain<>"" then %> <%=exploration%> <%end if%></td></tr>	

</table>
