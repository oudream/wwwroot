<%@ LANGUAGE="VBSCRIPT" %>
<!--#include file=config.asp -->


<script language="javascript">
<!--
function isok(theform)
{
    if (theform.placename.value=="")
    {
        alert("����д���λ��ʶ��");
        theform.placename.focus();
        return (false);
    }

}
-->
</script>
<script language=JavaScript>
<!--

function opw(url,name, width, height) { //v2.0
window.open(url,name,''+'width='+width+',height='+height+',scrollbars=yes'+'');
}
//-->
</script>
<!--#include file="top.asp"-->
<center>
  <table border="0" align="center"  width=100%  cellpadding="0" cellspacing="0" height="1">
    <tr> 
           <td valign="top">
        <table border=0 width=98% cellspacing=0 cellpadding=0>
          <tr bgcolor=#ffffff align=center valign=top> 
            <td> 
              <table border=0 width=100% cellspacing=0 cellpadding=2 style="border-collapse: collapse" bordercolor="#111111">
				<form method="POST"  action=?job=add onSubmit="return isok(this)">

              <%
if request.querystring("job")="add" then

adsconn.open adsdata
 
set adsrs=server.createobject("adodb.recordset")

if  request("place")="0" then
adssql="select * from place "
adsrs.open adssql,adsconn,1,3
adsrs.AddNew
else
adssql="select * from place where place="&trim(request("place"))
adsrs.open adssql,adsconn,1,3
end if
adsrs(1) = trim(request("placename"))
adsrs(2)= trim(request("placelei"))
adsrs(3)= trim(request("placehei"))
adsrs(4)= trim(request("placewid"))
adsrs.update
adsrs.close
set adsrs=nothing
adsconn.close
set adsconn=nothing
response.redirect "?"

end if

if request.querystring("job")="del" then
if  isnumeric(request("place"))=true then
adsconn.open adsdata
set adsrs=server.createobject("adodb.recordset")
adssql="select * from place where place="&trim(request("place"))
adsrs.open adssql,adsconn,3,3
adsrs.delete
adsrs.close
set adsrs=nothing
adsconn.close
set adsconn=nothing
response.redirect "?"
end if
end if
%>
              <tr bgcolor=#ffffff> 
                <td>��</td>
                <td> <input type=hidden name=place value="0" >
                  <p align="center">���λ��ʶ&nbsp;<input type=text name=placename size=30 maxlength=30>
                  <font color="#FF0000">15������ </font>�߶�<input type=text name=placehei value="60" size=3 maxlength=30> 
                  ���<input type=text name=placewid value="468" size=3 maxlength=30>&#12288;���� 
 <%Call Ggwlei(1)%>&nbsp; 
                  <input type=submit value=�������λ name=B1></p>
                <p align="center">
                  <font color="#808080">*** �뾡����Ҫʹ����ͬ�Ĺ��λ��ʶ���߶ȡ������ҪӦ���ڵ������ڴ�С�������������ã����ʵ����ã�����Ϊ��</font></td>
              </tr>
             
            </form>
          </table>
        </td>
      </tr>
    </table>
  </td>
    </tr>
  </table>
  
  <hr color="#808080" size="1">
<p> <font color="#FF0000">�����뷽����</font>���±����ݷŵ�Ԥ�����λ�ã��������е�<font color="#FF0000">���λID</font>��Ӧ��ȷ 
<font color="#808080">���ڹ��λ�б��в鿴</font><font color="#FF0000">���λID</font><br><input type="text" name="T1" size="100" value="<script language=javascript src=<%=DqUrl()%>/ad.asp?i=���λID></script>"></p>
<p align="center"> <font color=red><b>���й��λ�б�</b></font></p>
<p align=left><font color="#808080">*** ���ڸ߶ȡ������������ʵ�<b>���ֻ�ٷֱȻ�Ϊ���Զ�</b></font></p>
<div align="center">
  <center>
  <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#C0C0C0" width="100%" id="AutoNumber1">
    <tr>
      <td width="80" align="center" height="20" bgcolor="#F5F5F5">
      <font color="#FF0000">���λID</font></td>
      <td align="center" height="20" bgcolor="#F5F5F5">���λ��ʶ</td>
      <td align="center" height="20" bgcolor="#F5F5F5">�߶�</td>
      <td align="center" height="20" bgcolor="#F5F5F5">���</td>
      <td width="20%" align="center" height="20" bgcolor="#F5F5F5">���λ��ʾ��ʽ</td>
      <td width="30%" align="center" height="20" bgcolor="#F5F5F5">�� ��</td>
    </tr>
    
   <%
adsconn.open adsdata
set adsrs=server.createobject("adodb.recordset")
adssql="select * from place"
adsrs.open adssql,adsconn,1,1
do while not adsrs.eof 

%><form method="POST" action="?job=add"  onSubmit="return isok(this)">
    <tr>
     
      <td width="80" align="center"><font color=red><%=adsrs(0)%></font><input type=hidden name=place value="<%=adsrs(0)%>" >��</td>
      <td align="center"><input type=text name=placename value="<%=adsrs(1)%>" size=30 maxlength=30></td>
      <td align="center">
      <input type=text name=placehei value="<%=adsrs(3)%>" size=3 maxlength=30></td>
      <td align="center">
      <input type=text name=placewid value="<%=adsrs(4)%>" size=3 maxlength=30></td>
      <td width="20%" align="center">
     <%Call Ggwlei(adsrs("placelei"))%>&nbsp; 
</td>
      <td width="30%" align="center"><input type="submit" value="�޸�" name="B1">&nbsp;&nbsp; 
      <a href=?job=del&place=<%=adsrs(0)%>>ɾ��</a>&nbsp; <a href=list.asp?type=place&place=<%=adsrs(0)%>>���й����</a> <a href=javascript:opw('option.asp?id=<%=adsrs(0)%>&job=yulanggw','',800,600)>Ԥ�����</a></td>
    </tr>
    </form>
    <%adsrs.movenext
      loop
      adsrs.close
      set adsrs=nothing
adsconn.close
set adsconn=nothing
      %>
  </table>
  </center>
</div>
��
  <p align="left">
  <p align="left"><hr color="#808080" size="1">
<p align="left"><font color="#FF0000"><a name="˵��">���λ��ʾ��ʽ˵��</a>��</font></p>
<center>
  </p>
  <ul>
    <li><p align="left">
    ҳ��Ƕ��ѭ�������ǽ����λֱ������ĳҳ��һ�̶�λ�ã�����ͬһλ��ѭ����ʾ���λ�е����������������������ÿˢ��һ�ξͻ������ʾһ���µĹ����</p>
    </li>
    <li><p align="left">�����������룺���ϵ������Ź��λ�е��������������</p>
    </li>
    <li><p align="left">�����������룺�����Һ��Ź��λ�е��������������</p>
    </li>
    <li><p align="left">���Ϲ������룺���Ϲ�����ʾ���λ�е��������������</p>
    </li>
    <li><p align="left">����������룺���������ʾ���λ�е��������������</p>
    </li>
    <li><p align="left">����������ڣ�ҳ���ʱͬʱ����������ڣ�ÿ����������ʾһ��������������������ù��λ�е������������һ��</p>
    </li>
    <li><p align="left">
    ѭ���������ڣ�ҳ���ʱͬʱ����һ�����ڣ���ͬһ������ѭ����ʾ���λ�е�������棬������ÿˢ��һ�ξͻ��ڵ��������и�����ʾһ���µĹ����</p>
    </li>
  </ul>
  <!--#include file="boot.asp"-->
<p><br><br>
       

        ��</p>