<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--media=print ������Կ����ڴ�ӡʱ��Ч--> 
<style media=print> 
.Noprint{display:none;} 
.PageNext{page-break-after: always;} 
</style> 

<style> 
.tdp 
{ 
border-bottom: 1 solid #000000; 
border-left: 1 solid #000000; 
border-right: 0 solid #ffffff; 
border-top: 0 solid #ffffff; 
} 
.tabp 
{ 
border-color: #000000 #000000 #000000 #000000; 
border-style: solid; 
border-top-width: 2px; 
border-right-width: 2px; 
border-bottom-width: 1px; 
border-left-width: 1px; 
} 
.NOPRINT { 
font-family: "����"; 
font-size: 9pt; 
} 

</style>
<title></title></head>
<body topmargin="0">
<!--#include file="art_adminconn.asp" -->
  <table width="550" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td>
	  <center class="Noprint" >
    <OBJECT id=WebBrowser classid=CLSID:8856F961-340A-11D0-A96B-00C04FD705A2 height=0 width=0>
    </OBJECT>
    <input type=button value=��ӡ onclick=document.all.WebBrowser.ExecWB(6,1)>&nbsp;
    <input type=button value=ֱ�Ӵ�ӡ onclick=document.all.WebBrowser.ExecWB(6,6)>&nbsp;
    <input type=button value=ҳ������ onclick=document.all.WebBrowser.ExecWB(8,1)>&nbsp;
    <input name="button" type=button onClick=document.all.WebBrowser.ExecWB(7,1) value=��ӡԤ��>
      </center> 
</td>
    </tr>
  </table>
<%
sql="select * from art_artist where artist_id="&request("artist_id") 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1 

if rs.eof or rs.bof then 
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' None data ! ');</SCRIPT>")
response.End()
end if
%>
  <TABLE width=650 
                  border=0 align="center" cellPadding=5 cellSpacing=1 bgcolor="#000000" class="tabp">
    <TBODY>
      <TR bgcolor="#FFFFFF" class="tdp"> 
        
      <TD width="20%" align="right" class="tdp"> ȫ �ƣ�</TD>
        <TD class="tdp"> 
          <%=rs("allname")%> </TD>
      </TR>
      <TR bgcolor="#FFFFFF" class="tdp"> 
        
      <TD align="right" class="tdp"> �� �ţ�</TD>
        <TD class="tdp"> 
          <%=rs("simname")%>
          </TD>
      </TR>
      <TR bgcolor="#FFFFFF" class="tdp"> 
        
      <TD align="right" class="tdp"> �� ��</TD>
        <TD class="tdp"> 
		<%=rs("sex")%>
          </TD>
      </TR>
      <TR bgcolor="#FFFFFF" class="tdp"> 
        
      <TD align="right" class="tdp"> �� �壺</TD>
        <TD class="tdp">  
          <%=rs("nation")%>
          </TD>
      </TR>
      <TR bgcolor="#FFFFFF" class="tdp"> 
        
      <TD align="right" class="tdp"> ԭ ����</TD>
        <TD class="tdp"> 
          <%=rs("native")%>
          </TD>
      </TR>
      <TR bgcolor="#FFFFFF" class="tdp"> 
        
      <TD align="right" class="tdp"> �������£�</TD>
        <TD class="tdp"> 
          <%=rs("born_time")%>
          </TD>
      </TR>
      <TR bgcolor="#FFFFFF" class="tdp"> 
        
      <TD align="right" class="tdp"> ��ҵԺУ��</TD>
        <TD class="tdp"> 
          <%=rs("graduation")%>
          </TD>
      </TR>
      <TR bgcolor="#FFFFFF" class="tdp"> 
        
      <TD align="right" class="tdp"> ��ѧרҵ��</TD>
        <TD class="tdp"> 
          <%=rs("special")%>
          </TD>
      </TR>
      <TR bgcolor="#FFFFFF" class="tdp"> 
        
      <TD align="right" class="tdp"> ����ְ��</TD>
        <TD class="tdp"> 
          <%=rs("job_now")%>
          </TD>
      </TR>
      <TR bgcolor="#FFFFFF" class="tdp"> 
        
      <TD align="right" class="tdp"> ����ְ��</TD>
        <TD class="tdp"> 
         <%=rs("job_old")%>
          </TD>
      </TR>
      <TR bgcolor="#FFFFFF" class="tdp"> 
        
      <TD align="right" class="tdp"> ְ �ƣ�</TD>
        <TD class="tdp"> 
          <%=rs("job_rank")%>
          </TD>
      </TR>
      <TR bgcolor="#FFFFFF" class="tdp"> 
        
      <TD align="right" class="tdp"> �� ����</TD>
        <TD class="tdp"> 
          <%=rs("expert")%>
          </TD>
      </TR>
      <TR bgcolor="#FFFFFF" class="tdp"> 
        
      <TD align="right" class="tdp"> �����</TD>
        <TD class="tdp"> 
          <%=rs("awarded")%>
          </TD>
      </TR>
      <TR bgcolor="#FFFFFF" class="tdp"> 
        
      <TD align="right" class="tdp"> ������Ʒ��</TD>
        <TD class="tdp"> 
          <%=rs("standest")%>
          </TD>
      </TR>
      <TR bgcolor="#FFFFFF" class="tdp"> 
        
      <TD align="right" class="tdp"> ��Ȥ���ã�</TD>
        <TD class="tdp"> 
          <%=rs("interest")%>
          </TD>
      </TR>
      <TR bgcolor="#FFFFFF" class="tdp"> 
        
      <TD align="right" class="tdp"> �� ַ��</TD>
        <TD class="tdp"> 
          <%=rs("addr")%>
          </TD>
      </TR>
      <TR bgcolor="#FFFFFF" class="tdp"> 
        
      <TD align="right" class="tdp"> �� ����</TD>
        <TD class="tdp"> 
          <%=rs("tel")%>
          </TD>
      </TR>
      <TR bgcolor="#FFFFFF" class="tdp"> 
        
      <TD align="right" class="tdp"> �� �棺</TD>
        <TD class="tdp"> 
          <%=rs("fax")%>
          </TD>
      </TR>
     <TR bgcolor="#FFFFFF" class="tdp"> 
        
      <TD align="right" class="tdp"> �� �ࣺ</TD>
        <TD class="tdp"> 
          <%=rs("zip")%>
          </TD>
      </TR>
      <TR bgcolor="#FFFFFF" class="tdp"> 
        
      <TD align="right" class="tdp"> �� �䣺</TD>
        <TD class="tdp"> 
          <%=rs("email")%>
          </TD>
      </TR>
      <TR bgcolor="#FFFFFF" class="tdp"> 
        
      <TD align="right" class="tdp"> �� ע��</TD>
        <TD height="350" valign="top" class="tdp"> 
          <%=rs("memo")%>
          </TD>
      </TR>
      
    </TBODY>
  </TABLE>
</body>
</html>
