<html>
<title>
�鿴���������
</title>
<body>
<%  Set bc = Server.CreateObject("MSWC.BrowserType") %>  
<center>
<font size=5 color=red>�鿴���������</font><hr>
<TABLE BORDER=1> 
<TR><TD width=50>�����</TD><TD width=50>  <%= bc.browser  %>  </td>
<TR><TD>�汾</TD><TD>  <%= bc.version  %>  </TD></TR> 
<TR><TD>���</TD><TD>
<%  if (bc.frames = TRUE) then  %>  TRUE
<%  else  %>  FALSE
<%  end if  %> </td></TR> 
<TR><TD>��</TD><TD>
<%  if (bc.tables = TRUE) then  %>  TRUE 
<%  else  %> FALSE
<%  end if  %> </TD></TR> 
<TR><TD>��������</TD><TD> 
<%  if (bc.BackgroundSounds = TRUE) then  %>  TRUE 
<%  else  %> FALSE
<%  end if  %> </TD></TR> 
<TR><TD>VBScript</TD><TD> 
<%  if (bc.vbscript = TRUE) then  %>  TRUE 
<%  else  %> FALSE
<%  end if  %> </TD></TR> 
<TR><TD>JScript</TD><TD> 
<%  if (bc.javascript = TRUE) then  %>  TRUE 
<%  else  %> FALSE
<%  end if  %> </TD></TR> 
</TABLE> 
</center>
</body>
</html> 