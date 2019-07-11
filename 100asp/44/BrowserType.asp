<html>
<title>
查看浏览器属性
</title>
<body>
<%  Set bc = Server.CreateObject("MSWC.BrowserType") %>  
<center>
<font size=5 color=red>查看浏览器属性</font><hr>
<TABLE BORDER=1> 
<TR><TD width=50>浏览器</TD><TD width=50>  <%= bc.browser  %>  </td>
<TR><TD>版本</TD><TD>  <%= bc.version  %>  </TD></TR> 
<TR><TD>框架</TD><TD>
<%  if (bc.frames = TRUE) then  %>  TRUE
<%  else  %>  FALSE
<%  end if  %> </td></TR> 
<TR><TD>表单</TD><TD>
<%  if (bc.tables = TRUE) then  %>  TRUE 
<%  else  %> FALSE
<%  end if  %> </TD></TR> 
<TR><TD>背景音乐</TD><TD> 
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