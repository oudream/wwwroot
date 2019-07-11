<html>
<head>
<title>网上购物</title>
</head>
<body bgcolor=lightgreen>
<div align="center">
<p align=center><FONT face=华文行楷 size=6> 
  实例90&nbsp;&nbsp; 网上购物注册表</FONT>   </p>
<script language="VBScript">
<!--
function checkphone(str1)
dim i ,tmp_char,tmp_str
tmp_str=trim(str1)
for i=1 to len(tmp_str) 
tmp_char=mid(tmp_str,i,1)
if (asc(tmp_char)<asc("0"))or (asc(tmp_char)>asc("9")) then
checkphone=false
exit function
end if 
next
checkphone=true
end function
function checkmail(str1)
dim break,pre,post
break=inStr(1,str1,"@",0)
if break=0 then 
checkmail=false
exit function
end if 
checkmail=true
end function

sub submit1_onclick
if frm1.txt1.value="" then
msgbox "姓名不得为空,请重新输入！",0+16,"Warning"
exit sub
end if 
if frm1.txt2.value="" then
msgbox "密码不能为空，请重新输入！",0+16,"Warning"
exit sub
end if 
if frm1.txt3.value="" then
msgbox "联系地址不得为空，请重新输入！",0+16,"Warning"
exit sub
end if 
if (frm1.txt4.value="") or (checkphone(frm1.txt4.value)=false) then
msgbox "电话号码区号不得为空也不得为非数字符号，请重新输入！",0+16,"Warning"
exit sub
end if 
if (frm1.txt5.value="") or (checkphone(frm1.txt5.value)=false) or (len(trim(frm1.txt5.value))<6) then
msgbox "电话号码不得为空也不得为非数字符号，长度不得小于6位，请重新输入！",0+16,"Warning"
exit sub
end if 
if (frm1.txt6.value<>"") and (checkmail(frm1.txt6.value)=false) then
msgbox "E-mail地址输入错误，请重新输入！",0+16,"Warning"
exit sub
end if
frm1.submit 
end sub
-->
</script>
<form method="post" action="adduser.asp" name="frm1">
      <TABLE border=1 cellPadding=1 cellSpacing=1 width="600">
           <tr><td colspan=4 align=middle><FONT face=华文新魏 size=+2>个人信息表</FONT></td></tr>
        <TR>
          <TD>姓名：</TD>
          <TD><INPUT name=txt1></TD>
          <TD>年龄： </TD>
          <TD><SELECT name=sel1> <OPTION>18岁以下</OPTION> <OPTION 
              selected>18~35岁</OPTION> <OPTION>35~60岁</OPTION> 
              <OPTION>60岁以上</OPTION></SELECT> </TD></TR>
        <TR>
          <TD>password：</TD>
          <TD><INPUT type="password" name=txt2></TD>
          <TD>联系地址：</TD>
          <TD><INPUT name=txt3></TD></TR>
        <TR>
          <TD>联系电话： </TD>
          <TD><INPUT maxLength=4 name=txt4 size=4 value=010>- <INPUT 
            maxLength=8 name=txt5 size=8 style="HEIGHT: 22px; WIDTH: 87px"> </TD>
          <TD>E-mail: </TD><TD><INPUT name=txt6></TD></TR>
        <tr><td colspan=4 align=middle><INPUT name=submit1 type=button value=用户注册>
     </td></tr></TABLE>
</div></FORM>
</body>
</html>
