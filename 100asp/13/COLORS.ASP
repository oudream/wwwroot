<%@ LANGUAGE = VBScript %>
<%  Option Explicit %>
<%
Dim bgcolor, textcolor
'bgcolor 背景颜色
'textcolor 字体颜色
' 从Form表单中取得变量信息
bgcolor = Request.Form("bg_red") & Request.Form("bg_green") & Request.Form("bg_blue")
textcolor =  Request.Form("text_red") & Request.Form("text_green") & Request.Form("text_blue")
If len(bgcolor) = 6 Then
    bgcolor = " BGCOLOR=#" & bgcolor
Else
    bgcolor = ""
End If
If len(textcolor) = 6 Then
    textcolor = " COLOR=#" & textcolor
Else
    textcolor = ""
End If
%>
<TABLE BORDER="1" CELLSPACING="3" CELLPADDING="5"<%= bgcolor %>>
    <TR>	  
      <TD VALIGN="CENTER"><STRONG>
      <FONT SIZE="6"<%= textcolor %>>你喜欢哪种颜色</FONT></STRONG></TD>
   </TR>
</TABLE>
<FORM ACTION="colors.asp" METHOD="POST">
<TABLE BORDER="0" CELLSPACING="0" CELLPADDING="3">
<TR>
  <TD>&nbsp;</TD>
  <TD ALIGN="center">红</TD>
  <TD ALIGN="center">绿</TD>
  <TD ALIGN="center">蓝</TD>
</TR>
<TR>
<TD>背景颜色:</TD>
<TD>
<select NAME="bg_red">
<OPTION><%= Request.Form("bg_red") %></OPTION>
<OPTION>00</OPTION>
<OPTION>33</OPTION>
<OPTION>66</OPTION>
<OPTION>99</OPTION>
<OPTION>CC</OPTION>
<OPTION>FF</OPTION>
</select>
</TD>
<TD>
<select NAME="bg_green">
<OPTION><%= Request.Form("bg_green") %></OPTION>
<OPTION>00</OPTION>
<OPTION>33</OPTION>
<OPTION>66</OPTION>
<OPTION>99</OPTION>
<OPTION>CC</OPTION>
<OPTION>FF</OPTION>
</select>
</TD>
<TD>
<select NAME="bg_blue">
<OPTION><%= Request.Form("bg_blue") %></OPTION>
<OPTION>00</OPTION>
<OPTION>33</OPTION>
<OPTION>66</OPTION>
<OPTION>99</OPTION>
<OPTION>CC</OPTION>
<OPTION>FF</OPTION>
</select>
</TD>
</TR>
<TR>
<TD>字体颜色:</TD>
<TD>
<select NAME="text_red">
<OPTION><%= Request.Form("text_red") %></OPTION>
<OPTION>00</OPTION>
<OPTION>33</OPTION>
<OPTION>66</OPTION>
<OPTION>99</OPTION>
<OPTION>CC</OPTION>
<OPTION>FF</OPTION>
</select>
</TD>
<TD>
<select NAME="text_green">
<OPTION><%= Request.Form("text_green") %></OPTION>
<OPTION>00</OPTION>
<OPTION>33</OPTION>
<OPTION>66</OPTION>
<OPTION>99</OPTION>
<OPTION>CC</OPTION>
<OPTION>FF</OPTION>
</select>
</TD>
<TD>
<select NAME="text_blue">
<OPTION><%= Request.Form("text_blue") %></OPTION>
<OPTION>00</OPTION>
<OPTION>33</OPTION>
<OPTION>66</OPTION>
<OPTION>99</OPTION>
<OPTION>CC</OPTION>
<OPTION>FF</OPTION>
</select>
</TD>
</TR>
</TABLE>
<INPUT TYPE="submit" VALUE="显示效果">
</FORM>

