<HTML>
<TITLE>
柱状图
</TITLE>
<BODY>
<SCRIPT LANGUAGE="VBScript" RUNAT="SERVER"> 
function MakeColumn(title, numarray, labelarray, maxheight, maxwidth)  
 
 dim ColumnString  
 dim max  
 dim maxlength 
 dim tempnumarray 
 dim templabelarray 
 dim heightarray 
 Dim colorarray 
 Dim multiplier 
 
 if maxheight > 0 and maxwidth > 0 and ubound(labelarray) = ubound(numarray) then 
  colorarray = array("red","blue","yellow","navy","orange","purple","green") 
  templabelarray = labelarray 
  tempnumarray = numarray 
  heightarray = array() 
  max = 0 
  maxlength = 0 
  ColumnString = "<TABLE bgcolor='gold' border='6'><tr><td><TABLE border='0' cellspacing='1' cellpadding='0'>" & vbCrLf 
  for each stuff in tempnumarray 
   if stuff > max then max = stuff end if  
  next 
  multiplier = maxheight/max 
  for counter = 0 to ubound(tempnumarray) 
   if tempnumarray(counter) = max then  
    redim preserve heightarray(counter) 
    heightarray(counter) = maxheight 
   else 
    redim preserve heightarray(counter)  
    heightarray(counter) = tempnumarray(counter) * multiplier  
   end if  
  next  

  ColumnString = ColumnString & "<TR><TH colspan='" & ubound(tempnumarray)+1 & "'>" & _ 
     "<FONT SIZE='1'><U>" & title & "</TH></TR>" & vbCrLf & "<TR>" & vbCrLf 
   for counter = 0 to ubound(tempnumarray)  
    ColumnString = ColumnString & vbTab & "<TD valign='bottom' align='center' >" & _ 
    "<FONT SIZE='1'><table border='0' cellpadding='0' width='" & maxwidth & "'><tr><tr><td valign='bottom' bgcolor='"  
    ColumnString = ColumnString & colorarray(counter mod (ubound(colorarray)+1)) 
    ColumnString = ColumnString & "' height='" & round(heightarray(counter),2) & "'></td></tr></table>" 
    ColumnString = ColumnString & "<BR>" & tempnumarray(counter) 
    ColumnString = ColumnString & "</TD>" & vbCrLf 
   next 
  
  ColumnString = ColumnString & "</TR>" & vbCrLf 
 
  for each stuff in labelarray 
   if len(stuff) >= maxlength then maxlength = len(stuff) 
  next 
  
  for each stuff in labelarray 
   ColumnString = ColumnString & vbTab & "<TD align='center'><FONT SIZE='1'><B> "  
   for count = 0 to round((maxlength - len(stuff))/2) 
    ColumnString = ColumnString & " " 
   next 
   if maxlength mod 2 <> 0 then ColumnString = ColumnString & " " 
   ColumnString = ColumnString & stuff  
   for count = 0 to round((maxlength - len(stuff))/2) 
    ColumnString = ColumnString & " " 
   next 
   ColumnString = ColumnString & " </TD>" & vbCrLf 
  next 
    
  ColumnString = ColumnString & "</TABLE></td></tr></table>" & vbCrLf 
  MakeColumn = ColumnString 
 else 
  Response.Write "柱状图函数参数有错" 
 end if  
end function 

dim stuff 
dim labelstuff 

stuff = Array(72,39,60,42) 
labelstuff = Array("北京", "上海","广州","重庆") 
Response.Write MakeColumn("演示", stuff, labelstuff, 150,30) 

</SCRIPT> 
</BODY>
</HTML>