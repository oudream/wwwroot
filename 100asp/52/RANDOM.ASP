<%@ LANGUAGE = VBScript %> 
<%  Option Explicit     %>   
<HTML>     
<HEAD>         
<TITLE>用Tools组件生成随机整数</TITLE>     
</HEAD>      
<BODY BGCOLOR="White" topmargin="10" leftmargin="10">           
<!-- Display Header -->          
<font size="4" face="Arial, Helvetica">         
<b>用Tools组件生成随机整数</b></font><br>                
<hr size="1" color="#000000">  
<% Dim Tools  
Dim FoundFile, FileExists, FileDoesNotExist, NotFoundFile 
Dim RandInt, RandPosInt, RandIntBelow  
Set Tools = Server.CreateObject("MSWC.Tools") %>   
<h3>随机整数实例</h3>  
<% ' Random Integer  
RandInt = Tools.Random 
Response.Write "随机整数: " 
Response.Write RandInt 
Response.Write "<BR>"   

RandPosInt = Abs( Tools.Random ) 
Response.Write "随机整数的绝对值: " 
Response.Write RandPosInt 
Response.Write "<BR>"   

RandIntBelow =  Abs( Tools.Random ) Mod 200 
Response.write "小于200的随机整数: " 
Response.write RandIntBelow 
Response.write "<BR>" %>     
</BODY> </HTML> 