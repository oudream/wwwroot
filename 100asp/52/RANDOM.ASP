<%@ LANGUAGE = VBScript %> 
<%  Option Explicit     %>   
<HTML>     
<HEAD>         
<TITLE>��Tools��������������</TITLE>     
</HEAD>      
<BODY BGCOLOR="White" topmargin="10" leftmargin="10">           
<!-- Display Header -->          
<font size="4" face="Arial, Helvetica">         
<b>��Tools��������������</b></font><br>                
<hr size="1" color="#000000">  
<% Dim Tools  
Dim FoundFile, FileExists, FileDoesNotExist, NotFoundFile 
Dim RandInt, RandPosInt, RandIntBelow  
Set Tools = Server.CreateObject("MSWC.Tools") %>   
<h3>�������ʵ��</h3>  
<% ' Random Integer  
RandInt = Tools.Random 
Response.Write "�������: " 
Response.Write RandInt 
Response.Write "<BR>"   

RandPosInt = Abs( Tools.Random ) 
Response.Write "��������ľ���ֵ: " 
Response.Write RandPosInt 
Response.Write "<BR>"   

RandIntBelow =  Abs( Tools.Random ) Mod 200 
Response.write "С��200���������: " 
Response.write RandIntBelow 
Response.write "<BR>" %>     
</BODY> </HTML> 