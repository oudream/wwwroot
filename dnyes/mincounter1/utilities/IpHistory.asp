<%@LANGUAGE="JavaScript" CODEPAGE="936"%>
<script language="vbscript" runat="server">
	Function AnsiCode(vstrIn)
		Dim i, strReturn, innerCode, ThisChr
		Dim Hight8, Low8
		strReturn = "" 
		For i = 1 To Len(vstrIn) 
			ThisChr = Mid(vStrIn,i,1) 
			If Abs(Asc(ThisChr)) < &HFF Then 
				strReturn = strReturn & ThisChr 
			Else
				innerCode = Asc(ThisChr)
				If innerCode < 0 Then
					innerCode = innerCode + &H10000
				End If
				Hight8 = (innerCode And &HFF00) \ &HFF
				Low8 = innerCode And &HFF
				strReturn = strReturn & "%" & Hex(Hight8) & "%" & Hex(Low8)
			End If 
		Next 
		AnsiCode = strReturn 
	End Function
	
	Function DeCodeAnsi(s)
		Dim i, sTmp, sResult, sTmp1
		sResult = ""
		For i=1 To Len(s)
			If Mid(s,i,1)="%" Then
				sTmp = "&H" & Mid(s,i+1,2)
				If isNumeric(sTmp) Then
					If CInt(sTmp)=0 Then
						i = i + 2
					ElseIf CInt(sTmp)>0 And CInt(sTmp)<128 Then
						sResult = sResult & Chr(sTmp)
						i = i + 2
					Else
						If Mid(s,i+3,1)="%" Then
							sTmp1 = "&H" & Mid(s,i+4,2)
							If isNumeric(sTmp1) Then
								sResult = sResult & Chr(CInt(sTmp)*16*16 + CInt(sTmp1))
								i = i + 5
							End If
						Else
							sResult = sResult & Chr(sTmp)
							i = i + 2
						End If
					End If
				Else
					sResult = sResult & Mid(s,i,1)
				End If
			Else
				sResult = sResult & Mid(s,i,1)
			End If
		Next
		DeCodeAnsi = sResult
	End Function
</script>
<html>
<head>
<title>历史IP访问地列表 - COCOON Counter 6</title>
<style>
body{font-size:9pt;font-family:Verdana}
</style>
</head>
<body>
<u>COCOON Counter 6 IP归属地 历史记录</u>
<br><br>
<%
		function formatDateTime(d,t){
			function formatNumber(n){return (n+'').length==1?'0'+n:n;}
			var sY=d.getFullYear();sM=d.getMonth()+1,sD=d.getDate();
			var sH=formatNumber(d.getHours()),sN=formatNumber(d.getMinutes());sS=formatNumber(d.getSeconds());
			var sDate = sY+'-'+sM+'-'+sD, sTime = sH+':'+sN+':'+sS;
			switch(t){
				case 0: return sDate+' '+sTime;
				case 1: case 2: return sDate;
				case 3:	case 4: return sTime;
			}
		}


Response.write(DeCodeAnsi(Request.Cookies("cc_6_userIpSet")).replace(/\&/gim,'\r\n<br>\r\n'));
%>
<br><br><br>
-----------
<br>
Copyright(r) COCOON Studio 2000-2004.
</body>
</html>