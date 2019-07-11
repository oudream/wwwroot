<%
Response.Expires=0
Function bin2str(binstr) 
    Dim varlen,clow,ccc,skipflag 

    skipflag=0 
    ccc = "" 
    If Not IsNull(binstr) Then 
        varlen=LenB(binstr) 
        For i=1 To varlen 
            If skipflag=0 Then 
                clow = MidB(binstr,i,1)
                If AscB(clow) > 127 Then 
                    ccc =ccc & Chr(AscW(MidB(binstr,i+1,1) & clow)) 
                    skipflag=1 
                Else 
                    ccc = ccc & Chr(AscB(clow)) 
                End If 
            Else 
                skipflag=0 
            End If 
        Next 
    End If 
    bin2str = ccc 
End Function 


varByteCount = Request.TotalBytes
bnCRLF = chrB( 13 ) & chrB( 10 )
binHTTPHeader=Request.BinaryRead(varByteCount)        
Divider = LEFTB( binHTTPHeader,  INSTRB( binHTTPHeader, bnCRLF ) - 1 )

'开始读非文件域的数据
Do while lenB(binHTTPHeader)>46
    
    binHeaderData = LeftB(binHTTPHeader, INSTRB( binHTTPHeader, bnCRLF & bnCRLF )-1)
    strHeaderData=bin2str(binHeaderData)

    lngFieldNameStart=Instr(strHeaderData,"name="&chr(34))+Len("name="&chr(34))
    lngFieldNameEnd=Instr(lngFieldNameStart,strHeaderData,chr(34))
    
    
    strFieldName=Mid(strHeaderData,lngFieldNameStart,lngFieldNameEnd-lngFieldNameStart)
    strFieldName=Trim(strFieldName)
    strFieldName=Replace(strFieldName,vbcrlf,vbnullstring)
    
        '判断文件数据时候开始
    If strComp(strFieldName,"FileUploadStart",1)=0 Then
        binHTTPHeader=MIDB(binHTTPHeader,INSTRB( DataStart + 1, binHTTPHeader, divider ))
        exit do
    End if
    
    DataStart = INSTRB( binHTTPHeader, bnCRLF & bnCRLF ) + 4 
    DataEnd = INSTRB( DataStart + 1, binHTTPHeader, divider ) - DataStart

    binFieldValue=MIDB( binHTTPHeader, DataStart, DataEnd )
    strFieldValue=bin2str(binFieldValue)
    strFieldValue=Trim(strFieldValue)
    strFieldValue=Replace(strFieldValue,vbcrlf,vbnullstring)

    '非文件上传域变量赋值
    execute strFieldName&"="""&strFieldValue&""""
    
        
    binHTTPHeader=MIDB(binHTTPHeader,INSTRB( DataStart + 1, binHTTPHeader, divider ))
        
loop

'开始处理文件数据
i=0
Do while lenB(binHTTPHeader)>46
    ii=ii+1
    
    binHeaderData = LeftB(binHTTPHeader, INSTRB( binHTTPHeader, bnCRLF & bnCRLF )-1)
        
    strHeaderData=bin2str(binHeaderData)
    
    '读取上传文件的Content-Type
    lngFileContentTypeStart=Instr(strHeaderData,"Content-Type:")+Len("Content-Type:")
    strFileContentType=Trim(Mid(strHeaderData,lngFileContentTypeStart))
    strFileContentType=Replace(strFileContentType,vbCRLF,vbNullString)
    
    '读取上传的文件名
    lngFileNameStart=Instr(strHeaderData,"filename="&chr(34))+Len("filename="&chr(34))
    lngFileNameEnd=Instr(lngFileNameStart,strHeaderData,chr(34))
    strFileName=Mid(strHeaderData,lngFileNameStart,lngFileNameEnd-lngFileNameStart)
    strFileName=Trim(strFileName)
    strFileName=Replace(strFileName,vbCRLF,vbNullString)
    
    '读取上传文件数据
    DataStart = INSTRB( binHTTPHeader, bnCRLF & bnCRLF ) + 4 
    DataEnd = INSTRB( DataStart + 1, binHTTPHeader, divider ) - DataStart
    
    If strFileName<>"" Then
                
        binFieldValue=MIDB( binHTTPHeader, DataStart, DataEnd )
        
        '将上传的文件写入数据库
%>
<!--#include file="SHEI45FECSA.ASP" -->
<%
        SQL="select * from subs where id=" & id
        set rs=server.CreateObject("ADODB.Recordset")

rs.open sql,conn,1,3
if ii=1 then 
rs.update
rs("tu").AppendChunk binFieldValue
rs.update
end if
if ii=2 then 
rs.update
rs("stu").AppendChunk binFieldValue
rs.update
end if

        rs.close
        set rs=Nothing
        conn.Close
        set conn=Nothing
        
    End if
    
    binHTTPHeader=MIDB(binHTTPHeader,INSTRB( DataStart + 1, binHTTPHeader, divider ))
    
loop
%>
<%
'url="showsubs.asp?id=" & id
'response.redirect url
response.write"<SCRIPT language=JavaScript>alert('商品资料已添加或修改成功！');"
response.write"javascript:history.go(-2)</SCRIPT>"
%>
