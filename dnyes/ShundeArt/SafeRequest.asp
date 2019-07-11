<%
'SQL注入过滤
Function SafeRequest(ParaName,ParaType)
       '--- 传入参数 ---
       'ParaName:参数名称-字符型
       'ParaType:参数类型-数字型(1表示以上参数是数字，0表示以上参数为字符)

       Dim ParaValue
       ParaValue=Request(ParaName)
       If ParaType=1 then
              If not isNumeric(ParaValue) then
                     Response.write "参数" & ParaName & "必须为数字型！"
                     Response.end
              End if
       Else
              ParaValue=replace(ParaValue,"'","''")
       End if
       SafeRequest=ParaValue
End function
%>