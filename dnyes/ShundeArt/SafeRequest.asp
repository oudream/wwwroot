<%
'SQLע�����
Function SafeRequest(ParaName,ParaType)
       '--- ������� ---
       'ParaName:��������-�ַ���
       'ParaType:��������-������(1��ʾ���ϲ��������֣�0��ʾ���ϲ���Ϊ�ַ�)

       Dim ParaValue
       ParaValue=Request(ParaName)
       If ParaType=1 then
              If not isNumeric(ParaValue) then
                     Response.write "����" & ParaName & "����Ϊ�����ͣ�"
                     Response.end
              End if
       Else
              ParaValue=replace(ParaValue,"'","''")
       End if
       SafeRequest=ParaValue
End function
%>