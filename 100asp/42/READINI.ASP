<%
class IniFile
    private mFileName
    private mDictionary
    
    private sub class_initialize()
        set mDictionary=Server.CreateObject("Scripting.Dictionary")
    end sub
    
    public property Let FileName(vFileName)
        dim mFileSys 
        dim mIniFile
        dim mSection
        dim mKeyName
        dim mValue
        dim mString
        dim mStart
        mFileName=vFileName
        set mFileSys=server.CreateObject("Scripting.FileSystemObject")
        set mIniFile=mfilesys.OpenTextFile(mFileName,1)
        do while not mIniFile.AtEndOfStream
            mString=mIniFile.ReadLine
            if mString<>"" then
                if left(mString,1)<>";" then
                    if left(mString,1)="[" then
                        mSection=mid(mString,2,len(mstring)-2)
                    else
                        mStart=instr(mstring,"=")
                        mKeyName=left(mstring,mStart-1)
                        mDictionary.Add lcase(mSection) & "_" & lcase(mKeyName),mid(mString,mStart+1)
                    end if
                end if
            end if
        loop
        mIniFile.Close
        Set mFileSys=nothing
        set mIniFile=nothing
    end property
    
    Public Function GetIniValue(byval vSection,byval vKeyName)
        on error resume next
        GetIniValue=mDictionary(lcase(vSection) & "_" & lcase(vKeyName))
    End Function
    
End Class
%>
<Html>
<Title>读取INI文件</Title>
<Body>
<h2>读取INI文件</h2>
<%
dim mINI 
set mINI=new inifile
mINI.FileName = "d:\1.ini"
Response.Write "[Company]" & "<br>Phone="
Response.Write  mINI.GetIniValue("Company","phone")

%>
</Body>
</Html>    