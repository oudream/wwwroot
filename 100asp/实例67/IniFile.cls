VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IniFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Public Function GetIniFileString(ByVal FilePath As String, ByVal Section As String, ByVal KeyName As String) As String
    Dim lpReturnedString As String * 100
    Dim nSize As Long
    Dim lPo As Long
    Dim mFilePath, mSection, mKeyName As String
    mSection = Trim(Section)
    mKeyName = Trim(KeyName)
    mFilePath = Trim(FilePath)
    nSize = Len(lpReturnedString)
    lPo = GetPrivateProfileString(mSection, mKeyName, "", lpReturnedString, nSize, mFilePath)
    GetIniFileString = Mid(lpReturnedString, 1, lPo)
End Function

