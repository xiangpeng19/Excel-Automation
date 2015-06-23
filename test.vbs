Option Explicit
Private Declare Function FindWindowEx Lib "User32" Alias "FindWindowExA" (ByVal hWnd1, ByVal hWnd2, ByVal lpsz1, ByVal lpsz2) As Long 

Declare Function GetClassName Lib "User32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Declare Function IIDFromString Lib "ole32" (ByVal lpsz As Long, ByRef lpiid As UUID) As Long
Declare Function AccessibleObjectFromWindow Lib "oleacc" (ByVal hWnd As Long, ByVal dwId As Long, ByRef riid As UUID, ByRef ppvObject As Object) As Long

Type UUID 'GUID
  Data1 As Long
  Data2 As Integer
  Data3 As Integer
  Data4(7) As Byte
End Type

Type TargetWBType
    name As String
    returnObj As Object
    returnApp As Excel.Application
    returnWBIndex As Integer
End Type

Option Explicit