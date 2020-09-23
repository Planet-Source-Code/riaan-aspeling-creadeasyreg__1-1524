<div align="center">

## cReadEasyReg


</div>

### Description

A easy way to read the Registry. Most of the times I work with the registry I only want to read it, not write to it. PLEASE NOTE: This is a class module and all the code should be paste into a CLASS Module.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Riaan Aspeling](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/riaan-aspeling.md)
**Level**          |Unknown
**User Rating**    |4.2 (163 globes from 39 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Registry](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/registry__1-36.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/riaan-aspeling-creadeasyreg__1-1524/archive/master.zip)

### API Declarations

```
' Developed by : Riaan Aspeling
' Company :  Altered Reality Corporation
' Date :   1999-Mar-21
' Country :  South Africa
'
' Description : A Easy way to READ the registry
' Comment :  Most of the times a work with the registry is only want
'     to READ it, not write to it. Hope you guys/gals out there
'     could use this code.
' Problems :  If you do find any problems (not Microsoft related) let me
'     know at :
'        arc@iti.co.za
'     Have fun reading the registry ;-)
Option Explicit
'API's to use
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal HKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal HKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal HKey As Long, ByVal dwIndex As Long, ByVal lpName As String, cbName As Long, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal HKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal HKey As Long) As Long
'Enum's for the OpenRegistry function
Public Enum HKeys
 HKEY_CLASSES_ROOT = &H80000000
 HKEY_CURRENT_USER = &H80000001
 HKEY_LOCAL_MACHINE = &H80000002
 HKEY_USERS = &H80000003
 HKEY_PERFORMANCE_DATA = &H80000004
 HKEY_CURRENT_CONFIG = &H80000005
 HKEY_DYN_DATA = &H80000006
End Enum
'Right's for the OpenRegistry
Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_CREATE_LINK = &H20
Private Const SYNCHRONIZE = &H100000
Private Const KEY_ALL_ACCESS = ( _
         ( _
         STANDARD_RIGHTS_ALL Or _
         KEY_QUERY_VALUE Or _
         KEY_SET_VALUE Or _
         KEY_CREATE_SUB_KEY Or _
         KEY_ENUMERATE_SUB_KEYS Or _
         KEY_NOTIFY Or _
         KEY_CREATE_LINK _
         ) _
         And _
         ( _
         Not SYNCHRONIZE _
         ) _
        )
'Local var's to keep track of things happening
Dim RootHKey As HKeys
Dim SubDir As String
Dim HKey As Long
Dim OpenRegOk As Boolean
```


### Source Code

```
'This function will return a array of variant with all the subkey values
'eg.
'  Dim MyVariant As Variant, MyReg As New CReadEasyReg, i As Integer
'  If Not MyReg.OpenRegistry(HKEY_LOCAL_MACHINE, "Software\Microsoft") Then
'   MsgBox "Couldn't open the registry"
'   Exit Sub
'  End If
'  MyVariant = MyReg.GetAllSubDirectories
'  For i = LBound(MyVariant) To UBound(MyVariant)
'   Debug.Print MyVariant(i)
'  Next i
'  MyReg.CloseRegistry
Function GetAllSubDirectories() As Variant
On Error GoTo handelgetdirvalues
 Dim SubKey_Num As Integer
 Dim SubKey_Name As String
 Dim Length As Long
 Dim ReturnArray() As Variant
 If Not OpenRegOk Then Exit Function
 'Get the Dir List
 SubKey_Num = 0
 Do
  Length = 256
  SubKey_Name = Space$(Length)
  If RegEnumKey(HKey, SubKey_Num, SubKey_Name, Length) <> 0 Then
   Exit Do
  End If
  SubKey_Name = Left$(SubKey_Name, InStr(SubKey_Name, Chr$(0)) - 1)
  ReDim Preserve ReturnArray(SubKey_Num) As Variant
  ReturnArray(SubKey_Num) = SubKey_Name
  SubKey_Num = SubKey_Num + 1
 Loop
 GetAllSubDirectories = ReturnArray
 Exit Function
handelgetdirvalues:
 GetAllSubDirectories = Null
 Exit Function
End Function
'This function will return a array of variant with all the value names in a key
'eg.
'  Dim MyVariant As Variant, MyReg As New CReadEasyReg, i As Integer
'  If Not MyReg.OpenRegistry(HKEY_LOCAL_MACHINE, "HardWare\Description\System\CentralProcessor\0") Then
'   MsgBox "Couldn't open the registry"
'   Exit Sub
'  End If
'  MyVariant = MyReg.GetAllValues
'  For i = LBound(MyVariant) To UBound(MyVariant)
'   Debug.Print MyVariant(i)
'  Next i
'  MyReg.CloseRegistry
Function GetAllValues() As Variant
On Error GoTo handelgetdirvalues
 Dim lpData As String, KeyType As Long
 Dim BufferLengh As Long, vname As String, vnamel As Long
 Dim ReturnArray() As Variant, Index As Integer
 If Not OpenRegOk Then Exit Function
 'Get the Values List
 Index = 0
 Do
  lpData = String(250, " ")
  BufferLengh = 240
  vname = String(250, " ")
  vnamel = 240
  If RegEnumValue(ByVal HKey, ByVal Index, vname, vnamel, 0, KeyType, lpData, BufferLengh) <> 0 Then
   Exit Do
  End If
  vname = Left$(vname, InStr(vname, Chr$(0)) - 1)
  ReDim Preserve ReturnArray(Index) As Variant
  ReturnArray(Index) = vname
  Index = Index + 1
 Loop
 GetAllValues = ReturnArray
 Exit Function
handelgetdirvalues:
 GetAllValues = Null
 Exit Function
End Function
'This function will return a specific value from the registry
'eg.
'  Dim MyString As String, MyReg As New CReadEasyReg, i As Integer
'  If Not MyReg.OpenRegistry(HKEY_LOCAL_MACHINE, "HardWare\Description\System\CentralProcessor\0") Then
'   MsgBox "Couldn't open the registry"
'   Exit Sub
'  End If
'  MyString = MyReg.GetValue("Identifier")
'  Debug.Print MyString
'  MyReg.CloseRegistry
Function GetValue(ByVal VarName As String) As String
On Error GoTo handelgetavalue
 Dim i As Integer
 Dim SubKey_Value As String, TempStr As String
 Dim Length As Long
 Dim value_type As Long
 If Not OpenRegOk Then Exit Function
 'Read the value
 Length = 256
 SubKey_Value = Space$(Length)
 If RegQueryValueEx(HKey, VarName, 0&, value_type, ByVal SubKey_Value, Length) <> 0 Then
  GetValue = ""
  Exit Function
 End If
 Select Case value_type
  Case 1 'Text
   SubKey_Value = Left$(SubKey_Value, Length - 1)
  Case 3 'Binary
   SubKey_Value = Left$(SubKey_Value, Length - 1)
   TempStr = ""
   For i = 1 To Len(SubKey_Value)
    TempStr = TempStr & Format$(Hex(Asc(Mid$(SubKey_Value, i, 1))), "00") & " "
   Next i
   SubKey_Value = TempStr
  Case Else
   SubKey_Value = "value_type=" & value_type
 End Select
 GetValue = SubKey_Value
 Exit Function
handelgetavalue:
 GetValue = ""
 Exit Function
End Function
'This property returns the current KeyValue
Public Property Get RegistryRootKey() As HKeys
 RegistryRootKey = RootHKey
End Property
'This property returns the current 'Registry Directory' your in
Public Property Get SubDirectory() As String
 SubDirectory = SubDir
End Property
'This function open's the registry at a specific 'Registry Directory'
'eg.
'  Dim MyVariant As Variant, MyReg As New CReadEasyReg, i As Integer
'  If Not MyReg.OpenRegistry(HKEY_LOCAL_MACHINE, "") Then
'   MsgBox "Couldn't open the registry"
'   Exit Sub
'  End If
'  MyVariant = MyReg.GetAllSubDirectories
'  For i = LBound(MyVariant) To UBound(MyVariant)
'   Debug.Print MyVariant(i)
'  Next i
'  MyReg.CloseRegistry
Public Function OpenRegistry(ByVal RtHKey As HKeys, ByVal SbDr As String) As Boolean
On Error GoTo OpenReg
 If RtHKey = 0 Then
  OpenRegistry = False
  OpenRegOk = False
  Exit Function
 End If
 RootHKey = RtHKey
 SubDir = SbDr
 If OpenRegOk Then
  CloseRegistry
  OpenRegOk = False
 End If
 If RegOpenKeyEx(RootHKey, SubDir, 0&, KEY_ALL_ACCESS, HKey) <> 0 Then
  OpenRegistry = False
  Exit Function
 End If
 OpenRegOk = True
 OpenRegistry = True
 Exit Function
OpenReg:
 OpenRegOk = False
 OpenRegistry = False
 Exit Function
End Function
'This function should be called after you're done with the registry
'eg. (see other examples)
Public Function CloseRegistry() As Boolean
On Error Resume Next
 If RegCloseKey(HKey) <> 0 Then
  CloseRegistry = False
  Exit Function
 End If
 CloseRegistry = True
 OpenRegOk = False
End Function
Private Sub Class_Initialize()
 RootHKey = &H0
 SubDir = ""
 HKey = 0
 OpenRegOk = False
End Sub
Private Sub Class_Terminate()
On Error Resume Next
 If RegCloseKey(HKey) <> 0 Then
  Exit Sub
 End If
End Sub
```

