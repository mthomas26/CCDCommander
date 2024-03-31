Attribute VB_Name = "File_Commands"
' Copyright (C) 2004-2024 Matthew Thomas
'
' This file is part of CCD Commander.
'
' CCD Commander is free software: you can redistribute it and/or modify it under the terms of the GNU
' General Public License as published by the Free Software Foundation, version 3 of the License.
'
' CCD Commander is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without
' even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General
' Public License for more details.
'
' You should have received a copy of the GNU General Public License along with CCD Commander.
' If not, see <https://www.gnu.org/licenses/>.
'
'---------------------------------------------------------------------------------------------------------


Option Explicit

' Windows Registry Root Key Constants.
Public Const HKEY_CLASSES_ROOT = &H80000000
 
' Windows Registry Key Type Constants.
Public Const REG_OPTION_NON_VOLATILE = 0        ' Key is preserved when system is rebooted

Public Const REG_SZ = 1                         ' Unicode nul terminated string

' Function Error Constants.
Public Const ERROR_SUCCESS = 0

' Registry Access Rights.
Public Const SYNCHRONIZE = &H100000
Public Const READ_CONTROL = &H20000
Public Const STANDARD_RIGHTS_ALL = &H1F0000
Public Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
Public Const KEY_CREATE_LINK = &H20
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Public Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))

' Windows Registry API Declarations.
' Registry API To Open A Key.
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" _
  (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, _
  ByVal samDesired As Long, phkResult As Long) As Long

' Registry API To Create A New Key.
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" _
  (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, _
  ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, _
  ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long

' Registry API To Query A String Value.
Private Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" _
  (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
  lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
  ' Note that if you declare the lpData parameter as String, you must pass it By Value.

' Registry API To Query A Long (DWORD) Value.
Private Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" _
  (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
  lpType As Long, lpData As Long, lpcbData As Long) As Long

' Registry API To Query A NULL Value.
Private Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" _
  (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
  lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long

' Registry API To Set A String Value.
Private Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" _
  (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, _
  ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
  ' Note that if you declare the lpData parameter as String, you must pass it By Value.

' Registry API To Set A Long (DWORD) Value.
Private Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" _
  (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, _
  ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long

' Registry API To Delete A Key.
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" _
  (ByVal hKey As Long, ByVal lpSubKey As String) As Long

' Registry API To Delete A Key Value.
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" _
  (ByVal hKey As Long, ByVal lpValueName As String) As Long

' Registry API To Close A Key.
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Public Function Associate_File(Extension As String, Application As String, Identifier As String, Description As String, Icon As String)

  Dim lRtn    As Long     ' Returned Value From API Registry Call
  Dim hKey    As Long     ' Handle Of Open Key
  Dim lValue  As Long     ' Setting A Long Data Value
  Dim sValue  As String   ' Setting A String Data Value
  Dim lsize   As Long     ' Size Of String Data To Set
  Dim commandline As String
  

  ' Create The New Registry Key, the file extension
  lRtn = RegCreateKeyEx(HKEY_CLASSES_ROOT, Extension, 0&, vbNullString, REG_OPTION_NON_VOLATILE, _
                          KEY_ALL_ACCESS, 0&, hKey, lRtn)
  
  ' Check For An Error.
  If lRtn <> ERROR_SUCCESS Then
    ' MsgBox CreateErr
  End If
      
  lsize = Len(Identifier)      ' Get Size Of identifier String
  ' Set "(Default)" String Value to identifier
  lRtn = RegSetValueExString(hKey, "", 0&, REG_SZ, Identifier, lsize)
  ' Check For An Error.
  If lRtn <> ERROR_SUCCESS Then
      MsgBox "Error Setting String Value!"
      RegCloseKey (hKey)
      Exit Function
  End If

  ' Close The Registry Key.
  RegCloseKey (hKey)

  ' Create The New Registry Key, the file extension identifier
  lRtn = RegCreateKeyEx(HKEY_CLASSES_ROOT, Identifier, 0&, vbNullString, REG_OPTION_NON_VOLATILE, _
                          KEY_ALL_ACCESS, 0&, hKey, lRtn)
  
  ' Check For An Error.
  If lRtn <> ERROR_SUCCESS Then
    'MsgBox CreateErr
  End If
  
    lsize = Len(Description)      ' Get Size Of file type description String
  ' Set (Default) String Value to description of the file type
  lRtn = RegSetValueExString(hKey, "", 0&, REG_SZ, Description, lsize)
  ' Check For An Error.
  If lRtn <> ERROR_SUCCESS Then
      MsgBox "Error Setting String Value!"
      RegCloseKey (hKey)
      Exit Function
  End If

  ' Close The Registry Key.
  RegCloseKey (hKey)


  ' Create The New Registry Key, the default icon key within the identifier key
  lRtn = RegCreateKeyEx(HKEY_CLASSES_ROOT, (Identifier + "\DefaultIcon"), 0&, vbNullString, REG_OPTION_NON_VOLATILE, _
                          KEY_ALL_ACCESS, 0&, hKey, lRtn)
  
  ' Check For An Error.
  If lRtn <> ERROR_SUCCESS Then
    'MsgBox CreateErr
  End If
  
    lsize = Len(Icon)      ' Get Size Of String
  ' Set (Default) String Value to the full path name of the icon that will be associated with
  '    this file type
  
  lRtn = RegSetValueExString(hKey, "", 0&, REG_SZ, Icon, lsize)
  ' Check For An Error.
  If lRtn <> ERROR_SUCCESS Then
      MsgBox "Error Setting String Value!"
      RegCloseKey (hKey)
      Exit Function
  End If

  ' Close The Registry Key.
  RegCloseKey (hKey)



Identifier = Identifier + "\shell"
  ' Create The New Registry Key, the "shell" key within the identifier key

  lRtn = RegCreateKeyEx(HKEY_CLASSES_ROOT, Identifier, 0&, vbNullString, REG_OPTION_NON_VOLATILE, _
                          KEY_ALL_ACCESS, 0&, hKey, lRtn)
  
  ' Check For An Error.
  If lRtn <> ERROR_SUCCESS Then
    'MsgBox CreateErr
  End If
  
  ' Close The Registry Key.
  RegCloseKey (hKey)


Identifier = Identifier + "\open"
  ' Create The New Registry Key, the "open" command key within the shell key

  lRtn = RegCreateKeyEx(HKEY_CLASSES_ROOT, Identifier, 0&, vbNullString, REG_OPTION_NON_VOLATILE, _
                          KEY_ALL_ACCESS, 0&, hKey, lRtn)
  
  ' Check For An Error.
  If lRtn <> ERROR_SUCCESS Then
    'MsgBox CreateErr
  End If
  
  ' Close The Registry Key.
  RegCloseKey (hKey)


Identifier = Identifier + "\command"
  ' Create The New Registry Key, the "command"  key within the "open" command key

  lRtn = RegCreateKeyEx(HKEY_CLASSES_ROOT, Identifier, 0&, vbNullString, REG_OPTION_NON_VOLATILE, _
                          KEY_ALL_ACCESS, 0&, hKey, lRtn)
  
  ' Check For An Error.
  If lRtn <> ERROR_SUCCESS Then
    'MsgBox CreateErr
  End If

    commandline = (Chr$(34) + Application + Chr$(34) + " " + Chr$(34) + "%1" + Chr$(34))
    lsize = Len(commandline)      ' Get Size Of String
  ' Set (Default) String Value of the "command" key to the command line to be used to open the file
  lRtn = RegSetValueExString(hKey, "", 0&, REG_SZ, commandline, lsize)
  ' Check For An Error.
  If lRtn <> ERROR_SUCCESS Then
      MsgBox "Error Setting String Value!"
      RegCloseKey (hKey)
      Exit Function
  End If

  ' Close The Registry Key.
  RegCloseKey (hKey)

  




End Function

Public Function CheckFileAssociation(ByVal Extension As String) As String
    'read in the program name associated with this filetype
    CheckFileAssociation = ReadKey(HKEY_CLASSES_ROOT, Extension, "", "")
End Function


Private Function ReadKey(ByVal KeyName As String, ByVal SubKeyName As String, ByVal ValueName As String, ByVal DefaultValue As String) As String
    Dim sBuffer As String
    Dim lBufferSize As Long
    Dim ret&
    Dim lphKey&
    
    sBuffer = Space(255)
    lBufferSize = Len(sBuffer)
    ret& = RegOpenKeyEx(KeyName, SubKeyName, 0, KEY_READ, lphKey&)
    If ret& = ERROR_SUCCESS Then
        ret& = RegQueryValueExString(lphKey&, ValueName, 0, REG_SZ, sBuffer, lBufferSize)
        ret& = RegCloseKey(lphKey&)
    Else
        ret& = RegCloseKey(lphKey&)
    End If
    
    sBuffer = Trim(sBuffer)
    If sBuffer <> "" Then
        sBuffer = Left(sBuffer, Len(sBuffer) - 1)
    Else
        sBuffer = DefaultValue
    End If
    
    ReadKey = sBuffer
End Function

