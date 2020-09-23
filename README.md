<div align="center">

## Get Windows Version


</div>

### Description

Finds the OS version of Windows 95/SP1/OSR2, Win 98/SP1/SE, Win ME, Win NT 3.51/4.0, Windows 2000, Windows XP, Windows CE 1.0/2.0/2.1/3.0. (Revised Version)
 
### More Info
 
Eg.

lblOS.Caption = GetWindowsVersion

O.S's specified


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Mike Whitman](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/mike-whitman.md)
**Level**          |Intermediate
**User Rating**    |4.7 (14 globes from 3 users)
**Compatibility**  |VB 6\.0
**Category**       |[Windows System Services](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-system-services__1-35.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/mike-whitman-get-windows-version__1-35762/archive/master.zip)

### API Declarations

```
Public Declare Function GetVersionExA Lib "kernel32" (lpVersionInformation As OSVERSIONINFO) As Integer
```


### Source Code

```
'IN MODULE
Public Declare Function GetVersionExA Lib "kernel32" (lpVersionInformation As OSVERSIONINFO) As Integer
Public Type OSVERSIONINFO
 dwOSVersionInfoSize As Long
 dwMajorVersion As Long
 dwMinorVersion As Long
 dwBuildNumber As Long
 dwPlatformId As Long
 szCSDVersion As String * 128
End Type
Public Function GetWindowsVersion() As String
 Dim OSInfo As OSVERSIONINFO
 Dim Ret As Integer
 OSInfo.dwOSVersionInfoSize = 148
 OSInfo.szCSDVersion = Space$(128)
 Ret = GetVersionExA(OSInfo)
With OSInfo
 Select Case .dwPlatformId
  Case 1
   If .dwMinorVersion < 10 Then
    If .dwBuildNumber = 950 Then
     GetWindowsVersion = "Windows 95"
    ElseIf .dwBuildNumber > 950 Or .dwBuildNumber <= 1080 Then
     GetWindowsVersion = "Windows 95 SP1"
    Else
     GetWindowsVersion = "Windows 95 OSR2"
    End If
   ElseIf .dwMinorVersion = 10 Then
    If .dwBuildNumber = 1998 Then
     GetWindowsVersion = "Windows 98"
    ElseIf .dwBuildNumber > 1998 Or .dwBuildNumber < 2183 Then
     GetWindowsVersion = "Windows 98 SP1"
    ElseIf .dwBuildNumber >= 2183 Then
     GetWindowsVersion = "Windows 98 SE"
    End If
   Else
    GetWindowsVersion = "Windows ME"
   End If
  Case 2
   If .dwMajorVersion = 3 Then
    GetWindowsVersion = "Windows NT 3.51"
   ElseIf .dwMajorVersion = 4 Then
    GetWindowsVersion = "Windows NT 4.0"
   ElseIf .dwMajorVersion = 5 Then
    If .dwMinorVersion = 0 Then
     GetWindowsVersion = "Windows 2000"
    Else
     GetWindowsVersion = "Windows XP"
    End If
   End If
  Case 3
   If .dwMajorVersion = 1 Then
    GetWindowsVersion = "Windows CE 1.0"
   ElseIf .dwMajorVersion = 2 Then
    If .dwMinorVersion = 0 Then
     GetWindowsVersion = "Windows CE 2.0"
    Else
     GetWindowsVersion = "Windows CE 2.1"
    End If
   Else
    GetWindowsVersion = "Windows CE 3.0"
   End If
  Case Else
   GetWindowsVersion = "Unable to get Windows Version"
 End Select
End With
End Function
```

