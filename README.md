<div align="center">

## A Network Drive Mapping Module


</div>

### Description

Module used to map network drives to next available drive letter and to disconnect network drives.

Very Simple to use.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Bryan Lass](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/bryan-lass.md)
**Level**          |Intermediate
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/bryan-lass-a-network-drive-mapping-module__1-53008/archive/master.zip)





### Source Code

```
Option Explicit
Private Const CONNECT_UPDATE_PROFILE = &H1
Private Const RESOURCE_CONNECTED As Long = &H1&
 Public iDrive As Integer
 Public iFirst As Integer
 Public iFirstFree As Integer, sFirstFree As String
 Public sNextDrive As String
Public Declare Function GetDriveType Lib "kernel32" Alias _
 "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Const RESOURCE_GLOBALNET As Long = &H2&
Private Const RESOURCETYPE_DISK As Long = &H1&
Private Const RESOURCEDISPLAYTYPE_SHARE& = &H3
Private Const RESOURCEUSAGE_CONNECTABLE As Long = &H1&
Private Declare Function WNetAddConnection2 Lib "mpr.dll" _
 Alias "WNetAddConnection2A" (lpNetResource As NETCONNECT, _
 ByVal lpPassword As String, ByVal lpUserName As String, _
 ByVal dwFlags As Long) As Long
Private Declare Function WNetCancelConnection2 Lib "mpr.dll" _
 Alias "WNetCancelConnection2A" (ByVal lpName As String, _
 ByVal dwFlags As Long, ByVal fForce As Long) As Long
Private Type NETCONNECT
 dwScope As Long
 dwType As Long
 dwDisplayType As Long
 dwUsage As Long
 lpLocalName As String
 lpRemoteName As String
 lpComment As String
 lpProvider As String
End Type
Public Function MapDrive(LocalDrive As String, _
 RemoteDrive As String, Optional Username As String, _
 Optional Password As String) As Boolean
 Dim NetR As NETCONNECT
 NetR.dwScope = RESOURCE_GLOBALNET
 NetR.dwType = RESOURCETYPE_DISK
 NetR.dwDisplayType = RESOURCEDISPLAYTYPE_SHARE
 NetR.dwUsage = RESOURCEUSAGE_CONNECTABLE
 NetR.lpLocalName = Left$(LocalDrive, 1) & ":"
 NetR.lpRemoteName = RemoteDrive
 MapDrive = (WNetAddConnection2(NetR, Username, Password, _
 CONNECT_UPDATE_PROFILE) = 0)
End Function
Public Function DisconnectDrive(LocalDrive As String) As String
 DisconnectDrive = WNetCancelConnection2(Left$(LocalDrive, 1) & ":", _
 CONNECT_UPDATE_PROFILE, False) = 0
End Function
Public Function FindDrive() As String
 iDrive = 67
 Do
 iDrive = iDrive + 1
 sNextDrive = Chr$(iDrive) + ":\"
 iFirstFree = GetDriveType(sNextDrive)
 Loop Until iFirstFree = 1
 sFirstFree = Chr$(iDrive) + ":\"
 FindDrive = sFirstFree
End Function
'Syntax is as follows
Private sub NetConnect()
Dim UncPath As String
UncPath="\\server\folder\subfolder\subfolder\destinationfolder"
MapDrive FindDrive, UncPath
end sub
Private sub DropDrive()
Dim DrLetter as string
DrLetter= "e"' Any Letter you want
disconnectdrive drletter
end sub
```

