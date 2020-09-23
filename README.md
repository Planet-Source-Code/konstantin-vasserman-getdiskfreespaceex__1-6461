<div align="center">

## GetDiskFreeSpaceEx


</div>

### Description

This code shows how to use GetFreeDiskSpaceEx API. It works with the drives larger than 2GB as oppose to the old GetFreeDiskSpace API call. It will also work with Windows2000 per-user space quota, so the free disk space you get is actually what available to the user and not all the space available on the disk.
 
### More Info
 
Pass any valid path to the GetFreeSpace function. The path could be a local drive ("c:" or "c:\windows"), network drive ("x:" or "x:\MyFolder") or UNC path like "\\myserver\myshare".

Under Windows 2000: if per-user quotas are in use, the value returned by GetFreeSpace function may be less than the total number of free bytes on the disk.

This code can be modified to get total number of free bytes on the disk (without regard to the user quota) or total number of bytes on the disk.

GetFreeSpace function returns total number of free bytes on the disk that are available to the user associated with the calling thread. Return value is a Double.

This code will not work on versions of Windows 95 prior to OSR2.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Konstantin Vasserman](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/konstantin-vasserman.md)
**Level**          |Intermediate
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/konstantin-vasserman-getdiskfreespaceex__1-6461/archive/master.zip)

### API Declarations

```
Public Type ULong ' Unsigned Long
 Byte1 As Byte
 Byte2 As Byte
 Byte3 As Byte
 Byte4 As Byte
End Type
Public Type LargeInt ' Large Integer
 LoDWord As ULong
 HiDWord As ULong
 LoDWord2 As ULong
 HiDWord2 As ULong
End Type
Public Declare Function GetDiskFreeSpaceEx Lib "kernel32" Alias "GetDiskFreeSpaceExA" _
 (ByVal lpRootPathName As String, FreeBytesAvailableToCaller As LargeInt, _
 TotalNumberOfBytes As LargeInt, TotalNumberOfFreeBytes As LargeInt) As Long
```


### Source Code

```
Function GetFreeSpace(strPath as String) As Double
 Dim nFreeBytesToCaller As LargeInt
 Dim nTotalBytes As LargeInt
 Dim nTotalFreeBytes As LargeInt
 strPath = Trim(strPath)
 If Right(strPath, 1) <> "\" Then
  strPath = strPath & "\"
 End If
 If GetDiskFreeSpaceEx(strPath, nFreeBytesToCaller, nTotalBytes, nTotalFreeBytes) <> 0 Then
  GetFreeSpace = CULong( _
   nFreeBytesToCaller.HiDWord.Byte1, _
   nFreeBytesToCaller.HiDWord.Byte2, _
   nFreeBytesToCaller.HiDWord.Byte3, _
   nFreeBytesToCaller.HiDWord.Byte4) * 2 ^ 32 + _
   CULong(nFreeBytesToCaller.LoDWord.Byte1, _
   nFreeBytesToCaller.LoDWord.Byte2, _
   nFreeBytesToCaller.LoDWord.Byte3, _
   nFreeBytesToCaller.LoDWord.Byte4)
 End If
End Function
Function CULong(Byte1 As Byte, Byte2 As Byte, Byte3 As Byte, Byte4 As Byte) As Double
 CULong = Byte4 * 2 ^ 24 + Byte3 * 2 ^ 16 + Byte2 * 2 ^ 8 + Byte1
End Function
```

