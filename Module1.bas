Attribute VB_Name = "Module1"
Option Explicit

' Used by: GetEXEWordSize
Declare Function GetBinaryType Lib "kernel32" Alias "GetBinaryTypeA" (ByVal lpApplicationName As String, lpBinaryType As Long) As Long

' Used by: GetEXEWordSize
Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Integer

' Used by: GetEXEWordSize
Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
 
' Deprecated, but let it be for now
Type EXE
 Caption As String
 Handle As Long
 hWnd As Long
 Module As Long
 nSize As Byte
 Path As String
 PID As Long
End Type
 
' for OS info
Type OSVERSIONINFO
 OSVSize         As Long
 dwVerMajor      As Long
 dwVerMinor      As Long
 dwBuildNumber   As Long
 PlatformID      As Long
 szCSDVersion    As String * 128
End Type
 
Type SHFILEINFO
 hIcon As Long ' out: icon
 iIcon As Long ' out: icon index
 dwAttributes As Long ' out: SFGAO_ flags
 szDisplayName As String * 260 ' out: display name (or path)
 szTypeName As String * 80 ' out: type name
End Type

' custom struct, for data output
Type GetEXEWordSize_out
    Size As Byte
    Cause As String
    Desc As String
End Type

Function GetEXEWordSize_ToString(dat As GetEXEWordSize_out) As String
    ' format is kinda like zfill,
    ' https://bytes.com/topic/visual-basic/answers/778694-how-format-number-0000-a
    GetEXEWordSize_ToString = Format(dat.Size, "000") + ":" + dat.Cause
End Function

Function GetEXEWordSize(AppPath As String) As GetEXEWordSize_out
 'Try gathering info thru ShGetFileInfo first
 Dim SHFI As SHFILEINFO
 Dim lngResult   As Long
 Dim intLoWord   As Integer
 Dim intLoWordHiByte As Integer
 Dim intLoWordLoByte As Integer
 Dim strLOWORD   As String
 
 Dim ret As GetEXEWordSize_out
 
 lngResult = SHGetFileInfo(AppPath, 0, SHFI, Len(SHFI), &H2000)
  
 If lngResult = 0 Then ' If EXE cannot be read
  ret.Size = 0 ' TODO
  ret.Cause = ret.Cause + "lng-|"
  GetEXEWordSize = ret
  Exit Function
 Else
  ret.Cause = ret.Cause + "lng+|"
 End If
  
 intLoWord = lngResult And &HFFFF&
 intLoWordHiByte = intLoWord \ &H100 And &HFF&
 intLoWordLoByte = intLoWord And &HFF&
 strLOWORD = Chr$(intLoWordLoByte) & Chr$(intLoWordHiByte)
  
 Select Case strLOWORD
  Case "NE", "MZ" ' as far as I can tell,  both NE and MZ are 16bit
   ret.Size = 16
   ret.Cause = ret.Cause + "LOWORD:NEMZ|"
  Case "PE" ' If PE app, gather OS info
  ret.Cause = ret.Cause + "LOWORD:PE|"
   Dim OSV As OSVERSIONINFO
   With OSV
    .OSVSize = Len(OSV)
    GetVersionEx OSV
    If .PlatformID < 2 Then ' If PE app and Win 9x
     ret.Size = 32
     ret.Cause = ret.Cause + "PE&Win9x|"
    Else ' If PE app and Win NT or higher
     If .dwVerMajor >= 4 Then
       ret.Cause = ret.Cause + "PE&WinNT4|"
       ' Get info via GetBinaryType
       Dim BinaryType As Long
       GetBinaryType AppPath, BinaryType
       Select Case BinaryType
        Case 0 'SCS_32BIT_BINARY
         ' https://docs.microsoft.com/en-us/windows/win32/api/winbase/nf-winbase-getbinarytypew
         ret.Desc = " A 32-bit Windows-based application "
         ret.Cause = ret.Cause + "SCS_32BIT_BINARY|"
         ret.Size = 32
        Case 1 'SCS_DOS_BINARY
         ' https://users.cs.jmu.edu/abzugcx/Public/Student-Produced-Term-Projects/Operating-Systems-2003-FALL/MS-DOS-by-Dominic-Swayne-Fall-2003.pdf
         ' First known as 86-DOS, it was developed in about 6 weeks by Tim Paterson of Seattle Computer Products (SCP).  The OS was designed to operate on the company’s own 16-bit personal computers running the Intel 8086 microprocessor.  (Paterson, 1983a)
         ret.Cause = ret.Cause + "SCS_DOS_BINARY|"
         ret.Size = 16
        Case 2 'SCS_WOW_BINARY
         ' https://docs.microsoft.com/en-us/windows/win32/api/winbase/nf-winbase-getbinarytypew
         ret.Desc = "A 16-bit Windows-based application"
         ret.Cause = ret.Cause + "SCS_WOW_BINARY|"
         ret.Size = 16
        Case 3 'SCS_PIF_BINARY
         ' https://users.cs.jmu.edu/abzugcx/Public/Student-Produced-Term-Projects/Operating-Systems-2003-FALL/MS-DOS-by-Dominic-Swayne-Fall-2003.pdf
         ' First known as 86-DOS, it was developed in about 6 weeks by Tim Paterson of Seattle Computer Products (SCP).  The OS was designed to operate on the company’s own 16-bit personal computers running the Intel 8086 microprocessor.  (Paterson, 1983a)
         ret.Desc = " A PIF file that executes an MS-DOS – based application "
         ret.Cause = ret.Cause + "SCS_PIF_BINARY|"
         ret.Size = 16
        Case 4 'SCS_POSIX_BINARY
         ' https://en.wikipedia.org/wiki/Program_information_file
         ' ...
         ' https://stackoverflow.com/q/58986468
         ret.Cause = ret.Cause + "SCS_POSIX_BINARY|"
         ret.Size = 16 ' Posix word size unknown
        Case 5 'SCS_OS216_BINARY
         ' https://docs.microsoft.com/en-us/windows/win32/api/winbase/nf-winbase-getbinarytypew
         ret.Desc = " A 16-bit OS/2-based application "
         ret.Cause = ret.Cause + "SCS_OS216_BINARY|"
         ret.Size = 16
        Case 6 'SCS_64BIT_BINARY
         ' https://docs.microsoft.com/en-us/windows/win32/api/winbase/nf-winbase-getbinarytypew
         ret.Desc = " A 64-bit Windows-based application. "
         ret.Cause = ret.Cause + "SCS_64BIT_BINARY|"
         ret.Size = 64
       End Select
     Else ' However, if we have, say, Windows NT 3.51, then
      ret.Cause = ret.Cause + "PE&WinNT3.51|"
      ret.Size = 32
     End If
     
     GetEXEWordSize = ret
    End If
   End With
 End Select
 
End Function
