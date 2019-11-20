Attribute VB_Name = "Module1"
Option Explicit

' Used by: GetEXEWordSize
Declare Function GetBinaryType Lib "kernel32" Alias "GetBinaryTypeA" (ByVal lpApplicationName As String, lpBinaryType As Long) As Long

' Used by: GetEXEWordSize
Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Integer

' Used by: GetEXEWordSize
Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
 
Type EXE
 Caption As String
 Handle As Long
 hWnd As Long
 Module As Long
 nSize As Byte
 Path As String
 PID As Long
End Type
 
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

Function GetEXEWordSize(AppPath As String) As Byte
 'Try gathering info thru ShGetFileInfo first
 Dim SHFI As SHFILEINFO
 Dim lngResult   As Long
 Dim intLoWord   As Integer
 Dim intLoWordHiByte As Integer
 Dim intLoWordLoByte As Integer
 Dim strLOWORD   As String
 
 lngResult = SHGetFileInfo(AppPath, 0, SHFI, Len(SHFI), &H2000)
  
 If lngResult = 0 Then ' If EXE cannot be read
  GetEXEWordSize = 0 ' TODO
  Exit Function
 End If
  
 intLoWord = lngResult And &HFFFF&
 intLoWordHiByte = intLoWord \ &H100 And &HFF&
 intLoWordLoByte = intLoWord And &HFF&
 strLOWORD = Chr$(intLoWordLoByte) & Chr$(intLoWordHiByte)
  
 Select Case strLOWORD
  Case "NE", "MZ" ' as far as I can tell,  both NE and MZ are 16bit
   GetEXEWordSize = 2
  Case "PE" ' If PE app, gather OS info
   Dim OSV As OSVERSIONINFO
   With OSV
    .OSVSize = Len(OSV)
    GetVersionEx OSV
    If .PlatformID < 2 Then ' If PE app and Win 9x
     GetEXEWordSize = 4
    Else ' If PE app and Win NT or higher
     If .dwVerMajor >= 4 Then
       ' Get info via GetBinaryType
       Dim BinaryType As Long
       GetBinaryType AppPath, BinaryType
       Select Case BinaryType
        Case 0 'SCS_32BIT_BINARY
         GetEXEWordSize = 4
        Case 1 'SCS_DOS_BINARY
         GetEXEWordSize = 2
        Case 2 'SCS_WOW_BINARY
         GetEXEWordSize = 2
        Case 3 'SCS_PIF_BINARY
         GetEXEWordSize = 2
        Case 4 'SCS_POSIX_BINARY
         GetEXEWordSize = 0 ' Posix word size unknown
        Case 5 'SCS_OS216_BINARY
         GetEXEWordSize = 2
        Case 6 'SCS_64BIT_BINARY
         GetEXEWordSize = 8
       End Select
     Else ' However, if we have, say, Windows NT 3.51, then
      GetEXEWordSize = 4
     End If
     
    End If
   End With
 End Select
End Function
