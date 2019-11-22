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
 path As String
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
    ' https://stackoverflow.com/a/4875294/12258312
    ' https://stackoverflow.com/a/4876841/12258312
    Size As Byte
    
    Cause As String
    Desc As String * 512
End Type

Function GetEXEWordSize_ToString(dat As GetEXEWordSize_out) As String
    Dim buf As String
    
    ' format is kinda like zfill,
    ' https://bytes.com/topic/visual-basic/answers/778694-how-format-number-0000-a
    
    buf = nzfill(dat.Size, 3)
    buf = buf + IIf(Asc(dat.Cause) <> 0, ":" + dat.Cause, "")
    
    ' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/ltrim-rtrim-and-trim-functions
    buf = buf + IIf(Asc(dat.Desc) <> 0, "-> " + Trim(dat.Desc), "")
    
    GetEXEWordSize_ToString = buf
End Function

Function nzfill(i As Byte, n As Byte) As String
 Dim buf As String
 buf = String(n, "0")
 nzfill = Format(i, buf)
End Function

Function str2hexarray(s As String, Optional delim As String = " ") As String
 Dim i As Integer
 Dim r As String
 
 For i = 1 To Len(s)
  r = r + Hex(Asc(Mid(s, i, 1))) + delim
 Next
 str2hexarray = r
End Function

Function read_binary_file(path As String, Optional l As Integer = 2) As Byte()
    Dim nFile As Integer
    Debug.Assert (l > 0)
    
    nFile = FreeFile
    
    Open path For Binary Access Read As nFile Len = l
    If LOF(nFile) > 0 Then
        read_binary_file = Input(LOF(nFile), nFile)
        'ReDim read_binary_file(0 To LOF(nFile) - 1)
        'Get nFile, , read_binary_file
    End If
    Close nFile
End Function

Function GetEXEWordSize(AppPath As String, Optional maxRdLen As Integer = 8192) As GetEXEWordSize_out
 ' +8192 = 2000h = 2*(observed emphirical PE header start pos)
 'Try gathering info thru ShGetFileInfo first
 Dim SHFI As SHFILEINFO
 Dim lngResult   As Long
 Dim intLoWord   As Integer
 Dim intLoWordHiByte As Integer
 Dim intLoWordLoByte As Integer
 Dim strLOWORD   As String
 
 ' https://superuser.com/questions/358434/how-to-check-if-a-binary-is-32-or-64-bit-on-windows)
 ' answer in reverse endianness format though
 ' Hence (in HEX):
 ' 32-bit: 4C 01 -> 076 001 DEC
 ' 64-bit: 64 86 -> 100 134 DEC
 Const PE_HEADER As String = "PE" + vbNullChar + vbNullChar 'PE\0\0

 Dim SIGN32 As String * 2
 SIGN32 = Chr(76) + Chr(1)

 Dim SIGN64 As String * 2
 SIGN64 = Chr(100) + Chr(134)
 
 Dim ret As GetEXEWordSize_out
 
 lngResult = SHGetFileInfo(AppPath, 0, SHFI, Len(SHFI), &H2000)
  
 If lngResult = 0 Then ' If EXE cannot be read
  ret.Size = 0
  ret.Cause = ret.Cause + "lng-|"
  
  Dim pe_buf As String
  pe_buf = read_binary_file(AppPath, maxRdLen)
  
 
  'Dim iFileNo As Integer
  'iFileNo = FreeFile
  'Open "C:\Test.txt" For Output As #iFileNo
  'Print #iFileNo, str2hexarray(pe_buf)
  'Form1.Text2.Text = str2hexarray(pe_buf)
  'Close #iFileNo
  
  ' https://superuser.com/questions/358434/how-to-check-if-a-binary-is-32-or-64-bit-on-windows)
  Dim pe_pos As Long
  pe_pos = InStr(1, pe_buf, PE_HEADER, vbBinaryCompare)

  If (pe_pos > 0) Then
   Dim pe_nextbytes As String

   pe_nextbytes = Mid(pe_buf, pe_pos + Len(PE_HEADER), 2)
   If (Len(pe_nextbytes)) Then
    If (StrComp(pe_nextbytes, SIGN32, vbBinaryCompare) = 0) Then
     ret.Size = 32 '
     ret.Cause = ret.Cause + "Sign32|"
    ElseIf (StrComp(pe_nextbytes, SIGN64, vbBinaryCompare) = 0) Then
     ret.Size = 64 '
     ret.Cause = ret.Cause + "Sign64|"
    Else
     ret.Size = 0 '
     ret.Cause = ret.Cause + "Sign?? (" + str2hexarray(Mid(pe_buf, pe_pos, 10)) + ") @ " + Hex(pe_pos) + "|"
    End If
   End If
  Else
   ret.Size = 0 ' prefill
   ret.Cause = ret.Cause + "NonPE/NonExecutable?|"
  End If ' If (pe_pos > 0) Then
 Else ' if can be read
  ret.Cause = ret.Cause + "lng+|"

  
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
       ElseIf .dwVerMajor >= 4 Then ' If PE app and Win NT or higher
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
        ' WinNT is designed for 32 bits
        ret.Cause = ret.Cause + "PE&WinNT3.51|"
        ret.Size = 32
       End If
      End With
    End Select
 End If
 GetEXEWordSize = ret
End Function

