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
 input As String
 time As String
 
 ' https://stackoverflow.com/a/4875294/12258312
 ' https://stackoverflow.com/a/4876841/12258312
 wordsize As Byte
    
 walkthrough As String
 
 desc As String * 80
End Type

Function GetEXEWordSize_ToJson(dat As GetEXEWordSize_out) As String
 ' Its sure rudimental,
 ' but it works!
 
 Dim C34 As String * 1
 C34 = Chr(34)
 
 Dim C34P As String
 C34P = C34 + ": " + C34
 
 Dim buf As String
 With dat
  .walkthrough = .walkthrough + "walk!"
 
  buf = "{" + vbCrLf
  buf = buf + String(1, vbTab) + C34 + App.Title + C34 + ":{" + vbCrLf
  
  ' input + time
  buf = buf + String(2, vbTab) + C34 + "input" + C34P + .input + C34 + vbCrLf
  buf = buf + String(2, vbTab) + C34 + "time" + C34P + .time + C34 + vbCrLf
  
  ' wordsize
  buf = buf + String(2, vbTab) + C34 + "wordsize" + C34P
  buf = buf + zfill_byte(.wordsize, 3) + C34 + vbCrLf
  
  ' desc
  buf = buf + String(2, vbTab) + C34 + "desc" + C34P
  ' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/ltrim-rtrim-and-trim-functions
  buf = buf + IIf(Asc(.desc), Trim(.desc), "") + C34 + vbCrLf
     
  ' walkthrough
  buf = buf + String(2, vbTab) + C34 + "walkthrough" + C34P
  buf = buf + IIf(Asc(.walkthrough), .walkthrough, "") + C34 + vbCrLf
  
  ' end item
  buf = buf + String(1, vbTab) + "}" + vbCrLf
  
  ' end json
  buf = buf + "}" + vbCrLf
  
  GetEXEWordSize_ToJson = buf
 End With
End Function

Function zfill_byte(i As Byte, n As Byte) As String
 ' format is kinda like zfill,
 ' https://bytes.com/topic/visual-basic/answers/778694-how-format-number-0000-a
 Dim buf As String
 buf = String(n, "0")
 zfill_byte = Format(i, buf)
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
        read_binary_file = input(LOF(nFile), nFile)
        'ReDim read_binary_file(0 To LOF(nFile) - 1)
        'Get nFile, , read_binary_file
    End If
    Close nFile
End Function

Function GetEXEWordSize_prefill(s As GetEXEWordSize_out, AppPath As String)
 With s
  .walkthrough = "rdy|"
  .input = AppPath
  .time = Now
 End With
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
 GetEXEWordSize_prefill ret, AppPath
 
 lngResult = SHGetFileInfo(AppPath, 0, SHFI, Len(SHFI), &H2000)
  
 If lngResult = 0 Then ' If EXE cannot be read
  ret.wordsize = 0
  ret.walkthrough = ret.walkthrough + "SHGetFileInfo=BAD|"
  
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
     ret.wordsize = 32 '
     ret.walkthrough = ret.walkthrough + "Sign32|"
    ElseIf (StrComp(pe_nextbytes, SIGN64, vbBinaryCompare) = 0) Then
     ret.wordsize = 64 '
     ret.walkthrough = ret.walkthrough + "Sign64|"
    Else
     ret.wordsize = 0 '
     ret.walkthrough = ret.walkthrough + "Sign?? (" + str2hexarray(Mid(pe_buf, pe_pos, 10)) + ") @ " + Hex(pe_pos) + "|"
    End If
   End If
  Else
   ret.wordsize = 0 ' prefill
   ret.walkthrough = ret.walkthrough + "NonPE/NonExecutable?|"
  End If ' If (pe_pos > 0) ...
 Else ' if can be read
  ret.walkthrough = ret.walkthrough + "SHGetFileInfo=OK|"

  
    intLoWord = lngResult And &HFFFF&
    intLoWordHiByte = intLoWord \ &H100 And &HFF&
    intLoWordLoByte = intLoWord And &HFF&
    strLOWORD = Chr$(intLoWordLoByte) & Chr$(intLoWordHiByte)
     
    Select Case strLOWORD
     Case "NE", "MZ" ' as far as I can tell,  both NE and MZ are 16bit
      ret.wordsize = 16
      ret.walkthrough = ret.walkthrough + "LOWORD:NEMZ|"
     Case "PE" ' If PE app, gather OS info
      ret.walkthrough = ret.walkthrough + "LOWORD:PE|"
      Dim OSV As OSVERSIONINFO
      With OSV
       .OSVSize = Len(OSV)
       GetVersionEx OSV
       If .PlatformID < 2 Then ' If PE app and Win 9x
        ret.wordsize = 32
        ret.walkthrough = ret.walkthrough + "PE&Win9x|"
       ElseIf .dwVerMajor >= 4 Then ' If PE app and Win NT or higher
         ret.walkthrough = ret.walkthrough + "PE&WinNT4|"
         ' Get info via GetBinaryType
         Dim BinaryType As Long
         GetBinaryType AppPath, BinaryType
         Select Case BinaryType
          Case 0 'SCS_32BIT_BINARY
           ' https://docs.microsoft.com/en-us/windows/win32/api/winbase/nf-winbase-getbinarytypew
           ret.desc = " A 32-bit Windows-based application "
           ret.walkthrough = ret.walkthrough + "SCS_32BIT_BINARY|"
           ret.wordsize = 32
          Case 1 'SCS_DOS_BINARY
           ' https://users.cs.jmu.edu/abzugcx/Public/Student-Produced-Term-Projects/Operating-Systems-2003-FALL/MS-DOS-by-Dominic-Swayne-Fall-2003.pdf
           ' First known as 86-DOS, it was developed in about 6 weeks by Tim Paterson of Seattle Computer Products (SCP).  The OS was designed to operate on the company’s own 16-bit personal computers running the Intel 8086 microprocessor.  (Paterson, 1983a)
           ret.walkthrough = ret.walkthrough + "SCS_DOS_BINARY|"
           ret.wordsize = 16
          Case 2 'SCS_WOW_BINARY
           ' https://docs.microsoft.com/en-us/windows/win32/api/winbase/nf-winbase-getbinarytypew
           ret.desc = "A 16-bit Windows-based application"
           ret.walkthrough = ret.walkthrough + "SCS_WOW_BINARY|"
           ret.wordsize = 16
          Case 3 'SCS_PIF_BINARY
           ' https://users.cs.jmu.edu/abzugcx/Public/Student-Produced-Term-Projects/Operating-Systems-2003-FALL/MS-DOS-by-Dominic-Swayne-Fall-2003.pdf
           ' First known as 86-DOS, it was developed in about 6 weeks by Tim Paterson of Seattle Computer Products (SCP).  The OS was designed to operate on the company’s own 16-bit personal computers running the Intel 8086 microprocessor.  (Paterson, 1983a)
           ret.desc = " A PIF file that executes an MS-DOS – based application "
           ret.walkthrough = ret.walkthrough + "SCS_PIF_BINARY|"
           ret.wordsize = 16
          Case 4 'SCS_POSIX_BINARY
           ' https://en.wikipedia.org/wiki/Program_information_file
           ' ...
           ' https://stackoverflow.com/q/58986468
           ret.walkthrough = ret.walkthrough + "SCS_POSIX_BINARY|"
           ret.wordsize = 16 ' Posix word wordsize unknown
          Case 5 'SCS_OS216_BINARY
           ' https://docs.microsoft.com/en-us/windows/win32/api/winbase/nf-winbase-getbinarytypew
           ret.desc = " A 16-bit OS/2-based application "
           ret.walkthrough = ret.walkthrough + "SCS_OS216_BINARY|"
           ret.wordsize = 16
          Case 6 'SCS_64BIT_BINARY
           ' https://docs.microsoft.com/en-us/windows/win32/api/winbase/nf-winbase-getbinarytypew
           ret.desc = " A 64-bit Windows-based application. "
           ret.walkthrough = ret.walkthrough + "SCS_64BIT_BINARY|"
           ret.wordsize = 64
         End Select
       Else ' However, if we have, say, Windows NT 3.51, then
        ' WinNT is designed for 32 bits
        ret.walkthrough = ret.walkthrough + "PE&WinNT3.51|"
        ret.wordsize = 32
       End If
      End With
    End Select
 End If
 GetEXEWordSize = ret
End Function

