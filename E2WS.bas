Attribute VB_Name = "E2WS"
Option Explicit

' Used by: get_wordsize_from_info
Declare Function GetBinaryType Lib "kernel32" Alias "GetBinaryTypeA" (ByVal lpApplicationName As String, lpBinaryType As Long) As Long

' Used by: get_wordsize_from_info
Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Integer

' Used by: get_wordsize_from_info
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
Type wordsize_struct
 file As String
 args As String
 time As String
 code As Long
 
 ' https://stackoverflow.com/a/4875294/12258312
 ' https://stackoverflow.com/a/4876841/12258312
 wordsize As Byte
    
 walkthrough As String
 
 desc As String * 80
End Type

Type wordsize_params
 ' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/open-statement
 ' > Optional. Number less than or equal to 32,767 (bytes)
 '32,767 = 0x7FFF = max signed 2 bytes var val -> VB6 integer
 max_read_bytes As Integer
 mode As Long
 code As Long
End Type

'' In generic use
'' https://social.msdn.microsoft.com/Forums/sqlserver/en-US/d6e76731-8e3b-465f-9d5a-12c6498d6b6c/how-to-return-exit-code-from-vb6-form?forum=winforms
Private Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)

' for header detection
Const PE_HEADER As String = "PE" + vbNullChar + vbNullChar 'PE\0\0
Const JSON_PARAMS_DELIM As String = "," & vbCrLf

' In use by get_error_desc
Private Const ERROR_IRRECOVERABLE = -1
Private Const ERROR_SUCCESS = 0 ' all good
Private Const ERROR_INVALID_ARGS = 1 'problem with ARGS
Private Const ERROR_INVALID_MODE = 2 ' invalid mode
Private Const ERROR_UNKNOWN_PE_HEADER = 3
Private Const ERROR_WARNING_AMBIGUOUS_WORDSIZE = 7 ' as 777
Private Const ERROR_WARNING_BAD_LUCK = 13 ' https://en.wikipedia.org/wiki/13_(number)#Unlucky_13

' For parse_args
Private Const ARGS_FLAG_MODE As String = "M"
Private Const ARGS_FLAG_MAXRDLEN As String = "R" ' on mode = 2

' pseudo consts, see setup()
Public APP_NAME As String
Public VER As String
Public DEBUGGER As Boolean
Public C34 As String
Public SIGN32 As String
Public SIGN64 As String

Public Function setup()
 APP_NAME = "exe2wordsize"
 VER = App.Major & "." & App.Minor & App.Revision
 DEBUGGER = GetRunningInIDE()
 C34 = Chr(34)
 
 ' https://superuser.com/questions/358434/how-to-check-if-a-binary-is-32-or-64-bit-on-windows)
 ' answer in reverse endianness format though
 ' Hence (in HEX):
 ' 32-bit: 4C 01 -> 076 001 DEC
 ' 64-bit: 64 86 -> 100 134 DEC
 SIGN32 = Chr(76) + Chr(1)
 SIGN64 = Chr(100) + Chr(134)
End Function

Public Function wordsize_to_json(s As String) As String
 Dim dat As wordsize_struct
 dat = get_wordsize_from_info(s)
 wordsize_to_json = struct_to_json(dat)
End Function

Function struct_to_json(dat As wordsize_struct) As String
 ' Its sure rudimental,
 ' but it works!
 Dim buf As String
 
 
 With dat
  .walkthrough = .walkthrough & "Json|" ' for logging
 
  buf = "{" & vbCrLf
  buf = buf & String(1, vbTab) & C34 & App.Title & C34 & ":{" & vbCrLf
  
  ' file
  buf = buf & String(2, vbTab) & C34 & "file" & C34 & ": " & C34 & .file & C34 & JSON_PARAMS_DELIM
  
  ' args
  buf = buf & String(2, vbTab) & C34 & "args" & C34 & ": " & C34 & .args & C34 & JSON_PARAMS_DELIM
  
  ' time
  buf = buf & String(2, vbTab) & C34 & "time" & C34 & ": " & C34 & .time & C34 & JSON_PARAMS_DELIM
  
  ' code
  buf = buf & String(2, vbTab) & C34 & "code" & C34 & ": " & Str(.code) & JSON_PARAMS_DELIM
  
  ' code - desc
  buf = buf & String(2, vbTab) & C34 & "code_desc" & C34 & ": " & C34 & get_error_desc(.code) & C34 & JSON_PARAMS_DELIM
  
  ' wordsize
  buf = buf & String(2, vbTab) & C34 & "wordsize" & C34 & ": " & zfill_byte(.wordsize, 3) & JSON_PARAMS_DELIM
  
  ' desc
  buf = buf & String(2, vbTab) & C34 & "desc" & C34 & ": " & C34
  ' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/ltrim-rtrim-and-trim-functions
  buf = buf & IIf(Asc(.desc), Trim(.desc), "") & C34 & JSON_PARAMS_DELIM
     
  ' walkthrough
  buf = buf & String(2, vbTab) & C34 & "walkthrough" & C34 & ": " & C34
  buf = buf & IIf(Asc(.walkthrough), .walkthrough, "") & C34 & vbCrLf
  
  ' end item
  buf = buf & String(1, vbTab) & "}" & vbCrLf
  
  ' end json
  buf = buf & "}" & vbCrLf
  
  struct_to_json = buf
 End With
End Function

Function app_path()
  ' https://stackoverflow.com/a/12423852/12258312
  app_path = App.path & IIf(Right$(App.path, 1) <> "\", "\", "")
End Function

Function app_path_exe()
  app_path_exe = app_path() & App.EXEName & ".exe"
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
        read_binary_file = Input(LOF(nFile), nFile)
        'ReDim read_binary_file(0 To LOF(nFile) - 1)
        'Get nFile, , read_binary_file
    End If
    Close nFile
End Function

Private Function struct_prefill(s As wordsize_struct, AppPath As String, args As String)
 With s
  .walkthrough = "rdy|"
  .code = -1
  .file = AppPath
  .time = get_unix_time_mod
  .args = args
 End With
End Function

Function wordsize_params_init(ret As wordsize_params)
 With ret
  .max_read_bytes = 8192
  .mode = 0
  .code = 0
 End With

End Function

Function parse_args(s As String, Optional demiliter As String = " ", Optional sub_delimiter As String = "=") As wordsize_params
 On Error Resume Next
 
 Dim ret As wordsize_params
 wordsize_params_init ret

 If (Len(s) > 0) Then
  Dim a() As String
  Dim ac As Integer
  Dim i As Integer
  
  Dim b() As String
  Dim bc As Integer
  Dim bbuf As String
  
  Dim k As Integer
  
  
  a = Split(s, demiliter)
  ac = UBound(a)
  
  For i = 0 To ac
   b = Split(a(i), sub_delimiter)
   bbuf = UCase(Trim(b(0)))
   With ret
    Select Case bbuf
     Case ARGS_FLAG_MAXRDLEN
      .max_read_bytes = CInt(Trim(b(1)))
     Case ARGS_FLAG_MODE
      .mode = CLng(Trim(b(1)))
     Case Else
      .code = ERROR_INVALID_ARGS
      parse_args = ret
    End Select
   End With
  Next
 End If
 
 parse_args = ret
End Function


Function get_wordsize_from_info(AppPath As String, Optional args As String = "") As wordsize_struct
 ' +8192 = 2000h = 2*(observed emphirical PE header start pos)
 'Try gathering info thru ShGetFileInfo first
 Dim SHFI As SHFILEINFO
 Dim sh_read   As Long
 Dim intLoWord   As Integer
 Dim intLoWordHiByte As Integer
 Dim intLoWordLoByte As Integer
 Dim strLOWORD   As String
 
 sh_read = -1
 
 Dim ret As wordsize_struct
 struct_prefill ret, AppPath, args
 
 Dim params As wordsize_params
 params = parse_args(args)
 
 'MsgBox "Params end code: " + Str(params.code)
 'MsgBox "Params M: " + Str(params.mode)
 'MsgBox "Params R: " + Str(params.max_read_bytes)
 
 If (params.code <> 0) Then
  ret.code = params.code
 Else
  sh_read = SHGetFileInfo(AppPath, 0, SHFI, Len(SHFI), &H2000)
   
  If ((params.mode = 0) And (sh_read > 0)) Or (params.mode = 1) Then ' if can be read, successfully
  
   If (sh_read > 0) Then
    ret.walkthrough = ret.walkthrough + "SHGetFileInfo=OK|"
   End If
    
   ret.walkthrough = ret.walkthrough + "Mode:WinAPI|"
   
   intLoWord = sh_read And &HFFFF&
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
         ret.code = 0 ' Success!
        Case 1 'SCS_DOS_BINARY
         ' https://users.cs.jmu.edu/abzugcx/Public/Student-Produced-Term-Projects/Operating-Systems-2003-FALL/MS-DOS-by-Dominic-Swayne-Fall-2003.pdf
         ' First known as 86-DOS, it was developed in about 6 weeks by Tim Paterson of Seattle Computer Products (SCP).  The OS was designed to operate on the company’s own 16-bit personal computers running the Intel 8086 microprocessor.  (Paterson, 1983a)
         ret.walkthrough = ret.walkthrough + "SCS_DOS_BINARY|"
         ret.wordsize = 16
         ret.code = 0 ' Success!
        Case 2 'SCS_WOW_BINARY
         ' https://docs.microsoft.com/en-us/windows/win32/api/winbase/nf-winbase-getbinarytypew
         ret.desc = "A 16-bit Windows-based application"
         ret.walkthrough = ret.walkthrough + "SCS_WOW_BINARY|"
         ret.wordsize = 16
         ret.code = 0 ' Success!
        Case 3 'SCS_PIF_BINARY
         ' https://users.cs.jmu.edu/abzugcx/Public/Student-Produced-Term-Projects/Operating-Systems-2003-FALL/MS-DOS-by-Dominic-Swayne-Fall-2003.pdf
         ' First known as 86-DOS, it was developed in about 6 weeks by Tim Paterson of Seattle Computer Products (SCP).  The OS was designed to operate on the company’s own 16-bit personal computers running the Intel 8086 microprocessor.  (Paterson, 1983a)
         ret.desc = " A PIF file that executes an MS-DOS – based application "
         ret.walkthrough = ret.walkthrough + "SCS_PIF_BINARY|"
         ret.wordsize = 16
         ret.code = 0 ' Success!
        Case 4 'SCS_POSIX_BINARY
         ' https://en.wikipedia.org/wiki/Program_information_file
         ' ...
         ' https://stackoverflow.com/q/58986468
         ret.walkthrough = ret.walkthrough + "SCS_POSIX_BINARY|"
         ret.wordsize = 32 ' Posix word wordsize minimal and ambiguous
         ret.code = 7 ' Ambiguous
        Case 5 'SCS_OS216_BINARY
         ' https://docs.microsoft.com/en-us/windows/win32/api/winbase/nf-winbase-getbinarytypew
         ret.desc = " A 16-bit OS/2-based application "
         ret.walkthrough = ret.walkthrough + "SCS_OS216_BINARY|"
         ret.wordsize = 16
         ret.code = 0 ' Success!
        Case 6 'SCS_64BIT_BINARY
         ' https://docs.microsoft.com/en-us/windows/win32/api/winbase/nf-winbase-getbinarytypew
         ret.desc = " A 64-bit Windows-based application. "
         ret.walkthrough = ret.walkthrough + "SCS_64BIT_BINARY|"
         ret.wordsize = 64
         ret.code = 0 ' Success!
       End Select
      Else ' However, if we have, say, Windows NT 3.51, then
       ' WinNT is designed for 32 bits
       ret.walkthrough = ret.walkthrough + "PE&WinNT3.51|"
       ret.wordsize = 32
       ret.code = 0 ' Success!
      End If
     End With
    'End cases
   End Select
  ElseIf ((sh_read = 0) And (params.mode = 0)) Or (params.mode = 2) Then ' If EXE cannot be read
   ret.wordsize = 0
   If (sh_read = 0) Then
    ret.walkthrough = ret.walkthrough + "SHGetFileInfo=BAD|"
   End If
   
   ret.walkthrough = ret.walkthrough + "MODE:RAW|"
   
   Dim pe_buf As String
   Dim pe_buf_len As Long
   
   ret.walkthrough = ret.walkthrough + "Reading " + CStr(params.max_read_bytes) + " bytes|"
   pe_buf = read_binary_file(AppPath, params.max_read_bytes)
   
   pe_buf_len = Len(pe_buf)
   If (pe_buf_len <> params.max_read_bytes) Then
    ret.walkthrough = ret.walkthrough + "Glitchy WinAPI -> read " + CStr(pe_buf_len) + " bytes instead|"
   End If
  
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
    ret.walkthrough = ret.walkthrough + "PE header found|"
    
    Dim pe_nextbytes As String
    pe_nextbytes = Mid(pe_buf, pe_pos + Len(PE_HEADER), 2)
    If (Len(pe_nextbytes)) Then
     If (StrComp(pe_nextbytes, SIGN32, vbBinaryCompare) = 0) Then
      ret.code = 0
      ret.wordsize = 32 '
      ret.walkthrough = ret.walkthrough + "sign 32 detected|"
     ElseIf (StrComp(pe_nextbytes, SIGN64, vbBinaryCompare) = 0) Then
      ret.code = 0
      ret.wordsize = 64 '
      ret.walkthrough = ret.walkthrough + "sign 64 detected|"
     Else
      ret.wordsize = 0 '
      ret.walkthrough = ret.walkthrough + "UNKNOWN (" + str2hexarray(Mid(pe_buf, pe_pos, 10)) + ") @ " + Hex(pe_pos) + "|"
      ret.code = ERROR_UNKNOWN_PE_HEADER
     End If
    End If
   Else
    ret.wordsize = 0 ' prefill
    ret.walkthrough = ret.walkthrough + "NonPE/NonExecutable?|"
    ret.code = ERROR_WARNING_BAD_LUCK
   End If ' If (pe_pos > 0) ...
  Else ' invalid mode
   ret.code = ERROR_INVALID_MODE
  End If
 End If
 get_wordsize_from_info = ret
End Function

Public Sub output_err(errMsg As String)
    CLI.Sendln "Error: " & errMsg
End Sub
Private Function get_error_desc(ByRef iError As Long) As String
 Select Case iError
  Case ERROR_SUCCESS
   get_error_desc = "Success"
  Case ERROR_INVALID_ARGS
   get_error_desc = "Args are invalid"
  Case ERROR_INVALID_MODE
   get_error_desc = "Invalid mode"
  Case ERROR_UNKNOWN_PE_HEADER
   get_error_desc = "Unknown PE header"
  Case ERROR_IRRECOVERABLE
   get_error_desc = "The program encountered an irrecoverable error"
  Case ERROR_WARNING_AMBIGUOUS_WORDSIZE
   get_error_desc = "Success, but the wordsize seems alarmingly ambiguous"
  Case ERROR_WARNING_BAD_LUCK
   get_error_desc = "Fail: No luck analysing"
 End Select
End Function

Public Function quit(code As Long)
    On Error Resume Next

    CLI.Send vbNewLine

    If DEBUGGER Then
        Debug.Print "End"
    Else
        ExitProcess code
    End If
End Function

' https://stackoverflow.com/a/9068210
Public Function GetRunningInIDE() As Boolean
   Dim x As Long
   Debug.Assert Not TestIDE(x)
   GetRunningInIDE = x = 1
End Function

' https://stackoverflow.com/a/9068210
Private Function TestIDE(x As Long) As Boolean
    x = 1
End Function

' original, from simple_capture
Private Function get_unix_time(d As Date) As Long
 get_unix_time = DateDiff("s", "01/01/1970 00:00:00", d)
End Function

Private Function get_unix_time_mod() As String
 get_unix_time_mod = Hex(get_unix_time(Now))
End Function

