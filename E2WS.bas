Attribute VB_Name = "E2WS"
Option Explicit

' Used by: get_wordsize_from_info
Declare Function GetBinaryType Lib "kernel32" Alias "GetBinaryTypeA" (ByVal lpApplicationName As String, lpBinaryType As Long) As Long

' Used by: get_wordsize_from_info
Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Integer

' Used by: get_wordsize_from_info
Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
 
' Deprecated
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
 path As String
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
Private Const ERROR_UNCHANGED = -2 ' error code was never changed
Private Const ERROR_IRRECOVERABLE = -1
Private Const ERROR_SUCCESS = 0 ' all good
Private Const ERROR_INVALID_ARGS = 1 'problem with ARGS
Private Const ERROR_INVALID_MODE = 2 ' invalid mode
Private Const ERROR_INVALID_FILE = 3 ' invalid file
Private Const ERROR_UNKNOWN_PE_HEADER = 4 ' unknown PE header
Private Const ERROR_WARNING_AMBIGUOUS_WORDSIZE = 7 ' as 777
Private Const ERROR_WARNING_BAD_LUCK = 13 ' https://en.wikipedia.org/wiki/13_(number)#Unlucky_13

Private Const MODE_FLEXI = 0 ' Flexible
Private Const MODE_WINAPI = 1 ' WinAPI only
Private Const MODE_RAW = 2 ' RAW only

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
 
 ' https://superuser.com/a/889267/1113462
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
  .walkthrough = .walkthrough & "2JSON|" ' for logging
 
  buf = "{" & vbCrLf
  buf = buf & String(1, vbTab) & C34 & App.Title & C34 & ":{" & vbCrLf
  
  ' file
  buf = buf & String(2, vbTab) & C34 & "path" & C34 & ": " & C34 & .path & C34 & JSON_PARAMS_DELIM
  
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

' https://www.tek-tips.com/viewthread.cfm?qid=457979
' no WinAPI implementation
Public Function FileExists(FileName As String) As Boolean
  Dim intFileNum As Integer
  intFileNum = FreeFile
  On Error GoTo NoSuchFile:
  Open FileName For Input As #intFileNum
  Close intFileNum
  'MsgBox "File Exists - True"
  FileExists = True
  Exit Function
NoSuchFile:
  'MsgBox "File Exists - False"
  FileExists = False
  
End Function

Function read_binary_file(path As String, Optional target_len As Integer = 2) As Byte()
    Dim nFile As Long
    Dim file_len As Long
    Dim final_len As Long
    
    nFile = FreeFile
    
    Open path For Binary Access Read As nFile
    file_len = FileLen(path)
    final_len = IIf(file_len < target_len, file_len, target_len) ' min function implementation
    If LOF(nFile) > 0 Then
        read_binary_file = Input(final_len, nFile)
        'ReDim read_binary_file(0 To LOF(nFile) - 1)
        'Get nFile, , read_binary_file
    End If
    Close nFile
End Function


Private Function struct_prefill(s As wordsize_struct, AppPath As String, args As String)
 With s
  .walkthrough = "RDY|"
  .code = ERROR_UNCHANGED
  .path = AppPath
  .time = get_unix_time_mod
  .args = args
 End With
End Function

Function wordsize_params_init(ret As wordsize_params)
 With ret
  .max_read_bytes = 8192
  .mode = MODE_FLEXI
  .code = 0
 End With

End Function

Function set_wordsize(struct As wordsize_struct, wordsize As Byte, Optional error_code = ERROR_SUCCESS)
 With struct
  .wordsize = wordsize
   .code = error_code
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
 
 Dim ret As wordsize_struct
 struct_prefill ret, AppPath, args
 
 If (DEBUGGER = False) Then
  On Error GoTo ErrorHandler
 End If
 
 Dim SHFI As SHFILEINFO
 Dim sh_read   As Long
 Dim intLoWord   As Integer
 Dim intLoWordHiByte As Integer
 Dim intLoWordLoByte As Integer
 Dim strLOWORD   As String
 
 sh_read = -1

 Dim params As wordsize_params
 params = parse_args(args)
 
 'MsgBox "Params end code: " + Str(params.code)
 'MsgBox "Params M: " + Str(params.mode)
 'MsgBox "Params R: " + Str(params.max_read_bytes)
 
 If (params.code <> 0) Then
  ret.code = params.code
 ElseIf (params.code = 0) Then
  
  If (params.mode = MODE_FLEXI) Then
   ret.walkthrough = ret.walkthrough + "Mode:Flexi|"
  End If
  
  If (FileExists(AppPath)) Then
   ret.walkthrough = ret.walkthrough + "File:Found|"
  
   sh_read = SHGetFileInfo(AppPath, 0, SHFI, Len(SHFI), &H2000)
    
   If ((params.mode = MODE_FLEXI) And (sh_read > 0)) Or (params.mode = MODE_WINAPI) Then ' if can be read, successfully
   
    If (sh_read > 0) Then
     ret.walkthrough = ret.walkthrough + "SHGetFileInfo=OK|"
    End If
     
    ret.walkthrough = ret.walkthrough + "Mode:WinAPI|"
    
    intLoWord = sh_read And &HFFFF&
    intLoWordHiByte = intLoWord \ &H100 And &HFF&
    intLoWordLoByte = intLoWord And &HFF&
    strLOWORD = Chr$(intLoWordLoByte) & Chr$(intLoWordHiByte)
       
    Select Case strLOWORD
     Case "MZ"
      ' 1) MZ is older than NE and LX, thus NE is 16bit, thus MZ is <= 16 bit
      ' 2) https://superuser.com/a/1334151/1113462
      ret.walkthrough = ret.walkthrough + "LOWORD:MZ|"
      set_wordsize ret, 16
     Case "LX"
      ' 1) http://www.textfiles.com/programming/FORMATS/lxexe.txt
      ' 2) https://superuser.com/a/1334151/1113462
      ' 3) https://faydoc.tripod.com/formats/exe-LE.htm
      ret.walkthrough = ret.walkthrough + "LOWORD:LX|"
      set_wordsize ret, 32
     Case "LE"
      ' 1) https://moddingwiki.shikadi.net/wiki/Linear_Executable_(LX/LE)_Format
      ' 2) https://www.program-transformation.org/Transform/PcExeFormat
      ' 3) https://github.com/gameprive/win2k/blob/4abce2f1531739102d49db5c9f9e20e1e2d0de71/private/windbg64/langapi/include/exe_vxd.h
      ret.desc = " DOS4GW-based executable, 16/32bit mixed codebase "
      ret.walkthrough = ret.walkthrough + "LOWORD:LE|"
      set_wordsize ret, 32
     Case "NE" ' NE is 16-bit, ref: https://en.wikipedia.org/wiki/New_Executable
      ret.walkthrough = ret.walkthrough + "LOWORD:NE|"
      set_wordsize ret, 16
     Case "PE" ' If PE app, gather OS info
      ret.walkthrough = ret.walkthrough + "LOWORD:PE|"
      Dim OSV As OSVERSIONINFO
      With OSV
       .OSVSize = Len(OSV)
       GetVersionEx OSV
       If .PlatformID < 2 Then ' If PE app and Win 9x
        ret.walkthrough = ret.walkthrough + "PE&Win9x|"
        set_wordsize ret, 32
       ' If PE app and Windows NT 3.51, 4 or higher
       ' References:
       ' 1) https://www.geoffchappell.com/studies/windows/win32/kernel32/api/index.htm
       ' 2) https://www.swissdelphicenter.ch/en/showcode.php?id=126
       ElseIf (.dwVerMajor = 3 And Left(Str(.dwVerMinor), 1) = "5") Or (.dwVerMajor >= 4) Then
        ret.walkthrough = ret.walkthrough + "PE&WinNT-Modern|"
        ' Get info via GetBinaryType
        Dim BinaryType As Long
        GetBinaryType AppPath, BinaryType
        Select Case BinaryType
         Case 0 'SCS_32BIT_BINARY
          ' https://docs.microsoft.com/en-us/windows/win32/api/winbase/nf-winbase-getbinarytypew
          ret.desc = " A 32-bit Windows-based application "
          ret.walkthrough = ret.walkthrough + "SCS_32BIT_BINARY|"
          set_wordsize ret, 32
         Case 1 'SCS_DOS_BINARY
          ' https://users.cs.jmu.edu/abzugcx/Public/Student-Produced-Term-Projects/Operating-Systems-2003-FALL/MS-DOS-by-Dominic-Swayne-Fall-2003.pdf
          ' First known as 86-DOS, it was developed in about 6 weeks by Tim Paterson of Seattle Computer Products (SCP).  The OS was designed to operate on the company’s own 16-bit personal computers running the Intel 8086 microprocessor.  (Paterson, 1983a)
          ret.walkthrough = ret.walkthrough + "SCS_DOS_BINARY|"
          set_wordsize ret, 16
         Case 2 'SCS_WOW_BINARY
          ' https://docs.microsoft.com/en-us/windows/win32/api/winbase/nf-winbase-getbinarytypew
          ret.desc = "A 16-bit Windows-based application"
          ret.walkthrough = ret.walkthrough + "SCS_WOW_BINARY|"
          set_wordsize ret, 16
         Case 3 'SCS_PIF_BINARY
          ' https://users.cs.jmu.edu/abzugcx/Public/Student-Produced-Term-Projects/Operating-Systems-2003-FALL/MS-DOS-by-Dominic-Swayne-Fall-2003.pdf
          ' First known as 86-DOS, it was developed in about 6 weeks by Tim Paterson of Seattle Computer Products (SCP).  The OS was designed to operate on the company’s own 16-bit personal computers running the Intel 8086 microprocessor.  (Paterson, 1983a)
          ret.desc = "A PIF file that executes an MS-DOS – based application"
          ret.walkthrough = ret.walkthrough + "SCS_PIF_BINARY|"
          set_wordsize ret, 16
         Case 4 'SCS_POSIX_BINARY
          ' https://en.wikipedia.org/wiki/Program_information_file
          ' ...
          ' https://stackoverflow.com/q/58986468
          ret.walkthrough = ret.walkthrough + "SCS_POSIX_BINARY|"
          ret.desc = "Posix word wordsize is minimal and ambiguous"
          set_wordsize ret, 32, 7 ' Ambiguous - override
         Case 5 'SCS_OS216_BINARY
          ' https://docs.microsoft.com/en-us/windows/win32/api/winbase/nf-winbase-getbinarytypew
          ret.desc = " A 16-bit OS/2-based application "
          ret.walkthrough = ret.walkthrough + "SCS_OS216_BINARY|"
          set_wordsize ret, 16
         Case 6 'SCS_64BIT_BINARY
          ' https://docs.microsoft.com/en-us/windows/win32/api/winbase/nf-winbase-getbinarytypew
          ret.desc = " A 64-bit Windows-based application. "
          ret.walkthrough = ret.walkthrough + "SCS_64BIT_BINARY|"
          set_wordsize ret, 64
        End Select
       Else
        ' However, if we have, say, Windows NT 3.1 and PE format, then
        ' Minimum byte-size applicable is always returned
        ' However, as NE was a de-facto standard for 16-bit applications,
        ' NT 3.1 would only realistically have PE files is they are 32-bit
        ' References,
        ' https://en.wikipedia.org/wiki/New_Executable
        ' https://en.wikipedia.org/wiki/Portable_Executable
        ' https://en.wikipedia.org/wiki/Windows_NT_3.1
        ' https://en.wikipedia.org/wiki/File:NT_3.1_layers.png
        ret.walkthrough = ret.walkthrough + "PE&WinNT3.X|"
        set_wordsize ret, 32
       End If
      End With
     'End cases
    End Select
   ElseIf ((sh_read = 0) And (params.mode = MODE_FLEXI)) Or (params.mode = MODE_RAW) Then ' If EXE cannot be read
    ' exceptional - task is not finished yet
    ret.wordsize = 0
    If (sh_read = 0) Then
     ret.walkthrough = ret.walkthrough + "SHGetFileInfo=Bad|"
    End If
    
    ret.walkthrough = ret.walkthrough + "Mode:RAW|"
    
    Dim pe_buf As String
    Dim pe_buf_len As Long
    
    ret.walkthrough = ret.walkthrough + "Read <=" + CStr(params.max_read_bytes) + " bytes|"
    pe_buf = read_binary_file(AppPath, params.max_read_bytes)
    
    pe_buf_len = Len(pe_buf)
    If (pe_buf_len > params.max_read_bytes) Then
     ret.walkthrough = ret.walkthrough + "Glitchy WinAPI - it had read " + CStr(pe_buf_len) + " bytes instead|"
    End If
   
    'Dim iFileNo As Integer
    'iFileNo = FreeFile
    'Open "C:\Test.txt" For Output As #iFileNo
    'Print #iFileNo, str2hexarray(pe_buf)
    'Form1.Text2.Text = str2hexarray(pe_buf)
    'Close #iFileNo
    
    ' https://superuser.com/a/889267/1113462
    Dim pe_pos As Long
    pe_pos = InStr(1, pe_buf, PE_HEADER, vbBinaryCompare)
  
    If (pe_pos > 0) Then
     ret.walkthrough = ret.walkthrough + "PE header found|"
     
     Dim pe_nextbytes As String
     pe_nextbytes = Mid(pe_buf, pe_pos + Len(PE_HEADER), 2)
     If (Len(pe_nextbytes)) Then
      If (StrComp(pe_nextbytes, SIGN32, vbBinaryCompare) = 0) Then
       ret.walkthrough = ret.walkthrough + "sign 32 detected|"
       set_wordsize ret, 32
      ElseIf (StrComp(pe_nextbytes, SIGN64, vbBinaryCompare) = 0) Then
       ret.walkthrough = ret.walkthrough + "sign 64 detected|"
       set_wordsize ret, 64
      Else
       ret.walkthrough = ret.walkthrough + "UNKNOWN (" + str2hexarray(Mid(pe_buf, pe_pos, 10)) + ") @ " + Hex(pe_pos) + "|"
       set_wordsize ret, 0, ERROR_UNKNOWN_PE_HEADER
      End If
     End If
    Else
     set_wordsize ret, 0, ERROR_UNKNOWN_PE_HEADER
    End If ' If (pe_pos > 0) ...
   Else ' invalid mode
    set_wordsize ret, 0, ERROR_INVALID_MODE
   End If
  Else
   ret.walkthrough = ret.walkthrough + "File:Unavailable|"
   ret.code = ERROR_INVALID_FILE
  End If
 Else
  set_wordsize ret, 0, ERROR_INVALID_MODE
 End If
 get_wordsize_from_info = ret
 Exit Function
 
ErrorHandler:
  ret.desc = "Error #" + Str(Err.Number) + ": " + Err.Description
  ret.code = ERROR_IRRECOVERABLE
  ret.wordsize = 0
 
  get_wordsize_from_info = ret
End Function

Public Sub output_err(errMsg As String)
    CLI.Sendln "Error: " & errMsg
End Sub
Private Function get_error_desc(ByRef myError As Long) As String
 Select Case myError
  Case ERROR_SUCCESS
   get_error_desc = "Success"
  Case ERROR_INVALID_ARGS
   get_error_desc = "Args are invalid"
  Case ERROR_INVALID_MODE
   get_error_desc = "Invalid mode"
  Case ERROR_INVALID_FILE
   get_error_desc = "File cannot be accessed"
  Case ERROR_UNKNOWN_PE_HEADER
   get_error_desc = "Unknown PE header"
  Case ERROR_IRRECOVERABLE
   get_error_desc = "The program encountered an irrecoverable error"
  Case ERROR_WARNING_AMBIGUOUS_WORDSIZE
   get_error_desc = "Success, but the wordsize seems alarmingly ambiguous"
  Case ERROR_WARNING_BAD_LUCK
   get_error_desc = "Fail: No luck analysing"
  Case ERROR_UNCHANGED
   get_error_desc = "Coder's error: Error code was never changed"
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
 ' returns UPPERCASE HEX-string, can be lowercased at extra computing
 ' therefore, decided to keep as is
 get_unix_time_mod = Hex(get_unix_time(Now))
End Function

