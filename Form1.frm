VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   9150
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7860
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9150
   ScaleWidth      =   7860
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   5040
      TabIndex        =   7
      Top             =   2640
      Width           =   2535
   End
   Begin VB.TextBox Text3 
      Height          =   4575
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   6
      Text            =   "Form1.frx":0ECA
      Top             =   4440
      Width           =   7335
   End
   Begin VB.TextBox Text2 
      Height          =   5295
      Left            =   9840
      TabIndex        =   5
      Text            =   "Text2"
      Top             =   120
      Width           =   6975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Try Me!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   4
      Top             =   3360
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   2640
      Width           =   4455
   End
   Begin VB.Label Label4 
      Caption         =   "Argz"
      Height          =   495
      Left            =   5040
      TabIndex        =   8
      Top             =   2160
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "Output (result):"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   4080
      Width           =   4455
   End
   Begin VB.Label Label2 
      Caption         =   "Input (path):"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   2280
      Width           =   4455
   End
   Begin VB.Label Label1 
      Caption         =   "IF YOU SEE THIS FORM, YOU HAVENT COMPILED INTO CONSOLE APP AS REQUIRED. GUI BELOW FOR TESTING PURPOSES ONLY"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   5535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
 Dim dat  As wordsize_struct
 
 dat = E2WS.get_wordsize_from_info(Text1.Text, Text4.Text)
 
 Text3.Text = E2WS.struct_to_json(dat)
End Sub

Private Sub Form_Load()
 CLI.setup
 E2WS.setup
    
 ' Initialise args
 Dim argw() As String
 Dim argc As Integer
 Dim argt As String

 Dim mypath As String
 Text1.Text = E2WS.app_path_exe
 
 Me.Caption = App.EXEName & ": my NES was 128 bit mhmm"
 
 argt = Trim(Command)
 argw = Split(argt, "*")
 argc = UBound(argw) - LBound(argw) + 1 ' https://forums.windowssecrets.com/showthread.php/28214-counting-array-elements-(vb6)
 
 If (Len(argt)) Then ' If number of args suffice
  Dim path As String
  
  path = argw(0)
  path = Replace(path, Chr(34), "")
  
  Dim out As wordsize_struct
  
  If (argc > 1) Then
   Dim real_args As String
   real_args = Trim(argw(1))
   out = E2WS.get_wordsize_from_info(path, real_args)
   CLI.Sendln E2WS.struct_to_json(out)
   quit out.code
  Else
   out = E2WS.get_wordsize_from_info(path)
   CLI.Sendln E2WS.struct_to_json(out)
   quit out.code
  End If
 Else
  showHelp
  quit 0
 End If
End Sub
Private Sub showHelp()
 CLI.SetTextColour CLI.FOREGROUND_RED Or CLI.FOREGROUND_GREEN Or CLI.FOREGROUND_INTENSITY
 CLI.Sendln "exe2wordsize v" + VER
 CLI.SetTextColour CLI.FOREGROUND_RED Or CLI.FOREGROUND_GREEN Or CLI.FOREGROUND_BLUE Or CLI.FOREGROUND_INTENSITY
 CLI.Sendln ""
 
 CLI.SetTextColour CLI.FOREGROUND_RED Or CLI.FOREGROUND_INTENSITY
 CLI.Sendln "USAGE:"
 CLI.SetTextColour CLI.FOREGROUND_RED Or CLI.FOREGROUND_GREEN Or CLI.FOREGROUND_BLUE
 CLI.Sendln "exe2wordsize <path_to_app>"
 CLI.Sendln "exe2wordsize <path_to_app> * <args>"
 CLI.Sendln ""
 
 CLI.SetTextColour CLI.FOREGROUND_GREEN Or CLI.FOREGROUND_INTENSITY
 CLI.Sendln "FOR EXAMPLE:"
 CLI.SetTextColour CLI.FOREGROUND_RED Or CLI.FOREGROUND_GREEN Or CLI.FOREGROUND_BLUE
 CLI.Sendln "exe2wordsize " + "C:/Projects/idk/Project1.exe"
 CLI.Sendln "exe2wordsize " + Chr(34) + "C:/Projects/idk/Project1.exe" + Chr(34) + " * M=2 R=8192"
 CLI.Sendln ""
 
 
 CLI.SetTextColour CLI.FOREGROUND_GREEN Or CLI.FOREGROUND_BLUE Or CLI.FOREGROUND_INTENSITY
 CLI.Sendln "MANUAL:"
 CLI.SetTextColour CLI.FOREGROUND_RED Or CLI.FOREGROUND_GREEN Or CLI.FOREGROUND_BLUE
 CLI.Sendln "<path_to_app> - Path to your executable. " + Chr(34) + "-tolerable"
 CLI.Sendln ""
 CLI.Sendln "* - Delimiter required if you use args."
 CLI.Sendln "(Hint: Don't have to use asterick if no args required)"
 CLI.Sendln ""
 CLI.Sendln "<args> - Extra arguments, space-delimited. Supported args below,"
 CLI.Sendln "# M=(number) - Set analysis mode. Modes supported,"
 CLI.Sendln "## 0 - Automatic and flexible (Default)"
 CLI.Sendln "## 1 - Rely only on WinAPI. 64-bit input may be unreliable"
 CLI.Sendln "## 2 - Rely only on raw-reading. Only 32/64-bit detection, false-pos theoretically possible"
 CLI.Sendln "# R=(number) - In raw-reading mode (M=2), how many bytes to read at most for analysis"
 CLI.Sendln "  (Hint: Only applicable in MODE = 2. Unused in other modes)"
 
 
 CLI.SetTextColour CLI.FOREGROUND_RED Or CLI.FOREGROUND_BLUE Or CLI.FOREGROUND_INTENSITY
 CLI.Sendln "OUTPUT:"
 CLI.SetTextColour CLI.FOREGROUND_RED Or CLI.FOREGROUND_GREEN Or CLI.FOREGROUND_BLUE
 CLI.Sendln "In JSON format, rather straightforward"
End Sub

