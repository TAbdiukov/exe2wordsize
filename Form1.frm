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

    Beep
    dat = E2WS.get_wordsize_from_info(Text1.Text)
    Text3.Text = E2WS.struct_to_json(dat)
End Sub

Private Sub Form_Load()
 CLI.setup
 E2WS.setup

 Dim mypath As String
 Text1.Text = E2WS.app_path_exe
 
 Me.Caption = App.EXEName & ": my NES was 128 bit mhmm"
End Sub
Private Sub showHelp()
        CLI.SetTextColour CLI.FOREGROUND_RED Or CLI.FOREGROUND_GREEN Or CLI.FOREGROUND_INTENSITY
        CLI.Sendln "AppModeChange - CLI mod v" + VER
        CLI.SetTextColour CLI.FOREGROUND_RED Or CLI.FOREGROUND_GREEN Or CLI.FOREGROUND_BLUE Or CLI.FOREGROUND_INTENSITY
        CLI.Sendln "(Original GUI code by Nirsoft)"
        CLI.Sendln ""
        
        CLI.SetTextColour CLI.FOREGROUND_RED Or CLI.FOREGROUND_INTENSITY
        CLI.Sendln "USAGE:"
        CLI.SetTextColour CLI.FOREGROUND_RED Or CLI.FOREGROUND_GREEN Or CLI.FOREGROUND_BLUE
        CLI.Sendln "amc <path_to_app> <new_mode>"
        CLI.Sendln ""
        
        CLI.SetTextColour CLI.FOREGROUND_GREEN Or CLI.FOREGROUND_INTENSITY
        CLI.Sendln "FOR EXAMPLE:"
        CLI.SetTextColour CLI.FOREGROUND_RED Or CLI.FOREGROUND_GREEN Or CLI.FOREGROUND_BLUE
        CLI.Sendln "amc " + Chr(34) + "C:/Projects/My supa CLI project/Project1.exe" + Chr(34) + " 3"
        CLI.Sendln "(to set the Project1 application to the CLI mode)"
        CLI.Sendln ""
        
        
        CLI.SetTextColour CLI.FOREGROUND_GREEN Or CLI.FOREGROUND_BLUE Or CLI.FOREGROUND_INTENSITY
        CLI.Sendln "MANUAL:"
        CLI.SetTextColour CLI.FOREGROUND_RED Or CLI.FOREGROUND_GREEN Or CLI.FOREGROUND_BLUE
        CLI.Sendln "<path_to_app> - Path to your executable. " + Chr(34) + "-tolerable"
        CLI.Sendln ""
        CLI.Sendln "<new_mode> - New app SUBSYSTEM mode to set"
        CLI.Sendln "Informally, one'd need to only know of modes: 2 (CLI) and 3 (GUI)"
        CLI.Sendln "But below all known modes are listed:"
        
        Dim i As Integer
End Sub

