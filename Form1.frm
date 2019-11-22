VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5490
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   16950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   16950
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   5295
      Left            =   6000
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   120
      Width           =   10815
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
      TabIndex        =   5
      Top             =   3360
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   2640
      Width           =   4455
   End
   Begin VB.Label Label4 
      Caption         =   "IDK"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   4
      Top             =   4440
      Width           =   5535
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
    Dim dat  As GetEXEWordSize_out

    Beep
    dat = Module1.GetEXEWordSize(Text1.Text)
    Label4.Caption = GetEXEWordSize_ToString(dat)
End Sub

Private Sub Form_Load()
    Dim mypath As String

    ' https://stackoverflow.com/a/12423852/12258312
    mypath = App.path & IIf(Right$(App.path, 1) <> "\", "\", "") & App.EXEName & ".exe"
    Text1.Text = mypath
End Sub
 
