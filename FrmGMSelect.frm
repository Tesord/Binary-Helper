VERSION 5.00
Begin VB.Form FrmGMSelect 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Game Mode Select"
   ClientHeight    =   7875
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12180
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7875
   ScaleWidth      =   12180
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdBinHard 
      Caption         =   "Hardcore Mode"
      Height          =   735
      Left            =   9240
      TabIndex        =   6
      Top             =   6480
      Width           =   1815
   End
   Begin VB.CommandButton CmdBinEasy 
      Caption         =   "Beginner Mode"
      Height          =   735
      Left            =   1080
      TabIndex        =   3
      Top             =   6480
      Width           =   1815
   End
   Begin VB.CommandButton CmdDecHard 
      Caption         =   "Hardcore Mode"
      Height          =   735
      Left            =   9240
      TabIndex        =   2
      Top             =   2760
      Width           =   1815
   End
   Begin VB.CommandButton CmdDecNormal 
      Caption         =   "Intermediate Mode"
      Height          =   735
      Left            =   5160
      TabIndex        =   1
      Top             =   2760
      Width           =   1815
   End
   Begin VB.CommandButton CmdBinNormal 
      Caption         =   "Intermediate Mode"
      Height          =   735
      Left            =   5160
      TabIndex        =   5
      Top             =   6480
      Width           =   1815
   End
   Begin VB.CommandButton CmdDecEasy 
      Caption         =   "Beginner Mode"
      Height          =   735
      Left            =   1080
      TabIndex        =   0
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label LbBinGame3 
      Caption         =   $"FrmGMSelect.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   13
      Top             =   4800
      Width           =   11895
   End
   Begin VB.Label LbBinGame4 
      Caption         =   $"FrmGMSelect.frx":00DE
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   12
      Top             =   5400
      Width           =   11895
   End
   Begin VB.Label LbBinGame1 
      Caption         =   "Game Mode 2: Binary to Decimal"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   3840
      Width           =   3735
   End
   Begin VB.Label LbBinGame2 
      Caption         =   $"FrmGMSelect.frx":0234
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   10
      Top             =   4200
      Width           =   11895
   End
   Begin VB.Label LblDecGame3 
      Caption         =   $"FrmGMSelect.frx":0310
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   9
      Top             =   1080
      Width           =   11895
   End
   Begin VB.Label LblDecGame4 
      Caption         =   $"FrmGMSelect.frx":03F9
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   11895
   End
   Begin VB.Label LblDecGame2 
      Caption         =   $"FrmGMSelect.frx":054A
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   7
      Top             =   480
      Width           =   11895
   End
   Begin VB.Label LblDecGame1 
      Caption         =   "Game Mode 1: Decimal to Binary"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "FrmGMSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdBinEasy_Click()

FrmBinGM.Show
FrmGMSelect.Hide
FrmBinGM.Caption = "Game Mode 2 - Binary to Decimal: " & "Beginner    "

End Sub

Private Sub CmdBinHard_Click()

FrmBinGM.Show
FrmGMSelect.Hide
FrmBinGM.Caption = "Game Mode 2 - Binary to Decimal: " & "Hardcore    "

End Sub

Private Sub CmdBinNormal_Click()

FrmBinGM.Show
FrmGMSelect.Hide
FrmBinGM.Caption = "Game Mode 2 - Binary to Decimal: " & "Intermediate"

End Sub

Private Sub CmdDecEasy_Click()

FrmDecGM.Show
FrmGMSelect.Hide
FrmDecGM.Caption = "Game Mode 1 - Decimal to Binary: " & "Beginner    "

End Sub

Private Sub CmdDecHard_Click()

FrmDecGM.Show
FrmGMSelect.Hide
FrmDecGM.Caption = "Game Mode 1 - Decimal to Binary: " & "Hardcore    "

End Sub

Private Sub CmdDecNormal_Click()

FrmDecGM.Show
FrmGMSelect.Hide
FrmDecGM.Caption = "Game Mode 1 - Decimal to Binary: " & "Intermediate"

End Sub

Private Sub Form_Unload(Cancel As Integer)

FrmSelection.Show

End Sub
