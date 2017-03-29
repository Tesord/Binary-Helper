VERSION 5.00
Begin VB.Form FrmSelection 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Mode"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   2580
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   2580
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGameMode 
      Caption         =   "Game Mode"
      Height          =   735
      Left            =   360
      TabIndex        =   2
      Top             =   1920
      Width           =   1815
   End
   Begin VB.CommandButton CmdConvertMode 
      Caption         =   "Convert Mode"
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Select Mode"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold Condensed"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "FrmSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdConvertMode_Click()

FrmConvertMode.Show
FrmSelection.Hide

End Sub

Private Sub cmdGameMode_Click()

FrmGMSelect.Show
FrmSelection.Hide

End Sub
