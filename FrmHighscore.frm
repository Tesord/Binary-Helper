VERSION 5.00
Begin VB.Form FrmHighscore 
   ClientHeight    =   4725
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   ScaleHeight     =   4725
   ScaleWidth      =   6870
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox Listno 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   3030
      ItemData        =   "FrmHighscore.frx":0000
      Left            =   240
      List            =   "FrmHighscore.frx":0022
      TabIndex        =   7
      Top             =   1200
      Width           =   375
   End
   Begin VB.ListBox ListTime 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3060
      Left            =   4560
      TabIndex        =   3
      Top             =   1200
      Width           =   2055
   End
   Begin VB.ListBox ListName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3060
      Left            =   600
      TabIndex        =   0
      Top             =   1200
      Width           =   3255
   End
   Begin VB.ListBox ListScore 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3060
      Left            =   3840
      TabIndex        =   2
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Time (In secs)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   6
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Score"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   5
      Top             =   840
      Width           =   855
   End
   Begin VB.Label LblName 
      Alignment       =   2  'Center
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   840
      Width           =   3135
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Highschool Table"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "FrmHighscore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim scorelist(30) As String
Private Sub Form_Load()

Counter = 1

If Mid$(FrmBinGM.Caption, 34, 12) = "Beginner    " Then
    If Dir("N:\IT Misc\BtDGameRecord - Easy.txt") <> "" Then
        Open "N:\IT Misc\BtDGameRecord - Easy.txt" For Input As #1     'Create a new file with the concatenation result as name in slot 1
            Do Until EOF(1)
                Input #1, Data
                scorelist(Counter) = Data
                Counter = Counter + 1
                EOF (1)
            Loop
        Close #1      'Close Writing Operation
    End If

    For Counter = 1 To 10
        ListScore.AddItem scorelist(Counter)
    Next
    For Counter = 11 To 20
        ListTime.AddItem scorelist(Counter)
    Next
    For Counter = 21 To 30
        ListName.AddItem scorelist(Counter)
    Next

    FrmHighscore.Caption = "Highscore Table (Easy Mode)"

End If


If Mid$(FrmBinGM.Caption, 34, 12) = "Intermediate" Then
    If Dir("N:\IT Misc\BtDGameRecord - Normal.txt") <> "" Then
        Open "N:\IT Misc\BtDGameRecord - Normal.txt" For Input As #1     'Create a new file with the concatenation result as name in slot 1
            Do Until EOF(1)
                Input #1, Data
                scorelist(Counter) = Data
                Counter = Counter + 1
                EOF (1)
            Loop
        Close #1      'Close Writing Operation
    End If

    For Counter = 1 To 10
        ListScore.AddItem scorelist(Counter)
    Next
    For Counter = 11 To 20
        ListTime.AddItem scorelist(Counter)
    Next
    For Counter = 21 To 30
        ListName.AddItem scorelist(Counter)
    Next

    FrmHighscore.Caption = "Highscore Table (Normal Mode)"
    
End If


If Mid$(FrmBinGM.Caption, 34, 12) = "Hardcore    " Then
    If Dir("N:\IT Misc\BtDGameRecord - Hard.txt") <> "" Then
        Open "N:\IT Misc\BtDGameRecord - Hard.txt" For Input As #1     'Create a new file with the concatenation result as name in slot 1
            Do Until EOF(1)
                Input #1, Data
                scorelist(Counter) = Data
                Counter = Counter + 1
                EOF (1)
            Loop
        Close #1      'Close Writing Operation
    End If

    For Counter = 1 To 10
        ListScore.AddItem scorelist(Counter)
    Next
    For Counter = 11 To 20
        ListName.AddItem scorelist(Counter)
    Next
    
    Label3.Visible = False
    ListTime.Visible = False

    FrmHighscore.Caption = "Highscore Table (Hardcore Mode)"
    
End If


End Sub

