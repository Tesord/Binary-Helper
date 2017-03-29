VERSION 5.00
Begin VB.Form FrmDecGM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Game Mode 1 - Decimal to Binary"
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   9045
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrgame 
      Enabled         =   0   'False
      Left            =   6480
      Top             =   1680
   End
   Begin VB.CommandButton CmdShowhiscore 
      Caption         =   "Show High-Score"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7200
      TabIndex        =   22
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CommandButton cmdsavehighscore 
      Caption         =   "Save High-Score"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7200
      TabIndex        =   21
      Top             =   720
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   120
      TabIndex        =   6
      Top             =   2640
      Visible         =   0   'False
      Width           =   8775
      Begin VB.OptionButton Option4 
         Caption         =   "Option1"
         Height          =   315
         Left            =   4560
         TabIndex        =   20
         Top             =   720
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Option1"
         Height          =   255
         Left            =   1920
         TabIndex        =   19
         Top             =   720
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   255
         Left            =   1920
         TabIndex        =   18
         Top             =   360
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Option1"
         Height          =   315
         Left            =   4560
         TabIndex        =   17
         Top             =   360
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton CmdConfirm 
         Caption         =   "Confirm!"
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
         Left            =   3120
         TabIndex        =   16
         Top             =   2040
         Width           =   2415
      End
      Begin VB.ComboBox Combo7 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         ItemData        =   "FrmGameMode.frx":0000
         Left            =   5760
         List            =   "FrmGameMode.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   1440
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.ComboBox Combo6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         ItemData        =   "FrmGameMode.frx":0014
         Left            =   5280
         List            =   "FrmGameMode.frx":001E
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1440
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.ComboBox Combo5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         ItemData        =   "FrmGameMode.frx":0028
         Left            =   4800
         List            =   "FrmGameMode.frx":0032
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1440
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.ComboBox Combo4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         ItemData        =   "FrmGameMode.frx":003C
         Left            =   4320
         List            =   "FrmGameMode.frx":0046
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1440
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         ItemData        =   "FrmGameMode.frx":0050
         Left            =   3840
         List            =   "FrmGameMode.frx":005A
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1440
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         ItemData        =   "FrmGameMode.frx":0064
         Left            =   3360
         List            =   "FrmGameMode.frx":006E
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1440
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         ItemData        =   "FrmGameMode.frx":0078
         Left            =   2880
         List            =   "FrmGameMode.frx":0082
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1440
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.ComboBox Combo0 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         ItemData        =   "FrmGameMode.frx":008C
         Left            =   2400
         List            =   "FrmGameMode.frx":0096
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1440
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Lbloption1 
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
         Left            =   2160
         TabIndex        =   26
         Top             =   360
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Lbloption2 
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
         Left            =   4800
         TabIndex        =   25
         Top             =   360
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Lbloption3 
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
         Left            =   2160
         TabIndex        =   24
         Top             =   720
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Lbloption4 
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
         Left            =   4800
         TabIndex        =   23
         Top             =   720
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "128   64   32    16     8     4      2      1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   7
         Top             =   1080
         Visible         =   0   'False
         Width           =   3855
      End
   End
   Begin VB.CommandButton Cmdgamestart 
      Caption         =   "Game Start!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   1695
   End
   Begin VB.CommandButton cmdselectmode 
      Caption         =   "Select Mode"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Lbltime 
      Height          =   255
      Left            =   7440
      TabIndex        =   27
      Top             =   240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label LblDecTitle 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3360
      TabIndex        =   5
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Label lblbindesceasy 
      Caption         =   $"FrmGameMode.frx":00A0
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2040
      TabIndex        =   4
      Top             =   240
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.Label lblbindeschard 
      Caption         =   "Hardcore Mode: The user needs to enter the binary equivalent of random decimal values with no help from the Binary Aid Table!!!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2040
      TabIndex        =   3
      Top             =   240
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.Label lblbindescnormal 
      Caption         =   $"FrmGameMode.frx":013F
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2040
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   5055
   End
End
Attribute VB_Name = "FrmDecGM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mintcount As Integer
Dim clockorder As String
Dim multchans(4) As String
Dim queno As Integer
Dim choiceque(4) As OptionButton
Dim anslabel(4) As Label
Dim gamefinish As String

Dim decval As Integer
Dim finalbinval As String
Dim correctans As Integer
Dim correctanscount As Integer

Dim scorelist(31) As String
Dim scoreval(10, 2) As Integer

Private Sub CmdConfirm_Click()

If Mid$(FrmDecGM.Caption, 34, 12) = "Beginner    " Then
    
    finalbinval = 0
    
    Set choiceque(1) = Option1
    Set choiceque(2) = Option2
    Set choiceque(3) = Option3
    Set choiceque(4) = Option4

    For Counter = 1 To 4
        If Counter = correctans And choiceque(Counter).Value = True Then trueval = 1
    Next
    If trueval <> 1 Then
        MsgBox ("Wrong answer! The correct answer was ") & multchans(correctans) & "."
    Else
        correctanscount = correctanscount + 1
    End If
    
    finalbinval = ""
    
    Call DtBwholenoconv(decval, finalbinval)
    Call dtbeasymode(decval, finalbinval, correctans, multchans(), queno, gamefinish)
    
    For Counter = 1 To 4
        choiceque(Counter).Value = False
    Next
    
    If gamefinish = 1 Then
        MsgBox ("Game Finish! You have answered " & correctanscount & " question(s) correctly. Your time was " & mintcount & " seconds!")
        Frame1.Visible = False
        tmrgame.Enabled = False
        CmdShowhiscore.Visible = True
        cmdsavehighscore.Visible = True
    End If
    
End If


If Mid$(FrmDecGM.Caption, 34, 12) = "Intermediate" Then

    If Text1.Text <> "" Then
        If Text1.Text = finalbinval Then trueval = 1
    End If
    
    If trueval <> 1 Then
        MsgBox ("Wrong answer! The correct answer was ") & finalbinval & "."
    Else
        correctanscount = correctanscount + 1
    End If
    
    finalbinval = ""
    
    Call DtBwholenoconv(decval, finalbinval)
    Call dtbnormalmode(decval, queno, gamefinish)
    
    Text1.Text = ""
    
    If gamefinish = 1 Then
        MsgBox ("Game Finish! You have answered " & correctanscount & " question(s) correctly. Your time was " & mintcount & " seconds!")
        Frame1.Visible = False
        tmrgame.Enabled = False
        CmdShowhiscore.Visible = True
        cmdsavehighscore.Visible = True
    End If
    
End If


If Mid$(FrmDecGM.Caption, 34, 12) = "Hardcore    " Then
    
    If Text1.Text <> "" Then
        If Text1.Text = finalbinval Then trueval = 1
    End If
    
    If trueval <> 1 Then
        MsgBox ("Wrong answer! The correct answer was ") & finalbinval & "."
        finalbinval = ""
        Call DtBwholenoconv(decval, finalbinval)
        Call dtbhardmode(decval)
    Else
        correctanscount = correctanscount + 1
        finalbinval = ""
        Call DtBwholenoconv(decval, finalbinval)
        Call dtbhardmode(decval)
    End If
    
    Text1.Text = ""
    
End If

End Sub

Private Sub Cmdgamestart_Click()

Lbltime.Visible = True
Frame1.Visible = True
finalbinval = ""

If Mid$(FrmDecGM.Caption, 34, 12) = "Beginner    " Then

    Option1.Visible = True
    Option2.Visible = True
    Option3.Visible = True
    Option4.Visible = True
    Lbloption1.Visible = True
    Lbloption2.Visible = True
    Lbloption3.Visible = True
    Lbloption4.Visible = True
    Label1.Visible = True
    gamefinish = 0
    clockorder = 0
    mintcount = 0
    Cls
    
    Call DtBwholenoconv(decval, finalbinval)
    Call dtbeasymode(decval, finalbinval, correctans, multchans(), queno, gamefinish)
End If


If Mid$(FrmDecGM.Caption, 34, 12) = "Intermediate" Then
    Label1.Visible = True
    Combo0.Visible = True
    Combo1.Visible = True
    Combo2.Visible = True
    Combo3.Visible = True
    Combo4.Visible = True
    Combo5.Visible = True
    Combo6.Visible = True
    Combo7.Visible = True
    gamefinish = 0
    clockorder = 0
    mintcount = 0
    Cls
    
    Call DtBwholenoconv(decval, finalbinval)
    Call dtbnormalmode(decval, queno, gamefinish)
End If


If Mid$(FrmDecGM.Caption, 34, 12) = "Hardcore    " Then
    Combo0.Visible = True
    Combo1.Visible = True
    Combo2.Visible = True
    Combo3.Visible = True
    Combo4.Visible = True
    Combo5.Visible = True
    Combo6.Visible = True
    Combo7.Visible = True
    clockorder = 1
    mintcount = 100
    Cls
    
    Call DtBwholenoconv(decval, finalbinval)
    Call dtbhardmode(decval)
End If


tmrgame.Enabled = True

queno = 0
correctanscount = 0

cmdsavehighscore.Visible = False
CmdShowhiscore.Visible = False


End Sub

Private Sub Cmdselectmode_Click()

FrmSelection.Show
FrmDecGM.Hide

End Sub

Private Sub DtBwholenoconv(ByRef decval As Integer, ByRef finalbinval As String)

Randomize
decval = Int(256 * Rnd)

Do      'Initialize testing condition
    binarymut = 0       'Reset combined values of all binary-powers
    For countertemp = 1 To countermut       'Combine decimal-equivalent of all binary-powers up to the number of *countermut*
        binarymut = binarymut + (2 ^ (countertemp - 1))
    Next
    If binarymut >= Val(decval) Then
        exceedmut = 1
    Else
        countermut = countermut + 1
    End If
Loop Until exceedmut = 1

dummyval = decval

For Counter = 1 To countermut
    truebin = 0
    If (2 ^ (countermut - Counter)) <= Val(dummyval) Then
        truebin = 1
        dummyval = Val(dummyval) - (2 ^ (countermut - Counter))
    End If
    If truebin = 1 Then
        finalbinval = finalbinval & "1"
    Else
        finalbinval = finalbinval & "0"
    End If
Next

If finalbinval = "" Then finalbinval = "0"

End Sub

Private Sub dtbeasymode(ByVal decval As Integer, ByVal finalbinval As String, ByRef correctans As Integer, ByRef multchans() As String, ByRef queno As Integer, ByRef gamefinish As String)

Set anslabel(1) = Lbloption1
Set anslabel(2) = Lbloption2
Set anslabel(3) = Lbloption3
Set anslabel(4) = Lbloption4

correctans = Int((4 * Rnd) + 1)
If queno <> 10 Then
    queno = queno + 1
Else
    gamefinish = 1
End If

multchans(correctans) = finalbinval

LblDecTitle.Caption = decval

For Counter = 1 To 4
    If Counter <> correctans Then
        Do
            For Counter2 = 1 To 8
                Randomize
                multchans(Counter) = multchans(Counter) & Int(2 * Rnd)
            Next
        Loop Until multchans(Counter) <> multchans(correctans)
    End If
    anslabel(Counter).Caption = multchans(Counter)
Next

End Sub

Private Sub dtbnormalmode(ByVal decval As String, ByRef queno As Integer, ByRef gamefinish As String)

If queno <> 10 Then
    queno = queno + 1
Else
    gamefinish = 1
End If

LblDecTitle.Caption = decval

End Sub

Private Sub dtbhardmode(ByVal decval As String)

LblDecTitle.Caption = decval

End Sub

Private Sub tmrgame_Timer()

If clockorder = 0 Then
    mintcount = mintcount + 1
    Lbltime.Caption = "Time Taken: " & mintcount
End If

If clockorder = 1 Then
    mintcount = mintcount - 1
    Lbltime.Caption = "Time Left: " & mintcount
    
    If mintcount = 0 Then
        MsgBox ("Game Finish! You have answered " & correctanscount & " question(s) correctly.")
        Frame1.Visible = False
        tmrgame.Enabled = False
        CmdShowhiscore.Visible = True
        cmdsavehighscore.Visible = True
    End If
    
End If

End Sub

