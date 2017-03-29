VERSION 5.00
Begin VB.Form FrmBinGM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Game Mode 2 - Binary to Decimal"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   9045
   StartUpPosition =   3  'Windows Default
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
      TabIndex        =   8
      Top             =   840
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Timer tmrgame 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6480
      Top             =   1680
   End
   Begin VB.Frame Frame1 
      Height          =   2055
      Left            =   120
      TabIndex        =   21
      Top             =   2640
      Visible         =   0   'False
      Width           =   8775
      Begin VB.OptionButton Option4 
         Caption         =   "Option1"
         Height          =   315
         Left            =   4560
         TabIndex        =   3
         Top             =   720
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Option1"
         Height          =   255
         Left            =   1920
         TabIndex        =   2
         Top             =   720
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Option1"
         Height          =   195
         Left            =   4560
         TabIndex        =   1
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
         Left            =   3000
         TabIndex        =   4
         Top             =   1080
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3000
         MaxLength       =   3
         TabIndex        =   22
         Top             =   360
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   255
         Left            =   1920
         TabIndex        =   0
         Top             =   360
         Visible         =   0   'False
         Width           =   255
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
         TabIndex        =   27
         Top             =   720
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
         TabIndex        =   26
         Top             =   720
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
         TabIndex        =   24
         Top             =   360
         Visible         =   0   'False
         Width           =   1335
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
      TabIndex        =   7
      Top             =   1440
      Width           =   1695
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
      TabIndex        =   9
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton Cmdselectmode 
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
      TabIndex        =   6
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Lbltime 
      Height          =   255
      Left            =   7440
      TabIndex        =   23
      Top             =   240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label7 
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
      Height          =   495
      Left            =   5880
      TabIndex        =   20
      Top             =   1920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label6 
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
      Height          =   495
      Left            =   5400
      TabIndex        =   19
      Top             =   1920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label5 
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
      Height          =   495
      Left            =   4920
      TabIndex        =   18
      Top             =   1920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label4 
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
      Height          =   495
      Left            =   4440
      TabIndex        =   17
      Top             =   1920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label3 
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
      Height          =   495
      Left            =   3960
      TabIndex        =   16
      Top             =   1920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label2 
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
      Height          =   495
      Left            =   3480
      TabIndex        =   15
      Top             =   1920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label1 
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
      Height          =   495
      Left            =   3000
      TabIndex        =   14
      Top             =   1920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblbinaryaidtable 
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
      TabIndex        =   13
      Top             =   1440
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.Label Label0 
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
      Height          =   495
      Left            =   2520
      TabIndex        =   12
      Top             =   1920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblbindescnormal 
      Caption         =   $"FrmBinGMEasy.frx":0000
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
      Left            =   1920
      TabIndex        =   11
      Top             =   240
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.Label lblbindeschard 
      Caption         =   $"FrmBinGMEasy.frx":008D
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
      Left            =   1920
      TabIndex        =   10
      Top             =   240
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.Label lblbindesceasy 
      Caption         =   $"FrmBinGMEasy.frx":011A
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
      Left            =   1920
      TabIndex        =   5
      Top             =   240
      Visible         =   0   'False
      Width           =   5055
   End
End
Attribute VB_Name = "FrmBinGM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mintcount As Integer
Dim clockorder As String
Dim multchans(4) As Integer
Dim queno As Integer
Dim choiceque(4) As OptionButton
Dim quelabel(8) As Label
Dim anslabel(4) As Label
Dim gamefinish As String

Dim binval As String
Dim totaldecval As Integer
Dim correctans As Integer
Dim correctanscount As Integer

Dim scorelist(31) As String
Dim scoreval(10, 2) As Integer

Private Sub CmdConfirm_Click()

If Mid$(FrmBinGM.Caption, 34, 12) = "Beginner    " Then
    
    totaldecval = 0
    
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
    
    Call BtDwholenoconv(binval, totaldecval)
    Call btdeasymode(binval, totaldecval, correctans, multchans(), queno, gamefinish)
    
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


If Mid$(FrmBinGM.Caption, 34, 12) = "Intermediate" Then

    If Text1.Text <> "" Then
        If Text1.Text = totaldecval Then trueval = 1
    End If
    
    If trueval <> 1 Then
        MsgBox ("Wrong answer! The correct answer was ") & totaldecval & "."
    Else
        correctanscount = correctanscount + 1
    End If
    
    totaldecval = 0
    
    Call BtDwholenoconv(binval, totaldecval)
    Call btdnormalmode(binval, queno, gamefinish)
    
    Text1.Text = ""
    
    If gamefinish = 1 Then
        MsgBox ("Game Finish! You have answered " & correctanscount & " question(s) correctly. Your time was " & mintcount & " seconds!")
        Frame1.Visible = False
        tmrgame.Enabled = False
        CmdShowhiscore.Visible = True
        cmdsavehighscore.Visible = True
    End If
    
End If


If Mid$(FrmBinGM.Caption, 34, 12) = "Hardcore    " Then
    
    If Text1.Text <> "" Then
        If Text1.Text = totaldecval Then trueval = 1
    End If
    
    If trueval <> 1 Then
        MsgBox ("Wrong answer! The correct answer was ") & totaldecval & "."
        totaldecval = 0
        Call BtDwholenoconv(binval, totaldecval)
        Call btdhardmode(binval)
    Else
        correctanscount = correctanscount + 1
        totaldecval = 0
        Call BtDwholenoconv(binval, totaldecval)
        Call btdhardmode(binval)
    End If
    
    Text1.Text = ""
    
End If

End Sub

Private Sub Cmdgamestart_Click()

Label0.Visible = True
Label1.Visible = True
Label2.Visible = True
Label3.Visible = True
Label4.Visible = True
Label5.Visible = True
Label6.Visible = True
Label7.Visible = True

Lbltime.Visible = True
Frame1.Visible = True
totaldecval = 0

If Mid$(FrmBinGM.Caption, 34, 12) = "Beginner    " Then
    Option1.Visible = True
    Option2.Visible = True
    Option3.Visible = True
    Option4.Visible = True
    Lbloption1.Visible = True
    Lbloption2.Visible = True
    Lbloption3.Visible = True
    Lbloption4.Visible = True
    lblbinaryaidtable.Visible = True
    gamefinish = 0
    clockorder = 0
    mintcount = 0
    Cls
    
    Call BtDwholenoconv(binval, totaldecval)
    Call btdeasymode(binval, totaldecval, correctans, multchans(), queno, gamefinish)
End If


If Mid$(FrmBinGM.Caption, 34, 12) = "Intermediate" Then
    Text1.Visible = True
    lblbinaryaidtable.Visible = True
    gamefinish = 0
    clockorder = 0
    mintcount = 0
    Cls
    
    Call BtDwholenoconv(binval, totaldecval)
    Call btdnormalmode(binval, queno, gamefinish)
End If


If Mid$(FrmBinGM.Caption, 34, 12) = "Hardcore    " Then
    Text1.Visible = True
    clockorder = 1
    mintcount = 100
    Cls
    
    Call BtDwholenoconv(binval, totaldecval)
    Call btdhardmode(binval)
End If


tmrgame.Enabled = True

queno = 0
correctanscount = 0

cmdsavehighscore.Visible = False
CmdShowhiscore.Visible = False

End Sub

Private Sub cmdsavehighscore_Click()

Counter = 0
    
If Mid$(FrmBinGM.Caption, 34, 12) = "Beginner    " Then
    
    scorelist(30) = InputBox("Please enter your name.", "Enter name for high-score")
    If scorelist(30) = "" Then
        Exit Sub
    End If

    If Dir("N:\IT Misc\BtDGameRecord - Easy.txt") <> "" Then
        Open "N:\IT Misc\BtDGameRecord - Easy.txt" For Input As #1     'Create a new file with the concatenation result as name in slot 1
            Do Until EOF(1)
                Input #1, Data
                scorelist(Counter) = Data
                Counter = Counter + 1
                EOF (1)
            Loop
        Close #1
        For Counter = 0 To 9
            scoreval(Counter, 1) = Val(scorelist(Counter))
            scoreval(Counter, 2) = Val(scorelist(Counter + 10))
        Next
    
        scoreval(10, 1) = correctanscount
        scoreval(10, 2) = mintcount
        
        Call scoresortop(scoreval(), scorelist())
        Call timesortop(scoreval(), scorelist())

        Open "N:\IT Misc\BtDGameRecord - Easy.txt" For Output As #1
            For Counter = 0 To 9
                Print #1, scoreval(Counter, 1)
            Next
            For Counter = 0 To 9
                Print #1, scoreval(Counter, 2)
            Next
            For Counter = 20 To 29
                Print #1, scorelist(Counter)
            Next
        Close #1
    End If
    
End If


If Mid$(FrmBinGM.Caption, 34, 12) = "Intermediate" Then
    
    scorelist(30) = InputBox("Please enter your name.", "Enter name for high-score")
    If scorelist(30) = "" Then
        Exit Sub
    End If
    
    If Dir("N:\IT Misc\BtDGameRecord - Normal.txt") <> "" Then
        Open "N:\IT Misc\BtDGameRecord - Normal.txt" For Input As #1     'Create a new file with the concatenation result as name in slot 1
            Do Until EOF(1)
                Input #1, Data
                scorelist(Counter) = Data
                Counter = Counter + 1
                EOF (1)
            Loop
        Close #1
        For Counter = 0 To 9
            scoreval(Counter, 1) = Val(scorelist(Counter))
            scoreval(Counter, 2) = Val(scorelist(Counter + 10))
        Next
    
        scoreval(10, 1) = correctanscount
        scoreval(10, 2) = mintcount
        
        Call scoresortop(scoreval(), scorelist())
        Call timesortop(scoreval(), scorelist())

        Open "N:\IT Misc\BtDGameRecord - Normal.txt" For Output As #1
            For Counter = 0 To 9
                Print #1, scoreval(Counter, 1)
            Next
            For Counter = 0 To 9
                Print #1, scoreval(Counter, 2)
            Next
            For Counter = 20 To 29
                Print #1, scorelist(Counter)
            Next
        Close #1
    End If
    
End If


If Mid$(FrmBinGM.Caption, 34, 12) = "Hardcore    " Then

    scorelist(20) = InputBox("Please enter your name.", "Enter name for high-score")
    If scorelist(20) = "" Then
        Exit Sub
    End If

    If Dir("N:\IT Misc\BtDGameRecord - Hard.txt") <> "" Then
        Open "N:\IT Misc\BtDGameRecord - Hard.txt" For Input As #1     'Create a new file with the concatenation result as name in slot 1
            Do Until EOF(1)
                Input #1, Data
                scorelist(Counter) = Data
                Counter = Counter + 1
                EOF (1)
            Loop
        Close #1
        
        For Counter = 0 To 9
            scoreval(Counter, 1) = Val(scorelist(Counter))
        Next
    
        scoreval(10, 1) = correctanscount
        
        For Outer = 0 To 10     'The 1st Number being compared in a Simple Sort
            For Inner = Outer + 0 To 10     'The 2nd Number being compared in a Simple Sort
                If scoreval(Outer, 1) < scoreval(Inner, 1) Then      'i.e. If the 1st Number is greater than the 2nd number
                    Tempval = scoreval(Outer, 1)     'Temporary value required for the 'swap' operation, it contain the 1st value
                    tempval2 = scorelist(Outer + 10)
                    scoreval(Outer, 1) = scoreval(Inner, 1)     'Transfering 2nd position to 1st position
                    scorelist(Outer + 10) = scorelist(Inner + 10)
                    scoreval(Inner, 1) = Tempval      'Place temporary value to the 2nd position
                    scorelist(Inner + 10) = tempval2
                End If
            Next Inner
        Next Outer

        Open "N:\IT Misc\BtDGameRecord - Hard.txt" For Output As #1
            For Counter = 0 To 9
                Print #1, scoreval(Counter, 1)
            Next
            For Counter = 10 To 19
                Print #1, scorelist(Counter)
            Next
        Close #1
    End If
    
End If
    

cmdsavehighscore.Visible = False
    
End Sub

Private Sub Cmdselectmode_Click()

FrmSelection.Show
Unload FrmBinGM

End Sub

Private Sub CmdShowhiscore_Click()

FrmHighscore.Show

End Sub

Private Sub Option1_KeyPress(keyascii As Integer)

If keyascii = 13 Then
    Call CmdConfirm_Click
End If

End Sub

Private Sub Option2_KeyPress(keyascii As Integer)

If keyascii = 13 Then
    Call CmdConfirm_Click
End If

End Sub

Private Sub Option3_KeyPress(keyascii As Integer)

If keyascii = 13 Then
    Call CmdConfirm_Click
End If

End Sub

Private Sub Option4_KeyPress(keyascii As Integer)

If keyascii = 13 Then
    Call CmdConfirm_Click
End If

End Sub

Private Sub Text1_KeyPress(keyascii As Integer)

dotconfirm = 0
If InStr("0123456789", Chr(keyascii)) = 0 And keyascii <> 8 And keyascii <> 127 And keyascii <> 13 Then
    keyascii = 0
End If

If keyascii = 13 Then
    Call CmdConfirm_Click
End If

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

Private Sub BtDwholenoconv(ByRef binval As String, ByRef totaldecval As Integer)
binval = ""

For Counter = 1 To 8
    Randomize
    binval = binval & Int(2 * Rnd)
Next

For Counter = 1 To Len(binval)
If lastbin <> 1 And Len(binval) <> 1 Then
    If Mid$(binval, Len(binval) - Counter, 1) = "1" Then
        totaldecval = totaldecval + 2 ^ (Counter)
    End If
End If
If Counter = Len(binval) - 1 Then lastbin = 1
If lastbin = 1 And Mid$(binval, Len(binval), 1) = "1" And Counter = Len(binval) Then totaldecval = totaldecval + 1
Next

If Len(binval) = 1 Then totaldecval = 1

End Sub

Private Sub scoresortop(ByRef scoreval() As Integer, ByRef scorelist() As String)

For Outer = 0 To 10     'The 1st Number being compared in a Simple Sort
    For Inner = Outer + 0 To 10     'The 2nd Number being compared in a Simple Sort
        If scoreval(Outer, 1) < scoreval(Inner, 1) Then      'i.e. If the 1st Number is greater than the 2nd number
            Tempval = scoreval(Outer, 1)     'Temporary value required for the 'swap' operation, it contain the 1st value
            tempval2 = scoreval(Outer, 2)
            tempval3 = scorelist(Outer + 20)
            scoreval(Outer, 1) = scoreval(Inner, 1)     'Transfering 2nd position to 1st position
            scoreval(Outer, 2) = scoreval(Inner, 2)
            scorelist(Outer + 20) = scorelist(Inner + 20)
            scoreval(Inner, 1) = Tempval      'Place temporary value to the 2nd position
            scoreval(Inner, 2) = tempval2
            scorelist(Inner + 20) = tempval3
        End If
    Next Inner
Next Outer

End Sub

Private Sub timesortop(ByRef scoreval() As Integer, ByRef scorelist() As String)

Dim samescoval(10) As Integer

nocounter = 0

For Outer = 0 To 9     'The 1st Number being compared in a Simple Sort
    For Inner = Outer + 1 To 9     'The 2nd Number being compared in a Simple Sort
        If scoreval(Outer, 1) = scoreval(Inner, 1) Then      'i.e. If the 1st Number is greater than the 2nd number
            samescoval(Counter) = Inner
            nocounter = nocounter + 1
        End If
        If scoreval(Inner, 1) < scoreval(Outer, 1) Then
            Inner = 9
        End If
    Next Inner
    For Counter = 1 To nocounter     'The 2nd Number being compared in a Simple Sort
        If scoreval(samescoval(Counter), 2) < scoreval(Outer, 2) Then     'i.e. If the 1st Number is greater than the 2nd number
            Tempval = scoreval(samescoval(Counter), 2)      'Temporary value required for the 'swap' operation, it contain the 1st value
            tempval2 = scorelist(samescoval(Counter) + 20)
            scoreval(samescoval(Counter), 2) = scoreval(Outer, 2)      'Transfering 2nd position to 1st position
            scorelist(samescoval(Counter) + 20) = scorelist(Outer + 20)
            scoreval(Outer, 2) = Tempval       'Place temporary value to the 2nd position
            scorelist(Outer + 20) = tempval2
        End If
    Next
    nocounter = 0
Next Outer



End Sub

Private Sub btdeasymode(ByVal binval As String, ByVal totaldecval As Integer, ByRef correctans As Integer, ByRef multchans() As Integer, ByRef queno As Integer, ByRef gamefinish As String)

Set quelabel(1) = Label0
Set quelabel(2) = Label1
Set quelabel(3) = Label2
Set quelabel(4) = Label3
Set quelabel(5) = Label4
Set quelabel(6) = Label5
Set quelabel(7) = Label6
Set quelabel(8) = Label7

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

multchans(correctans) = totaldecval

For Counter = 1 To 8
    quelabel(Counter) = Mid$(binval, Counter, 1)
Next

For Counter = 1 To 4
    If Counter <> correctans Then
        Do
            multchans(Counter) = Int(256 * Rnd)
        Loop Until multchans(Counter) <> multchans(correctans)
    End If
    anslabel(Counter).Caption = multchans(Counter)
Next

End Sub

Private Sub btdnormalmode(ByVal binval As String, ByRef queno As Integer, ByRef gamefinish As String)

Set quelabel(1) = Label0
Set quelabel(2) = Label1
Set quelabel(3) = Label2
Set quelabel(4) = Label3
Set quelabel(5) = Label4
Set quelabel(6) = Label5
Set quelabel(7) = Label6
Set quelabel(8) = Label7

If queno <> 10 Then
    queno = queno + 1
Else
    gamefinish = 1
End If

For Counter = 1 To 8
    quelabel(Counter) = Mid$(binval, Counter, 1)
Next

End Sub

Private Sub btdhardmode(ByVal binval As String)

Set quelabel(1) = Label0
Set quelabel(2) = Label1
Set quelabel(3) = Label2
Set quelabel(4) = Label3
Set quelabel(5) = Label4
Set quelabel(6) = Label5
Set quelabel(7) = Label6
Set quelabel(8) = Label7

For Counter = 1 To 8
    quelabel(Counter) = Mid$(binval, Counter, 1)
Next

End Sub
