VERSION 5.00
Begin VB.Form FrmConvertMode 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Convert Mode"
   ClientHeight    =   10785
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11340
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10785
   ScaleWidth      =   11340
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbdecneg 
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
      ItemData        =   "FrmConvertMode.frx":0000
      Left            =   2640
      List            =   "FrmConvertMode.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1920
      Width           =   495
   End
   Begin VB.TextBox TxtResult 
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
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   3360
      Width           =   9975
   End
   Begin VB.CommandButton Cmdshowstep 
      Caption         =   "Show Steps"
      Height          =   615
      Left            =   9120
      TabIndex        =   8
      Top             =   2160
      Visible         =   0   'False
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
      TabIndex        =   9
      Top             =   480
      Width           =   1575
   End
   Begin VB.CommandButton CmdconvertBtD 
      Caption         =   "Convert"
      Height          =   495
      Left            =   6600
      TabIndex        =   6
      Top             =   2640
      Width           =   1815
   End
   Begin VB.CommandButton CmdconvertDtB 
      Caption         =   "Convert"
      Height          =   495
      Left            =   6600
      TabIndex        =   3
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Frame FrmTutorial 
      Height          =   6855
      Left            =   120
      TabIndex        =   12
      Top             =   3840
      Width           =   11175
   End
   Begin VB.ComboBox cmbbinneg 
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
      ItemData        =   "FrmConvertMode.frx":0014
      Left            =   2640
      List            =   "FrmConvertMode.frx":001E
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2640
      Width           =   495
   End
   Begin VB.TextBox TxtbinTdec 
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
      Left            =   3240
      MaxLength       =   23
      TabIndex        =   4
      Top             =   2640
      Width           =   3255
   End
   Begin VB.TextBox TxtdecTbin 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   1
      EndProperty
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
      Left            =   3240
      MaxLength       =   23
      TabIndex        =   1
      Top             =   1920
      Width           =   3255
   End
   Begin VB.Label Lbldescription3 
      Caption         =   "Result is displayed at bottem textbox and can be copied."
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
      TabIndex        =   15
      Top             =   1320
      Width           =   8895
   End
   Begin VB.Label Lbldescription2 
      Caption         =   $"FrmConvertMode.frx":0028
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   14
      Top             =   720
      Width           =   8895
   End
   Begin VB.Label lblResult 
      Alignment       =   2  'Center
      Caption         =   "Result"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   3360
      Width           =   855
   End
   Begin VB.Label LblDecimalsec 
      Alignment       =   2  'Center
      Caption         =   "Binary to Decimal"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   11
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Label LblBinarysec 
      Alignment       =   2  'Center
      Caption         =   "Decimal to Binary"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   10
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label Lbldescription1 
      Caption         =   $"FrmConvertMode.frx":00B2
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   0
      Top             =   240
      Width           =   8895
   End
End
Attribute VB_Name = "FrmConvertMode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim binval As String
Dim binfloatpos As Integer
Dim totaldecval As Double

Dim decval As String
Dim INTdecval As Double
Dim decfloatpos As Integer
Dim finalbinval As String

Private Sub Cmdselectmode_Click()

FrmSelection.Show
FrmConvertMode.Hide

End Sub

Private Sub Form_Load()

cmbdecneg.ListIndex = 0
cmbbinneg.ListIndex = 0

End Sub

Private Sub TxtdecTbin_KeyPress(KeyAscii As Integer)

dotconfirm = 0
If InStr("0123456789.", Chr(KeyAscii)) = 0 And KeyAscii <> 8 And KeyAscii <> 127 Then
    KeyAscii = 0
Else
    For dotcounter = 1 To Len(TxtdecTbin.Text)
        If Asc(Mid$(TxtdecTbin.Text, dotcounter, 1)) = 46 And dotconfirm < 1 Then dotconfirm = dotconfirm + 1
        If KeyAscii = 46 And dotconfirm = 1 Then KeyAscii = 0
    Next
End If

End Sub

Private Sub TxtbinTdec_KeyPress(KeyAscii As Integer)

dotconfirm = 0
If InStr("10.", Chr(KeyAscii)) = 0 And KeyAscii <> 8 And KeyAscii <> 127 Then
    KeyAscii = 0
Else
    For dotcounter = 1 To Len(TxtbinTdec.Text)
        If Asc(Mid$(TxtbinTdec.Text, dotcounter, 1)) = 46 And dotconfirm < 1 Then dotconfirm = dotconfirm + 1
        If KeyAscii = 46 And dotconfirm = 1 Then KeyAscii = 0
    Next
End If

End Sub

Private Sub CmdconvertDtB_Click()

decval = TxtdecTbin.Text
finalbinval = ""

For Counter = 1 To Len(decval)
    If Mid$(decval, Counter, 1) = "." Then
        decfloatpos = Counter
        Call DtBfloatpointconv(decval, INTdecval, decfloatpos, finalbinval)
        If cmbdecneg.ListIndex = 1 Then finalbinval = "1" & finalbinval
        TxtResult.Text = finalbinval
        Exit Sub
    End If
Next
Call DtBwholenoconv(decval, finalbinval)
If cmbdecneg.ListIndex = 1 Then finalbinval = "1" & finalbinval
TxtResult.Text = finalbinval

End Sub

Private Sub DtBwholenoconv(ByVal decval As String, ByRef finalbinval As String)

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

For Counter = 1 To countermut
    truebin = 0
    If (2 ^ (countermut - Counter)) <= Val(decval) Then
        truebin = 1
        decval = Val(decval) - (2 ^ (countermut - Counter))
    End If
    If truebin = 1 Then
        finalbinval = finalbinval & "1"
    Else
        finalbinval = finalbinval & "0"
    End If
Next

If finalbinval = "" Then finalbinval = "0"

End Sub

Private Sub DtBfloatpointconv(ByVal decval As String, ByVal decfloatpos As Integer, ByVal INTdecval As Double, ByRef finalbinval As String)

INTdecval = Round(decval)

Do      'Initialize testing condition
    binarymut = 0       'Reset combined values of all binary-powers
    For countertemp = 1 To countermut       'Combine decimal-equivalent of all binary-powers up to the number of *countermut*
        binarymut = binarymut + (2 ^ (countertemp - 1))
    Next
    If binarymut >= INTdecval Then
        exceedmut = 1
    Else
        countermut = countermut + 1
    End If
Loop Until exceedmut = 1

For Counter = 1 To countermut
    truebin = 0
    If (2 ^ (countermut - Counter)) <= Val(decval) Then
        truebin = 1
        decval = Val(decval) - (2 ^ (countermut - Counter))
    End If
    If truebin = 1 Then
        finalbinval = finalbinval & "1"
    Else
        finalbinval = finalbinval & "0"
    End If
Next

If finalbinval = "" Then finalbinval = "0"
finalbinval = finalbinval & "."
Counter = 1

Do
    truebin = 0
    If (2 ^ -(Counter)) <= Val(decval) Then
        truebin = 1
        decval = Val(decval) - (2 ^ -(Counter))
    End If
    If truebin = 1 Then
        finalbinval = finalbinval & "1"
    Else
        finalbinval = finalbinval & "0"
    End If
    Counter = Counter + 1
Loop Until decval = 0

End Sub

Private Sub CmdconvertBtD_Click()

binval = TxtbinTdec.Text
totaldecval = 0

For Counter = 1 To Len(binval)
    If Mid$(binval, Counter, 1) = "." Then
        binfloatpos = Counter
        Call BtDfloatpointconv(binval, binfloatpos, totaldecval)
        finalvalstr = totaldecval
        If cmbbinneg.ListIndex = 1 Then finalvalstr = "-" & totaldecval
        TxtResult.Text = finalvalstr
        Exit Sub
    End If
Next
Call BtDwholenoconv(binval, totaldecval)
finalvalstr = totaldecval
If cmbbinneg.ListIndex = 1 Then finalvalstr = "-" & totaldecval
TxtResult.Text = finalvalstr

End Sub

Private Sub BtDfloatpointconv(ByVal binval As String, ByVal binfloatpos, ByRef totaldecval As Double)

For Counter = 1 To Len(binval)
    If binfractpoint <> 1 Then
        If Mid$(binval, Counter, 1) = "1" Then
            totaldecval = totaldecval + 2 ^ (binfloatpos - Counter - 1)
        End If
    End If
    If Mid$(binval, Counter, 1) = "." Then binfractpoint = 1
    If binfractpoint = 1 Then
        If Mid$(binval, Counter, 1) = "1" Then
        totaldecval = totaldecval + 2 ^ -(Counter - binfloatpos)
        End If
    End If
Next

End Sub

Private Sub BtDwholenoconv(ByRef binval As String, ByRef totaldecval As Double)

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
