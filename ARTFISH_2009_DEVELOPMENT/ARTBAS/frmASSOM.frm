VERSION 5.00
Begin VB.Form frmASSOM 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11100
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7320
   ScaleMode       =   0  'User
   ScaleWidth      =   13284.19
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdFINISH 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   5.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      MousePointer    =   1  'Arrow
      Picture         =   "frmASSOM.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6360
      Width           =   735
   End
   Begin VB.CommandButton cmdPRINT 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   960
      MousePointer    =   1  'Arrow
      Picture         =   "frmASSOM.frx":1DF2
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6360
      Width           =   735
   End
   Begin VB.ListBox lstRESULTS 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5580
      Left            =   5880
      TabIndex        =   4
      Top             =   600
      Width           =   5055
   End
   Begin VB.ListBox lstSITES 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2700
      Left            =   120
      MultiSelect     =   1  'Simple
      TabIndex        =   3
      Top             =   3480
      Width           =   5655
   End
   Begin VB.ListBox lstSTRATA 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1740
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   5655
   End
   Begin VB.CommandButton cmdASSOCIATE 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   4320
      MousePointer    =   1  'Arrow
      Picture         =   "frmASSOM.frx":2074
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton cmdDEL 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   1800
      Picture         =   "frmASSOM.frx":22F6
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6360
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FFFF&
      Caption         =   " 04"
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   7080
      Width           =   255
   End
   Begin VB.Label lblRESULTS 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "???"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   7
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label lblSITES 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "???"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   3120
      Width           =   2295
   End
   Begin VB.Label lblSTRATA 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "???"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   2415
   End
End
Attribute VB_Name = "frmASSOM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private MNAME(), MNCODE()
Private SINAME(), SICODE()
Private NSIT, NMN, APPFLAG

Private NSIT_SEL, NMN_SEL As Integer

Private MNAME_SEL(), MNCODE_SEL()
Private MNSEL(), SISEL()
Private kres, kold As Integer
Private aseq As Integer
Private SINAME_SEL(), SICODE_SEL()
Private Sub cmdASSOCIATE_Click()

Dim response As String
Dim I, ii, J, jj As Integer

If lstSTRATA.SelCount = 0 Or lstSITES.SelCount = 0 Then
    Beep
    response = MsgBox(msgtab(76), vbExclamation, " ")
    Exit Sub
    End If

NSIT_SEL = lstSITES.SelCount: NMN_SEL = lstSTRATA.SelCount

ReDim mnseq_sel(1 To NMN_SEL), MNAME_SEL(1 To NMN_SEL)
ReDim MNCODE_SEL(1 To NMN_SEL), mnsort_sel(1 To NMN_SEL)
ReDim siseq_sel(1 To NSIT_SEL), SINAME_SEL(1 To NSIT_SEL)
ReDim SICODE_SEL(1 To NSIT_SEL), sisort_sel(1 To NSIT_SEL)

J = 0

For I = 0 To lstSTRATA.ListCount - 1

If lstSTRATA.Selected(I) = True Then

   ii = lstSTRATA.ItemData(I)
   MNSEL(ii) = 1
   J = J + 1
   MNAME_SEL(J) = MNAME(ii)
   MNCODE_SEL(J) = MNCODE(ii)
   
   End If

Next I

lstSTRATA.ListIndex = 0

J = 0

For I = 0 To lstSITES.ListCount - 1

If lstSITES.Selected(I) = True Then

   ii = lstSITES.ItemData(I)
   SISEL(ii) = 1
   J = J + 1
   SINAME_SEL(J) = SINAME(ii)
   SICODE_SEL(J) = SICODE(ii)
   
   End If

Next I

lstSITES.ListIndex = 0

kold = kres

For I = 1 To NMN_SEL

kres = kres + 1
lstRESULTS.AddItem MNAME_SEL(I)

For J = 1 To NSIT_SEL
kres = kres + 1
lstRESULTS.AddItem Space(5) + SINAME_SEL(J)
Next J
Next I

lstRESULTS.ListIndex = kold

Call UPDATE_ASSOC
Call RELOAD_STRATA
Call RELOAD_SITES

If lstSTRATA.ListCount + lstSITES.ListCount = 0 Then
   cmdFINISH.Visible = True
   cmdPRINT.Visible = True
   cmdASSOCIATE.Visible = False
   End If
   
If lstSTRATA.ListCount + lstSITES.ListCount <> 0 _
   And lstSTRATA.ListCount * lstSITES.ListCount = 0 Then
   
   Beep
   
   cmdFINISH.Visible = False
   cmdPRINT.Visible = False
   cmdASSOCIATE.Visible = False
   cmdDEL.Visible = False
   
   response = MsgBox(msgtab(76), vbExclamation, " ")
   
   Close #2
   
   Call Form_Load
   
   End If
    
End Sub
Private Sub cmdDEL_Click()
Dim response As Integer

Beep

response = MsgBox(msgtab(36), vbCritical + vbDefaultButton2 + vbOKCancel, msgtab(65))

If response = 2 Then Exit Sub

Close #2

Dim fnm

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_ASSOMN.TXT"

If Dir(fnm) <> "" Then Kill fnm

Load frmTABLES
Unload frmASSOM
frmTABLES.Show

End Sub
Private Sub cmdFINISH_Click()

Close #2

frmASSOM.MousePointer = 13
Load frmTABLES
Unload frmASSOM
frmTABLES.Show

End Sub
Private Sub Form_Load()

Set Picture = LoadPicture(APPROOT + "\ARTBAS\PICS_RUNTIME\SCREEN_04.JPG")

Dim fnm

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_ASSOMN.TXT"

Open fnm For Output As #2

cmdASSOCIATE.Visible = True
cmdDEL.Visible = True
cmdFINISH.Visible = False
cmdPRINT.Visible = False

kres = 0: lstRESULTS.Clear: aseq = 0

frmASSOM.Caption = monthtab(current_month) + " " + _
                    Format(current_year, "0000") + " - " + msgtab(63)
                    
lblSTRATA.Caption = msgtab(41): lblSITES.Caption = msgtab(42)
lblRESULTS.Caption = msgtab(75)

cmdASSOCIATE.ToolTipText = msgtab(74)
cmdDEL.ToolTipText = msgtab(9)
cmdFINISH.ToolTipText = msgtab(51)
cmdPRINT.ToolTipText = msgtab(52)

Call LOAD_STRATA
Call LOAD_SITES

End Sub
Private Sub LOAD_STRATA()

Dim I, XXX, fnm

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_MAJOR.TXT"

NMN = 0

lstSTRATA.Clear

Open fnm For Input As #1

Do Until EOF(1)

Line Input #1, XXX

NMN = NMN + 1

ReDim Preserve MNAME(1 To NMN), MNCODE(1 To NMN), MNSEL(1 To NMN)

MNCODE(NMN) = Val(Mid(XXX, 1, 4))
MNAME(NMN) = Mid(XXX, 6, 30)

I = NMN

lstSTRATA.AddItem MNAME(I)
MNSEL(I) = 0
lstSTRATA.ItemData(lstSTRATA.NewIndex) = I

Loop

Close #1

lstSTRATA.ListIndex = 0


End Sub
Private Sub LOAD_SITES()

Dim I, XXX, fnm

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_MINOR.TXT"

NSIT = 0

lstSITES.Clear

Open fnm For Input As #1

Do Until EOF(1)

Line Input #1, XXX

NSIT = NSIT + 1

ReDim Preserve SINAME(1 To NSIT), SICODE(1 To NSIT), SISEL(1 To NSIT)

SICODE(NSIT) = Val(Mid(XXX, 1, 4))
SINAME(NSIT) = Mid(XXX, 6, 30)

I = NSIT

lstSITES.AddItem SINAME(I)
SISEL(I) = 0
lstSITES.ItemData(lstSITES.NewIndex) = I

Loop

Close #1

lstSITES.ListIndex = 0

End Sub
Private Sub UPDATE_ASSOC()

Dim I, J, XXX, fnm

For I = 1 To NMN_SEL

Print #2, Format(MNCODE_SEL(I), "0000") + " " _
                + Left(MNAME_SEL(I) + Space(30), 30) + " " _
                + Format(NSIT_SEL, "0000")

For J = 1 To NSIT_SEL

Print #2, Space(5) + Format(SICODE_SEL(J), "0000") + " " + _
         Left(SINAME_SEL(J) + Space(30), 30)

Next J
Next I

End Sub
Private Sub RELOAD_STRATA()

Dim I As Integer

lstSTRATA.Clear

For I = 1 To NMN

If MNSEL(I) = 1 Then GoTo loop1

lstSTRATA.AddItem MNAME(I)
lstSTRATA.ItemData(lstSTRATA.NewIndex) = I

loop1:

Next I

End Sub
Private Sub RELOAD_SITES()

Dim I As Integer

lstSITES.Clear

For I = 1 To NSIT

If SISEL(I) = 1 Then GoTo loop2

lstSITES.AddItem SINAME(I)
lstSITES.ItemData(lstSITES.NewIndex) = I

loop2:

Next I

End Sub
Private Sub cmdPRINT_Click()

Close #2

Dim fnm, XXX, na

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_ASSOMN.TXT"

Open fnm For Input As #2

Printer.FontBold = True
Printer.FontName = "Courier"
Printer.FontName = "Courier New"
Printer.FontSize = 11

Dim I, J, pageno, lineno

pageno = 0

GoSub CHANGE_PAGE

Do Until EOF(2)

Line Input #2, XXX

na = Val(Mid(XXX, 37, 4))

Printer.Print Tab(5); String(72, "-")

Printer.Print Tab(5); Left(XXX, 36)

Printer.Print

For J = 1 To na

Line Input #2, XXX

Printer.Print Tab(10); Left(XXX, 36)

lineno = lineno + 1

If lineno > 55 Then GoSub CHANGE_PAGE

Next J

Loop

Printer.EndDoc
Close #2

Exit Sub

'========================
CHANGE_PAGE:

lineno = 0
pageno = pageno + 1
If pageno > 1 Then Printer.NewPage

Printer.Print

Printer.Print Tab(5); frmASSOM.Caption

Printer.Print

Return
'====================================

End Sub

