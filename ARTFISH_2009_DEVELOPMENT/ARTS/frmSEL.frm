VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmSEL 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ARTSER"
   ClientHeight    =   7515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10815
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   3  'I-Beam
   Moveable        =   0   'False
   ScaleHeight     =   7515
   ScaleWidth      =   10815
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPRINT 
      BackColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   7920
      MousePointer    =   1  'Arrow
      Picture         =   "frmSEL.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   6480
      Width           =   855
   End
   Begin VB.CommandButton cmdGUIDE 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8880
      MousePointer    =   1  'Arrow
      Picture         =   "frmSEL.frx":0282
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6480
      Width           =   855
   End
   Begin VB.CommandButton cmdDEL 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   1080
      Picture         =   "frmSEL.frx":24E4
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6720
      Width           =   615
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   240
      Picture         =   "frmSEL.frx":2766
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6720
      Width           =   615
   End
   Begin VB.ListBox lstMINOR 
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H00800000&
      Height          =   2580
      Left            =   120
      MultiSelect     =   1  'Simple
      TabIndex        =   7
      Top             =   3600
      Width           =   5415
   End
   Begin VB.ListBox lstSPECIES 
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H00800000&
      Height          =   2580
      Left            =   5640
      MultiSelect     =   1  'Simple
      TabIndex        =   6
      Top             =   3600
      Width           =   5055
   End
   Begin VB.ListBox lstBG 
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H00800000&
      Height          =   2580
      Left            =   5640
      MultiSelect     =   1  'Simple
      TabIndex        =   5
      Top             =   360
      Width           =   5055
   End
   Begin VB.ListBox lstMAJOR 
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H00800000&
      Height          =   2580
      Left            =   120
      MultiSelect     =   1  'Simple
      TabIndex        =   4
      Top             =   360
      Width           =   5415
   End
   Begin VB.CommandButton cmdEXIT 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   10320
      MousePointer    =   1  'Arrow
      Picture         =   "frmSEL.frx":2870
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6960
      Width           =   375
   End
   Begin VB.CommandButton cmdRETURN 
      BackColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   9840
      MousePointer    =   1  'Arrow
      Picture         =   "frmSEL.frx":2AF2
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6480
      Width           =   855
   End
   Begin VB.CommandButton cmdFULL 
      BackColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   120
      Picture         =   "frmSEL.frx":2D74
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6480
      Width           =   855
   End
   Begin VB.CommandButton cmdSEL 
      BackColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   1080
      Picture         =   "frmSEL.frx":336E
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6480
      Width           =   855
   End
   Begin RichTextLib.RichTextBox rtsNOPRICE 
      Height          =   6255
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   11033
      _Version        =   393217
      BackColor       =   12648447
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      MousePointer    =   1
      Appearance      =   0
      TextRTF         =   $"frmSEL.frx":3968
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FFFF&
      Caption         =   "02"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   15
      Top             =   7320
      Width           =   255
   End
   Begin VB.Label lblSPECIES 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Species"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   5640
      TabIndex        =   11
      Top             =   3360
      Width           =   3975
   End
   Begin VB.Label lblMINOR 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Minor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   3360
      Width           =   3975
   End
   Begin VB.Label lblBG 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Boat/gear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   5640
      TabIndex        =   9
      Top             =   120
      Width           =   3975
   End
   Begin VB.Label lblMAJOR 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Major"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "frmSEL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private SELFLAG, FNM5, FNM6, ADDFISH, ADDVALUE

Private NMJ, NMN, NBG, NSP
Private TMJC(), TMJN(), TMNC(), TMNN(), TBGC(), TBGN(), TSPC(), TSPN(), TASSO()
Private SELMJ(), SELMN(), SELBG(), SELSP()
Private REPJ(), REPM(), REPB(), REPS()

Private Sub cmdDEL_click()

lstMAJOR.Visible = False
lstMINOR.Visible = False
lstBG.Visible = False
lstSPECIES.Visible = False

lblMAJOR.Visible = False
lblMINOR.Visible = False
lblBG.Visible = False
lblSPECIES.Visible = False

cmdOK.Visible = False
cmdDEL.Visible = False

cmdFULL.Visible = True
cmdSEL.Visible = True

End Sub
Private Sub cmdEXIT_Click()

Call write_parms

End

End Sub
Private Sub cmdFULL_Click()

cmdFULL.MousePointer = 13
frmSEL.MousePointer = 13

Dim dbn1, dbn2

dbn1 = APPROOT + "\ARTS\WORK\CT" + Format(CURY, "0000") + ".MDB"
dbn2 = APPROOT + "\ARTS\WORK\WGT" + Format(CURY, "0000") + ".MDB"

FileCopy dbn1, dbn2

dbn1 = APPROOT + "\ARTS\WORK\CS" + Format(CURY, "0000") + ".MDB"
dbn2 = APPROOT + "\ARTS\WORK\WGS" + Format(CURY, "0000") + ".MDB"

FileCopy dbn1, dbn2

dbn1 = APPROOT + "\ARTS\WORK\CT" + Format(CURY, "0000") + ".MDB"
dbn2 = APPROOT + "\ARTS\WORK\WT" + Format(CURY, "0000") + ".MDB"

FileCopy dbn1, dbn2

dbn1 = APPROOT + "\ARTS\WORK\CS" + Format(CURY, "0000") + ".MDB"
dbn2 = APPROOT + "\ARTS\WORK\WS" + Format(CURY, "0000") + ".MDB"

FileCopy dbn1, dbn2

cmdFULL.MousePointer = 1
frmSEL.MousePointer = 1

Open APPROOT + "\ARTS\WORK\WSEL.TXT" For Output As #1
Print #1, msgtab(43)
Print #1, String(40, "=")
Print #1, msgtab(44)

Close #1

Call CHECK_PRICES

Load frmCOMP
Unload frmSEL
frmCOMP.Show

End Sub
Private Sub cmdGUIDE_Click()

HTYPE = "20"

HFNM = APPROOT + "\ARTS\HELP\" + current_language + "HELP" + HTYPE + ".rtf"

If Dir(HFNM) = "" Then Exit Sub

frmSEL.Enabled = False
Load frmGUIDE
frmGUIDE.Show

End Sub
Private Sub cmdOK_Click()

cmdOK.MousePointer = 13

Call MULT_MAJOR
Call MULT_MINOR
Call MULT_BG
Call MULT_SPECIES

If lstMAJOR.SelCount = 0 And lstMINOR.SelCount = 0 And lstBG.SelCount = 0 And lstSPECIES.SelCount = 0 Then
   Call cmdDEL_click
   cmdOK.MousePointer = 1
   Exit Sub
   End If

Call FILTER_SPECIES
Call FILTER_TOT

cmdOK.MousePointer = 1

If NCURS = 0 Then

   Dim resp, wnm
   
   wnm = APPROOT + "\ARTS\WORK\WT" + Format(CURY, "0000") + ".MDB"
   If Dir(wnm) <> "" Then Kill wnm
   
   wnm = APPROOT + "\ARTS\WORK\WS" + Format(CURY, "0000") + ".MDB"
   If Dir(wnm) <> "" Then Kill wnm
   
   resp = MsgBox(msgtab(14), vbCritical + vbOKOnly, " ")
   Call cmdDEL_click
   Exit Sub
   End If

Call WRITE_SELECTED

Dim dbn1, dbn2

dbn1 = APPROOT + "\ARTS\WORK\WT" + Format(CURY, "0000") + ".MDB"
dbn2 = APPROOT + "\ARTS\WORK\WGT" + Format(CURY, "0000") + ".MDB"

If Dir(dbn1) <> "" Then FileCopy dbn1, dbn2

dbn1 = APPROOT + "\ARTS\WORK\WS" + Format(CURY, "0000") + ".MDB"
dbn2 = APPROOT + "\ARTS\WORK\WGS" + Format(CURY, "0000") + ".MDB"

If Dir(dbn1) <> "" Then FileCopy dbn1, dbn2

Call CHECK_PRICES

Load frmCOMP
Unload frmSEL
frmCOMP.Show

End Sub
Private Sub WRITE_SELECTED()

Open APPROOT + "\ARTS\WORK\WSEL.TXT" For Output As #1

Print #1, msgtab(31)
Print #1, String(40, "=")

Dim I, K, J

K = 0

For I = 1 To NMJ
J = TMJC(I)
If SELMJ(J) = "Y" Then K = K + 1
Next I

If K = NMJ Then

   Print #1, "100 %"
   Print #1, " "
   
   GoTo PRINT_MINOR
   End If

For I = 1 To NMJ
J = TMJC(I)
If SELMJ(J) <> "Y" Then GoTo NEXT_MAJOR
Print #1, TMJN(I)

NEXT_MAJOR:

Next I

Print #1, " "

PRINT_MINOR:

Print #1, msgtab(32)
Print #1, String(40, "=")

K = 0

For I = 1 To NMN
J = TMNC(I)
If SELMN(J) = "Y" Then K = K + 1
Next I

If K = NMN Then

   Print #1, "100 %"
   Print #1, " "
   
   GoTo PRINT_BG
   End If

For I = 1 To NMN
J = TMNC(I)
If SELMN(J) <> "Y" Then GoTo NEXT_MINOR
Print #1, TMNN(I)

NEXT_MINOR:

Next I

Print #1, " "

PRINT_BG:

K = 0

Print #1, msgtab(33)
Print #1, String(40, "=")

For I = 1 To NBG
J = TBGC(I)
If SELBG(J) = "Y" Then K = K + 1
Next I

If K = NBG Then

   Print #1, "100 %"
   Print #1, " "
   
   GoTo PRINT_SPECIES
   End If

For I = 1 To NBG
J = TBGC(I)
If SELBG(J) <> "Y" Then GoTo NEXT_BG
Print #1, TBGN(I)

NEXT_BG:

Next I

Print #1, " "

PRINT_SPECIES:

K = 0

Print #1, msgtab(34)
Print #1, String(40, "=")

For I = 1 To NSP
J = TSPC(I)
If SELSP(J) = "Y" Then K = K + 1
Next I

If K = NSP Then

   Print #1, "100 %"
   Print #1, " "
   
   Close #1
   Exit Sub
   End If

For I = 1 To NSP
J = TSPC(I)
If SELSP(J) <> "Y" Then GoTo NEXT_SP
Print #1, TSPN(I)

NEXT_SP:

Next I

Print #1, " "

Close #1

End Sub
Private Sub MULT_MAJOR()

ReDim SELMJ(1 To 10000)

Dim I

For I = 1 To NMJ
SELMJ(TMJC(I)) = "Y"
Next I

If lstMAJOR.SelCount = 0 Then Exit Sub

For I = 1 To NMJ
SELMJ(TMJC(I)) = "N"
Next I

For I = 0 To lstMAJOR.ListCount - 1

If lstMAJOR.Selected(I) = True Then
   SELMJ(TMJC(I + 1)) = "Y"
   End If

Next I

End Sub
Private Sub MULT_MINOR()

ReDim SELMN(1 To 10000)

Dim I, J, K

For I = 1 To NMN
SELMN(TMNC(I)) = "Y"
Next I

If lstMINOR.SelCount = 0 Then
   For I = 1 To NMN
   SELMN(TMNC(I)) = "N"
   Next I
   
   For I = 1 To NMN
   K = TASSO(I)
   If SELMJ(K) = "Y" Then SELMN(TMNC(I)) = "Y"
   Next I
   Exit Sub
   End If

For I = 1 To NMJ
J = TMJC(I): SELMJ(J) = "N"
Next I

For I = 1 To NMN
SELMN(TMNC(I)) = "N"
Next I

For I = 0 To lstMINOR.ListCount - 1

If lstMINOR.Selected(I) = True Then
   SELMN(TMNC(I + 1)) = "Y"
   J = TASSO(I + 1): SELMJ(J) = "Y"
   End If

Next I

End Sub
Private Sub MULT_BG()

ReDim SELBG(1 To 10000)

Dim I

For I = 1 To NBG
SELBG(TBGC(I)) = "Y"
Next I

If lstBG.SelCount = 0 Then Exit Sub

For I = 1 To NBG
SELBG(TBGC(I)) = "N"
Next I

For I = 0 To lstBG.ListCount - 1

If lstBG.Selected(I) = True Then
   SELBG(TBGC(I + 1)) = "Y"
   End If

Next I

End Sub
Private Sub MULT_SPECIES()

ReDim SELSP(1 To 10000)

Dim I

For I = 1 To NSP
SELSP(TSPC(I)) = "Y"
Next I

If lstSPECIES.SelCount = 0 Then Exit Sub

For I = 1 To NSP
SELSP(TSPC(I)) = "N"
Next I

For I = 0 To lstSPECIES.ListCount - 1

If lstSPECIES.Selected(I) = True Then
   SELSP(TSPC(I + 1)) = "Y"
   End If

Next I

End Sub

Private Sub cmdRETURN_Click()

Load frmARTS00
Unload frmSEL
frmARTS00.Show

End Sub
Private Sub cmdSEL_Click()

lstMAJOR.Visible = True
lstMINOR.Visible = True
lstBG.Visible = True
lstSPECIES.Visible = True

lblMAJOR.Visible = True
lblMINOR.Visible = True
lblBG.Visible = True
lblSPECIES.Visible = True

cmdOK.Visible = True
cmdDEL.Visible = True

cmdFULL.Visible = False
cmdSEL.Visible = False

SELFLAG = "N"

Call LOAD_MAJOR
Call LOAD_MINOR
Call LOAD_BG
Call LOAD_SPECIES

End Sub
Private Sub Form_Load()

cmdPRINT.Visible = False

COMPGEN = "NO"

rtsNOPRICE.Visible = False

Set Picture = LoadPicture(APPROOT + "\ARTS\PICS_RUNTIME\SCREEN_02.JPG")

Open APPROOT + "\ARTS\CONTROL\COMPUTE.TXT" For Output As #1
Print #1, "NOCOMPUTE"
Close #1

Dim fnm, XXX, J

fnm = APPROOT + "\ARTS\TABLES\Y" + Format(CURY, "0000") + "_MONTHS.TXT"

If Dir(fnm) = "" Then End

Open fnm For Input As #1

Line Input #1, XXX

Close #1

ReDim VALID_MONTHS(1 To 12)

For J = 1 To 12
VALID_MONTHS(J) = Mid(XXX, 5 + J, 1)
Next J

Dim DBN

DBN = APPROOT + "\ARTS\WORK\W*.*"

If Dir(DBN) <> "" Then Kill DBN

lstMAJOR.Visible = False
lstMINOR.Visible = False
lstBG.Visible = False
lstSPECIES.Visible = False

lblMAJOR.Visible = False
lblMINOR.Visible = False
lblBG.Visible = False
lblSPECIES.Visible = False

cmdOK.Visible = False
cmdDEL.Visible = False

frmSEL.MousePointer = 1

frmSEL.Caption = msgtab(15) + ": " + Format(CURY, "0000") + " - " + _
                 msgtab(46)

cmdEXIT.ToolTipText = msgtab(3)
cmdRETURN.ToolTipText = msgtab(13)

lblMAJOR.Caption = msgtab(31)
lblMINOR.Caption = msgtab(32)
lblBG.Caption = msgtab(33)
lblSPECIES.Caption = msgtab(34)

cmdFULL.ToolTipText = msgtab(9)
cmdSEL.ToolTipText = msgtab(10)
cmdOK.ToolTipText = msgtab(11)
cmdDEL.ToolTipText = msgtab(12)

cmdGUIDE.ToolTipText = msgtab(6)

Call LOAD_REPTABS

End Sub
Private Sub LOAD_MAJOR()

Dim I, fnm, XXX

fnm = APPROOT + "\ARTS\TABLES\Y" + Format(CURY, "0000") + "_MAJOR.TXT"

If Dir(fnm) = "" Then Exit Sub

NMJ = 0

Open fnm For Input As #1

lstMAJOR.Clear

Do Until EOF(1)

Line Input #1, XXX

NMJ = NMJ + 1

ReDim Preserve TMJC(1 To NMJ), TMJN(1 To NMJ)

TMJC(NMJ) = Val(Left(XXX, 4)): TMJN(NMJ) = Mid(XXX, 6, 30)

lstMAJOR.AddItem TMJN(NMJ)

Loop

Close #1

End Sub
Private Sub LOAD_MINOR()

Dim I, fnm, XXX

fnm = APPROOT + "\ARTS\TABLES\Y" + Format(CURY, "0000") + "_MINOR.TXT"

If Dir(fnm) = "" Then Exit Sub

NMN = 0

Open fnm For Input As #1

lstMINOR.Clear

Do Until EOF(1)

Line Input #1, XXX

NMN = NMN + 1

ReDim Preserve TMNC(1 To NMN), TMNN(1 To NMN), TASSO(1 To NMN)

TMNC(NMN) = Val(Left(XXX, 4)): TMNN(NMN) = Mid(XXX, 6, 30)
TASSO(NMN) = Val(Mid(XXX, 37, 4))

lstMINOR.AddItem TMNN(NMN)

Loop

Close #1

End Sub
Private Sub LOAD_BG()

Dim I, fnm, XXX

fnm = APPROOT + "\ARTS\TABLES\Y" + Format(CURY, "0000") + "_BG.TXT"

If Dir(fnm) = "" Then Exit Sub

NBG = 0

Open fnm For Input As #1

lstBG.Clear

Do Until EOF(1)

Line Input #1, XXX

NBG = NBG + 1

ReDim Preserve TBGC(1 To NBG), TBGN(1 To NBG)

TBGC(NBG) = Val(Left(XXX, 4)): TBGN(NBG) = Mid(XXX, 6, 30)

lstBG.AddItem TBGN(NBG)

Loop

Close #1

End Sub
Private Sub LOAD_SPECIES()

Dim I, fnm, XXX

fnm = APPROOT + "\ARTS\TABLES\Y" + Format(CURY, "0000") + "_SPECIES.TXT"

If Dir(fnm) = "" Then Exit Sub

NSP = 0

Open fnm For Input As #1

lstSPECIES.Clear

Do Until EOF(1)

Line Input #1, XXX

NSP = NSP + 1

ReDim Preserve TSPC(1 To NSP), TSPN(1 To NSP)

TSPC(NSP) = Val(Left(XXX, 4)): TSPN(NSP) = Mid(XXX, 6, 30)

lstSPECIES.AddItem TSPN(NSP)

Loop

Close #1

End Sub
Private Sub FILTER_TOT()

cmdOK.MousePointer = 13

Dim DBN, wnm, XKEY, I

DBN = APPROOT + "\ARTS\WORK\CT" + Format(CURY, "0000") + ".MDB"

If Dir(DBN) = "" Then Exit Sub

wnm = APPROOT + "\ARTS\WORK\WT" + Format(CURY, "0000") + ".MDB"

FileCopy DBN, wnm

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(wnm)
Set prm_record = prm_database.OpenRecordset("ASITAB")

With prm_record

.Index = "primarykey"
.MoveFirst

Do Until .EOF

XKEY = ![akey]

I = Val(Mid(XKEY, 2, 4))

If SELMJ(I) <> "Y" Then
   .Delete
   GoTo NEXT_REC
   End If
   
I = Val(Mid(XKEY, 8, 4))

If SELMN(I) <> "Y" Then
   .Delete
   GoTo NEXT_REC
   End If

I = Val(Mid(XKEY, 14, 4))

If SELBG(I) <> "Y" Then
   .Delete
   GoTo NEXT_REC
   End If

NEXT_REC:

.MoveNext

Loop

NCURT = .RecordCount

End With

prm_record.Close
prm_database.Close

cmdOK.MousePointer = 1

End Sub
Private Sub FILTER_SPECIES()

Dim IBG

cmdOK.MousePointer = 13

Dim DBN, wnm, XKEY, I

DBN = APPROOT + "\ARTS\WORK\CS" + Format(CURY, "0000") + ".MDB"

If Dir(DBN) = "" Then Exit Sub

wnm = APPROOT + "\ARTS\WORK\WS" + Format(CURY, "0000") + ".MDB"

FileCopy DBN, wnm

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(wnm)
Set prm_record = prm_database.OpenRecordset("ASITAB")

With prm_record

.Index = "primarykey"
.MoveFirst

Do Until .EOF

XKEY = ![akey]

I = Val(Mid(XKEY, 2, 4))

If SELMJ(I) <> "Y" Then
   .Delete
   GoTo NEXT_REC
   End If
   
I = Val(Mid(XKEY, 8, 4))

If SELMN(I) <> "Y" Then
   .Delete
   GoTo NEXT_REC
   End If

I = Val(Mid(XKEY, 14, 4))

If SELBG(I) <> "Y" Then
   .Delete
   GoTo NEXT_REC
   End If

I = Val(Mid(XKEY, 20, 4))

If SELSP(I) <> "Y" Then
   .Delete
   GoTo NEXT_REC
   End If
   
NEXT_REC:

.MoveNext

Loop

For IBG = 1 To 10000
SELBG(IBG) = "N"
Next IBG

NCURS = .RecordCount

.MoveFirst

Do Until .EOF

IBG = Val(Mid(![akey], 14, 4))

SELBG(IBG) = "Y"

.MoveNext

Loop

End With

prm_record.Close
prm_database.Close

cmdOK.MousePointer = 1

End Sub
Private Sub CHECK_PRICES()

Dim I, ix, FFC, FFE, FFP, FFU, FFV, FFW, FFN, RR, cc, PP
Dim TC(1 To 13), TP(1 To 13), TV(1 To 13), TF(1 To 13)
Dim INDJ, INDM, INDB, INDS, XKEY, CURYX

CURYX = Format(CURY, "0000")

Dim DBN, XKEY1, XKEY2, NREC, IREC

FNM5 = APPROOT + "\ARTS\CONTROL\NOPRICES.TXT"
Open FNM5 For Output As #5

FNM6 = APPROOT + "\ARTS\CONTROL\NOFISH.TXT"
Open FNM6 For Output As #6

Print #5, Tab(2); msgtab(107)
Print #5, " "

Print #6, Tab(2); msgtab(115)
Print #6, " "

DBN = APPROOT + "\ARTS\WORK\WS" + Format(CURY, "0000") + ".MDB"

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(DBN)
Set prm_record = prm_database.OpenRecordset("ASITAB")

With prm_record

TCATCH = 0: PCATCH = 0: TVALUE = 0: TFISH = 0: FCATCH = 0

NREC = .RecordCount: IREC = 0

.Index = "primarykey"

.MoveFirst

Do Until .EOF

IREC = IREC + 1

XKEY = ![akey]

INDJ = Val(Mid(XKEY, 2, 4))
INDM = Val(Mid(XKEY, 8, 4))
INDB = Val(Mid(XKEY, 14, 4))
INDS = Val(Mid(XKEY, 20, 4))

For I = 1 To 12

ix = Format(I, "00")

FFC = "C" + ix:  TC(I) = .Fields(FFC)
FFP = "P" + ix:  TP(I) = .Fields(FFP)
FFV = "V" + ix:  TV(I) = .Fields(FFV)
FFN = "F" + ix:  TF(I) = .Fields(FFN)

TCATCH = TCATCH + TC(I)
TVALUE = TVALUE + TV(I)
TFISH = TFISH + TF(I)

If TF(I) <> 0 Or TC(I) = 0 Then GoTo NEXT_I1

FCATCH = FCATCH + TC(I)

Print #6, Tab(2); ix + "/" + CURYX; _
          Tab(10); RTrim(REPJ(INDJ)) + " / " + _
                  RTrim(REPM(INDM)) + " / " + _
                  RTrim(REPB(INDB)) + " / " + _
                  RTrim(REPS(INDS))

NEXT_I1:

If TP(I) <> 0 Or TC(I) = 0 Then GoTo NEXT_I

Print #5, Tab(2); ix + "/" + CURYX; _
          Tab(10); RTrim(REPJ(INDJ)) + " / " + _
                  RTrim(REPM(INDM)) + " / " + _
                  RTrim(REPB(INDB)) + " / " + _
                  RTrim(REPS(INDS))

PCATCH = PCATCH + TC(I)

NEXT_I:

Next I

.MoveNext

Loop

End With

prm_record.Close
prm_database.Close

Print #5, " "
Print #5, Tab(2); Format(100 * PCATCH / TCATCH, "####0.000") + " % " + msgtab(104)
Print #5, Tab(2); msgtab(105)

Close #5

Print #6, " "
Print #6, Tab(2); Format(100 * FCATCH / TCATCH, "####0.000") + " % " + msgtab(116)
Print #6, Tab(2); msgtab(105)

Close #6

SYSTEM_VALUES = "YES"

If TVALUE = 0 Then
   SYSTEM_VALUES = "NO"
   Open FNM5 For Output As #5
   Print #5, msgtab(119)
   Close #5
   End If

SYSTEM_FISH = "YES"
If TFISH = 0 Then
   SYSTEM_FISH = "NO"
   Open FNM6 For Output As #6
   Print #6, msgtab(120)
   Close #6
   End If

PCATCH_VALUES = PCATCH / TCATCH: PCATCH_FISH = FCATCH / TCATCH

If SYSTEM_VALUES = "YES" And PCATCH_VALUES = 0 Then
   Open FNM5 For Output As #5
   Print #5, msgtab(117)
   Close #5
   End If

If SYSTEM_FISH = "YES" And PCATCH_FISH = 0 Then
   Open FNM6 For Output As #6
   Print #6, msgtab(118)
   Close #6
   End If

End Sub
Private Sub LOAD_REPTABS()

Dim I, J, K, fnm, XXX

ReDim REPJ(1 To 2000), REPM(1 To 2000), REPB(1 To 2000), REPS(1 To 2000)

For I = 1 To 2000
REPJ(I) = " ": REPM(I) = " ": REPB(I) = " ": REPS(I) = " "
Next I

fnm = APPROOT + "\ARTS\TABLES\Y" + Format(CURY, "0000") + "_MAJOR.TXT"

'=====================================================
Open fnm For Input As #5

Do While Not EOF(5)
Line Input #5, XXX
J = Val(Left(XXX, 4)): REPJ(J) = Mid(XXX, 6, 30)
Loop

Close #5
'=====================================================
fnm = APPROOT + "\ARTS\TABLES\Y" + Format(CURY, "0000") + "_MINOR.TXT"

Open fnm For Input As #5

Do While Not EOF(5)
Line Input #5, XXX
J = Val(Left(XXX, 4)): REPM(J) = Mid(XXX, 6, 30)
Loop

Close #5
'=====================================================
fnm = APPROOT + "\ARTS\TABLES\Y" + Format(CURY, "0000") + "_BG.TXT"

Open fnm For Input As #5

Do While Not EOF(5)
Line Input #5, XXX
J = Val(Left(XXX, 4)): REPB(J) = Mid(XXX, 6, 30)
Loop

Close #5
'=====================================================
fnm = APPROOT + "\ARTS\TABLES\Y" + Format(CURY, "0000") + "_SPECIES.TXT"

Open fnm For Input As #5

Do While Not EOF(5)
Line Input #5, XXX
J = Val(Left(XXX, 4)): REPS(J) = Mid(XXX, 6, 30)
Loop

Close #5

End Sub
