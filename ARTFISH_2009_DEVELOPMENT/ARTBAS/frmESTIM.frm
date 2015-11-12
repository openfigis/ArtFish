VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmESTIM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ARTPLAN 1"
   ClientHeight    =   7215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10710
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   1  'Arrow
   Moveable        =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   10710
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGUIDE 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8520
      MousePointer    =   1  'Arrow
      Picture         =   "frmESTIM.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6240
      Width           =   735
   End
   Begin VB.CommandButton cmdPRINT 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      MousePointer    =   1  'Arrow
      Picture         =   "frmESTIM.frx":2262
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6240
      Width           =   735
   End
   Begin VB.ListBox lstMNBG 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   4380
      Left            =   240
      MultiSelect     =   2  'Extended
      TabIndex        =   5
      Top             =   1200
      Width           =   10215
   End
   Begin VB.CommandButton cmdGO 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      MousePointer    =   1  'Arrow
      Picture         =   "frmESTIM.frx":24E4
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6240
      Width           =   735
   End
   Begin VB.CommandButton cmdQUIT 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   10080
      MousePointer    =   1  'Arrow
      Picture         =   "frmESTIM.frx":2766
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6600
      Width           =   375
   End
   Begin VB.CommandButton cmdRETURN 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   9360
      MousePointer    =   1  'Arrow
      Picture         =   "frmESTIM.frx":29E8
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5880
      Width           =   1095
   End
   Begin ComctlLib.ProgressBar pgbFILES 
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   6480
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   1
   End
   Begin RichTextLib.RichTextBox rtsLOG 
      Height          =   5535
      Left            =   240
      TabIndex        =   9
      Top             =   120
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   9763
      _Version        =   393217
      BackColor       =   12648447
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      MousePointer    =   1
      TextRTF         =   $"frmESTIM.frx":2C6A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FFFF&
      Caption         =   " 10"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   12
      Top             =   6960
      Width           =   255
   End
   Begin VB.Label lblF 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
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
      Height          =   255
      Left            =   8760
      TabIndex        =   8
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label lblA 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
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
      Height          =   255
      Left            =   7440
      TabIndex        =   7
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label lblCUR 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   840
      Width           =   10215
   End
   Begin VB.Label lblTIT 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
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
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   6000
      Width           =   3495
   End
End
Attribute VB_Name = "frmESTIM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private NBG, NS, BGN(), BGC(), MNN(), MNC()
Private SNM(), ASO(), NACT, OKACT(), MNBGC(), MNBG(), MNBGA(), MNBGF()
Private CURMBC, CURMBN, CURIND, CURMINOR, CURBG
Private CURBAC, CURCPUE, CURN, CURCV, CURSN, CURSTD
Private RESBAC, RESCPUE, EFLAG, LFLAG, SPEN()
Private Sub cmdGO_Click()

Dim fnm

fnm = APPROOT + "\ARTBAS\RESULTS\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "*.*"

If Dir(fnm) <> "" Then Kill fnm

fnm = APPROOT + "\ARTBAS\RESULTS\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_ESTIM.TXT"

Open fnm For Output As #1

Close #1

lblTIT.Visible = True
lblTIT.Caption = msgtab(114)
pgbFILES.Visible = True

pgbFILES.Min = 0: pgbFILES.Max = 22

Call LOAD_BG
pgbFILES.Value = 1

Call LOAD_STRATA
pgbFILES.Value = 2

Call LOAD_SITES
pgbFILES.Value = 3

Call LOAD_ASSO
pgbFILES.Value = 4

Call CLEAN_FRAME
pgbFILES.Value = 5

Call CLEAN_ACTIVE
pgbFILES.Value = 6

Call CREATE_ACTIVE
pgbFILES.Value = 7

Call UPDATE_ACTIVE
pgbFILES.Value = 8

Call CLEAN_EFFORT
pgbFILES.Value = 9

Call CLEAN_LANDINGS
pgbFILES.Value = 10

Call LOAD_SPECIES
pgbFILES.Value = 11

Call CLEAN_BYSPECIES
pgbFILES.Value = 12

Call SETUP_EFF
pgbFILES.Value = 13

Call SETUP_LAND2
pgbFILES.Value = 14

Call RAISE_SPECIES
pgbFILES.Value = 15

Call LOAD_MNBG
pgbFILES.Value = 16

Call MAIN_LOOP
pgbFILES.Value = 17

lstMNBG.Visible = False
lblCUR.Visible = False

Call CONTINUE_COMP
pgbFILES.Value = 18

Call CONTINUE_COMP2
pgbFILES.Value = 19

Call CREATE_LOGS
pgbFILES.Value = 20

Call JOIN_LOGS
pgbFILES.Value = 21

Call SPLIT_DATA
pgbFILES.Value = 22

rtsLOG.FileName = APPROOT + "\ARTBAS\RESULTS\WLOG.TXT"
rtsLOG.Refresh
rtsLOG.Visible = True
rtsLOG.MousePointer = 1
frmESTIM.MousePointer = 1
cmdPRINT.Visible = True

lblTIT.Visible = False
pgbFILES.Visible = False
cmdGO.Visible = False

End Sub

Private Sub cmdGUIDE_Click()

HTYPE = "C0"

HFNM = APPROOT + "\ARTBAS\HELP\" + current_language + "HELP" + HTYPE + ".rtf"

If Dir(HFNM) = "" Then Exit Sub

frmESTIM.Enabled = False
Load frmGUIDE
frmGUIDE.Show

End Sub

Private Sub cmdPRINT_Click()

Printer.FontBold = True
Printer.FontName = "Courier"
Printer.FontName = "Courier New"
Printer.FontSize = 10
Printer.FontItalic = False

Dim fnm, XXX

fnm = APPROOT + "\ARTBAS\RESULTS\WLOG.TXT"

Printer.Print " "
Printer.Print " "
Printer.Print " "

Open fnm For Input As #1

Do Until EOF(1)

Line Input #1, XXX
Printer.Print Tab(5); XXX

Loop

Close #1

Printer.EndDoc

End Sub
Private Sub cmdQUIT_Click()

Call CHECK_BACKUP_COMPLETE

Dim fnm

fnm = APPROOT + "\ARTBAS\RESULTS\W*.*"

If Dir(fnm) <> "" Then Kill fnm

cmdRETURN.MousePointer = 13
cmdQUIT.MousePointer = 13

Call KILL_ARTBASIC_FOLDER

Call write_parms
Unload frmESTIM

End Sub
Private Sub cmdRETURN_Click()
cmdRETURN.MousePointer = 13

Dim fnm

fnm = APPROOT + "\ARTBAS\RESULTS\W*.*"

If Dir(fnm) <> "" Then Kill fnm

Load frmARTB01

Unload frmESTIM
frmARTB01.Show
End Sub
Private Sub Form_Load()

Set Picture = LoadPicture(APPROOT + "\ARTBAS\PICS_RUNTIME\SCREEN_10.JPG")

Dim fnm

fnm = APPROOT + "\ARTBAS\RESULTS\W*.*"

If Dir(fnm) <> "" Then Kill fnm

frmESTIM.Caption = monthtab(current_month) + " " + _
                    Format(current_year, "0000") + " - " + msgtab(111)

lblCUR.Visible = False
pgbFILES.Visible = False
lblTIT.Visible = False
rtsLOG.Visible = False
lstMNBG.Visible = False

cmdPRINT.Visible = False
cmdPRINT.ToolTipText = msgtab(52)

cmdGO.ToolTipText = msgtab(113)
cmdRETURN.ToolTipText = msgtab(49)
cmdQUIT.ToolTipText = msgtab(3)
cmdGUIDE.ToolTipText = msgtab(243)

lblTIT.Caption = msgtab(114)

lblA.Caption = msgtab(84)
lblF.Caption = msgtab(86)

lblA.Visible = False
lblF.Visible = False

End Sub
Private Sub CREATE_ACTIVE()

Dim fnm

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_ACTIVE.TXT"

Open fnm For Input As #1

Dim XXX, yyy, K

Dim dbn As String

FileCopy APPROOT + "\ARTBAS\STRUS\ACTIVE.MDB", APPROOT + "\ARTBAS\RESULTS\WACTIVE.MDB"

dbn = APPROOT + "\ARTBAS\RESULTS\WACTIVE.MDB"

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(dbn)
Set prm_record = prm_database.OpenRecordset("ACTAB")

With prm_record

.Index = "sortkey"

Do Until EOF(1)

Line Input #1, XXX

K = Val(Mid(XXX, 2, 4))

If MNN(K) = " " Then GoTo CONT_LOOP

K = Val(Mid(XXX, 8, 4))

If BGN(K) = " " Then GoTo CONT_LOOP

.AddNew

![akey] = Left(XXX, 11)
![aNO] = CDbl(Mid(XXX, 26, 15))
![SortKey] = Mid(XXX, 13, 12)
![afr] = 0

K = Val(Mid(XXX, 2, 4))
![adescr] = MNN(K) + " "
K = Val(Mid(XXX, 8, 4))
![adescr] = ![adescr] + BGN(K)

.Update

CONT_LOOP:

Loop

Close #1

End With

prm_record.Close
prm_database.Close

End Sub
Private Sub LOAD_BG()

ReDim BGC(1 To 10000), BGN(1 To 10000)

Dim I, XXX, fnm, BGCODE, BGNAME

For I = 1 To 10000
BGN(I) = " "
Next I

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_BG.TXT"

Open fnm For Input As #1

Do Until EOF(1)

Line Input #1, XXX

BGCODE = CDbl(Mid(XXX, 1, 4))
BGNAME = Mid(XXX, 6, 30)

BGN(BGCODE) = BGNAME

Loop

Close #1

End Sub
Private Sub LOAD_STRATA()

Dim I, XXX, fnm, MNCODE, MNAME

ReDim MNN(1 To 10000)

For I = 1 To 10000
MNN(I) = " "
Next I

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_MINOR.TXT"

Open fnm For Input As #1

Do Until EOF(1)

Line Input #1, XXX

MNCODE = CDbl(Mid(XXX, 1, 4))
MNAME = Mid(XXX, 6, 30)
MNN(MNCODE) = Left(MNAME + Space(30), 30)

Loop

Close #1

End Sub
Private Sub CLEAN_ACTIVE()

Dim fnm, XXX, yyy, K, PART1, PART2, vvv

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_ACTIVE.TXT"

Open fnm For Input As #1

Open APPROOT + "\ARTBAS\RESULTS\WACTIVE.TXT" For Output As #2

Do Until EOF(1)

Line Input #1, XXX

K = Val(Mid(XXX, 2, 4))

If MNN(K) = " " Then GoTo CONT_LOOP

K = Val(Mid(XXX, 8, 4))

If BGN(K) = " " Then GoTo CONT_LOOP

PART1 = Left(XXX, 25): PART2 = Mid(XXX, 26, 5)

vvv = CDbl(PART2)

If vvv > CURCAL Then vvv = CURCAL

Print #2, PART1 + Format(vvv, "00.00")

CONT_LOOP:

Loop

Close #1
Close #2

FileCopy APPROOT + "\ARTBAS\RESULTS\WACTIVE.TXT", fnm

End Sub
Private Sub LOAD_SITES()

Dim I, XXX, fnm, SICODE, SINAME

ReDim SNM(1 To 10000)

For I = 1 To 10000
SNM(I) = " "
Next I

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_SITES.TXT"

Open fnm For Input As #1

Do Until EOF(1)

Line Input #1, XXX

SICODE = Val(Mid(XXX, 1, 4))
SINAME = Mid(XXX, 6, 30)

SNM(SICODE) = SINAME

Loop

Close #1

End Sub
Private Sub LOAD_SPECIES()

Dim I, XXX, fnm, SICODE, SINAME

ReDim SPEN(1 To 10000)

For I = 1 To 10000
SPEN(I) = " "
Next I

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_Species.TXT"

Open fnm For Input As #1

Do Until EOF(1)

Line Input #1, XXX

SICODE = Val(Mid(XXX, 1, 4))
SINAME = Mid(XXX, 6, 30)

SPEN(SICODE) = SINAME

Loop

Close #1

End Sub
Private Sub CLEAN_FRAME()

Dim fnm, XXX, yyy, K

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_FRAME.TXT"

Open fnm For Input As #1

Open APPROOT + "\ARTBAS\RESULTS\WFRAME.TXT" For Output As #2

Do Until EOF(1)

Line Input #1, XXX

K = Val(Mid(XXX, 2, 4))

If SNM(K) = " " Then GoTo CONT_LOOP

K = Val(Mid(XXX, 8, 4))

If BGN(K) = " " Then GoTo CONT_LOOP

Print #2, XXX

CONT_LOOP:

Loop

Close #1
Close #2

FileCopy APPROOT + "\ARTBAS\RESULTS\WFRAME.TXT", fnm

End Sub
Private Sub LOAD_ASSO()

Dim I, J, K, XXX, yyy, fnm, NAS

ReDim ASO(1 To 10000)

For I = 1 To 10000
ASO(I) = 0
Next I

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_ASSOSI.TXT"

Open fnm For Input As #1

Do Until EOF(1)

Line Input #1, XXX
J = Val(Left(XXX, 4))

NAS = CDbl(Mid(XXX, 37, 4))

For I = 1 To NAS
Line Input #1, yyy
K = Val(Mid(yyy, 6, 4))
ASO(K) = J

Next I

CONT_LOOP:

Loop

Close #1

End Sub
Private Sub UPDATE_ACTIVE()

Dim fnm

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_FRAME.TXT"

Open fnm For Input As #1

Dim XXX, yyy, K, I, J, L, XKEY

Dim dbn As String

dbn = APPROOT + "\ARTBAS\RESULTS\WACTIVE.MDB"

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(dbn)
Set prm_record = prm_database.OpenRecordset("ACTAB")

With prm_record

.Index = "primarykey"

Do Until EOF(1)

Line Input #1, XXX

K = Val(Mid(XXX, 2, 4))
J = ASO(K)
L = CDbl(Mid(XXX, 26, 15))

XKEY = "S" + Format(J, "0000") + "+" + "B" + Mid(XXX, 8, 4)

.Seek "=", XKEY

If .NoMatch = True Then GoTo CONT_LOOP

.Edit

![afr] = ![afr] + L

.Update

CONT_LOOP:

Loop

Close #1

End With

prm_record.Close
prm_database.Close

End Sub
Private Sub CLEAN_EFFORT()

Call LOAD_FRAMEDB

Call SETUP_EFF

Dim fnm, XXX, yyy, K, PART1, PART2, J, XKEY

Dim dbn, dbn2 As String

dbn = APPROOT + "\ARTBAS\RESULTS\WEFF.MDB"
dbn2 = APPROOT + "\ARTBAS\RESULTS\WFRAME.MDB"

Dim prm_database As Database, prm_record As Recordset
Dim prm2_database As Database, prm2_record As Recordset

Set prm_database = OpenDatabase(dbn)
Set prm_record = prm_database.OpenRecordset("ETAB")

Set prm2_database = OpenDatabase(dbn2)
Set prm2_record = prm2_database.OpenRecordset("FRTAB")

With prm_record

.MoveFirst

Do Until .EOF

K = Val(Mid(![ekey], 2, 4)): J = K

If SNM(K) = " " Then GoTo CONT_LOOP

K = Val(Mid(![ekey], 8, 4))

If BGN(K) = " " Then GoTo CONT_LOOP

.Edit

![emnc] = ASO(J)

prm2_record.Index = "primarykey"

prm2_record.Seek "=", Left(![ekey], 11)

If prm2_record.NoMatch = True Then End

.Edit

![efrm] = prm2_record![FNO]

.Update

CONT_LOOP:

.MoveNext

Loop

prm2_record.Close
prm2_database.Close

fnm = APPROOT + "\ARTBAS\EFFORT\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_ESAMPLES.TXT"

Open fnm For Output As #1

.MoveFirst

Do Until .EOF

If ![ekey] <> "ZZZZZZZZZZZZZZZ" And ![eact] <> 0 And ![efrm] > 0 Then
   
   Print #1, ![ekey] + " " + _
             Format(![emnc], "0000") + " " + _
             Format(![eact], "000000.000") + " " + _
             Format(![esmp], "000000.000") + " " + _
             Format(![efrm], "000000.000") + " " + ![erec]
             
   End If
     
.MoveNext

Loop

End With

Close #1

prm_record.Close
prm_database.Close

End Sub
Private Sub LOAD_FRAMEDB()

Dim fnm

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_FRAME.TXT"

Open fnm For Input As #1

Dim XXX, yyy

Dim dbn As String

FileCopy APPROOT + "\ARTBAS\STRUS\FRAME.MDB", APPROOT + "\ARTBAS\RESULTS\WFRAME.MDB"

dbn = APPROOT + "\ARTBAS\RESULTS\WFRAME.MDB"

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(dbn)
Set prm_record = prm_database.OpenRecordset("FRTAB")

With prm_record

Do Until EOF(1)

Line Input #1, XXX

.AddNew

![fkey] = Left(XXX, 11)
![FNO] = CDbl(Mid(XXX, 26, 15))

.Update

Loop

Close #1

End With

prm_record.Close
prm_database.Close

End Sub
Private Sub CLEAN_LANDINGS()

Dim fnm, XXX, yyy, K, PART1, PART2, J

fnm = APPROOT + "\ARTBAS\LANDINGS\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_LSAMPLES.TXT"

Open fnm For Input As #1

Open APPROOT + "\ARTBAS\RESULTS\WLAND.TXT" For Output As #2

Do Until EOF(1)

Line Input #1, XXX

PART1 = Mid(XXX, 1, 84): PART2 = Mid(XXX, 89, 12)

K = Val(Mid(XXX, 91, 4)): J = K

If SNM(K) = " " Then GoTo CONT_LOOP

K = Val(Mid(XXX, 97, 4))

If BGN(K) = " " Then GoTo CONT_LOOP

Print #2, PART1 + Format(ASO(J), "0000") + PART2

CONT_LOOP:

Loop

Close #1
Close #2

FileCopy APPROOT + "\ARTBAS\RESULTS\WLAND.TXT", fnm

Call SETUP_LAND

End Sub
Private Sub CLEAN_BYSPECIES()

Dim fnm, XXX, yyy, K, PART1, PART2, J

fnm = APPROOT + "\ARTBAS\LANDINGS\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_LSPECIES.TXT"

Open fnm For Input As #1

Open APPROOT + "\ARTBAS\RESULTS\WSPECIES.TXT" For Output As #2

Do Until EOF(1)

Line Input #1, XXX

PART1 = Mid(XXX, 1, 94): PART2 = Mid(XXX, 99, 12)

K = Val(Mid(XXX, 101, 4)): J = K

If SNM(K) = " " Then GoTo CONT_LOOP

K = Val(Mid(XXX, 107, 4))

If BGN(K) = " " Then GoTo CONT_LOOP

K = Val(Mid(XXX, 10, 4))

If SPEN(K) = " " Then GoTo CONT_LOOP

Print #2, PART1 + Format(ASO(J), "0000") + PART2

CONT_LOOP:

Loop

Close #1
Close #2

FileCopy APPROOT + "\ARTBAS\RESULTS\WSPECIES.TXT", fnm

End Sub
Private Sub LOAD_MNBG()

Dim dbn As String
Dim I

dbn = APPROOT + "\ARTBAS\RESULTS\WACTIVE.MDB"

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(dbn)
Set prm_record = prm_database.OpenRecordset("ACTAB")

With prm_record

.Index = "SORTKEY"

NACT = .RecordCount: I = 0

ReDim MNBG(1 To NACT), MNBGA(1 To NACT), MNBGF(1 To NACT), OKACT(1 To NACT), MNBGC(1 To NACT)

.MoveFirst

Do Until .EOF

I = I + 1

MNBG(I) = ![adescr]: MNBGA(I) = ![aNO]: MNBGF(I) = ![afr]
MNBGC(I) = "M" + Right(![akey], 10) + "+S0000"

.MoveNext

Loop

End With

prm_record.Close
prm_database.Close

lstMNBG.Clear

For I = 1 To NACT

OKACT(I) = " "

lstMNBG.AddItem MNBG(I) + " " + _
        Right(Space(7) + LTrim(Format(MNBGA(I), "##0.000")), 7) + " " + _
        Right(Space(10) + LTrim(Format(MNBGF(I), "#####0.000")), 10)

Next I

lblCUR.Visible = True
lstMNBG.Visible = True

lblA.Visible = True
lblF.Visible = True

End Sub
Private Sub RELOAD_MNBG()

lstMNBG.Clear

lstMNBG.Enabled = False

Dim I

For I = 1 To NACT

If OKACT(I) <> " " Then GoTo CONT_LOOP

lstMNBG.AddItem MNBG(I) + " " + _
        Right(Space(7) + LTrim(Format(MNBGA(I), "##0.000")), 7) + " " + _
        Right(Space(10) + LTrim(Format(MNBGF(I), "#####0.000")), 10)

CONT_LOOP:

Next I

lstMNBG.Refresh

End Sub
Private Sub lstMNBG_Click()
Call RELOAD_MNBG
End Sub
Private Sub MAIN_LOOP()

Dim I, J

For I = 1 To NACT
OKACT(I) = " "
Next I

For I = 1 To NACT

OKACT(I) = "Y"

CURIND = I
CURMBC = MNBGC(I)
CURMBN = MNBG(I)
CURMINOR = Val(Mid(CURMBC, 2, 4))
CURBG = Mid(CURMBC, 7, 5)

lblCUR.Caption = MNBG(I) + " " + _
        Right(Space(7) + LTrim(Format(MNBGA(I), "##0.000")), 7) + " " + _
        Right(Space(10) + LTrim(Format(MNBGF(I), "#####0.000")), 10)

lblCUR.Refresh

Call RELOAD_MNBG
Call CREATE_TOTDB

If MNBGA(I) = 0 Or MNBGF(I) = 0 Then GoTo CONT_LOOP

Call COMPUTE_CPUE

If LFLAG <> "Y" Then GoTo CONT_LOOP

Call COMPUTE_BAC

If CURBAC <= 0 Then GoTo CONT_LOOP

Call EFF_SVAR
Call UPDATE_ESVAR
Call EFF_TVAR
Call UPDATE_ETVAR
Call COMPUTE_EFFORT

Call COMPUTE_CPUE

If CURCPUE <= 0 Then GoTo CONT_LOOP

Call CPUE_TVAR
Call UPDATE_LTVAR
Call CPUE_SVAR
Call UPDATE_LSVAR
Call COMPUTE_CATCH

Call COMPUTE_Q

CONT_LOOP:

Next I

End Sub
Private Sub CREATE_TOTDB()

If Dir(APPROOT + "\ARTBAS\RESULTS\WTOT.MDB") = "" Then
FileCopy APPROOT + "\ARTBAS\STRUS\ESTOT.MDB", APPROOT + "\ARTBAS\RESULTS\WTOT.MDB"
End If

Dim dbn As String

dbn = APPROOT + "\ARTBAS\RESULTS\WTOT.MDB"

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(dbn)
Set prm_record = prm_database.OpenRecordset("ESTAB")

With prm_record

.Index = "primarykey"

.AddNew

![estkey] = CURMBC
![estdes] = CURMBN
![cal] = CURCAL
![FRNO] = MNBGF(CURIND)
![actno] = MNBGA(CURIND)
![popn] = ![actno] * ![FRNO]
![smpn] = 0

.Update

End With

prm_record.Close
prm_database.Close

End Sub
Private Sub SETUP_EFF()

Dim fnm, dbn, XXX, xcode, xmnc, xact, xsmp, xfrm, xrec

fnm = APPROOT + "\ARTBAS\EFFORT\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_ESAMPLES.TXT"

Open fnm For Input As #1

FileCopy APPROOT + "\ARTBAS\STRUS\EFFORT.MDB", APPROOT + "\ARTBAS\RESULTS\WEFF.MDB"

dbn = APPROOT + "\ARTBAS\RESULTS\WEFF.MDB"

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(dbn)
Set prm_record = prm_database.OpenRecordset("ETAB")

prm_record.MoveFirst
prm_record.Delete

With prm_record

.Index = "primarykey"

Do Until EOF(1)

Line Input #1, XXX

.AddNew

![ekey] = Left(XXX, 15)

xmnc = Val(Mid(XXX, 2, 4))

![emnc] = ASO(xmnc)
![eact] = CDbl(Mid(XXX, 22, 10))
![esmp] = CDbl(Mid(XXX, 33, 10))
![efrm] = CDbl(Mid(XXX, 44, 10))
![erec] = Mid(XXX, 55, 15)

.Update

Loop

Close #1

End With

prm_record.Close
prm_database.Close

End Sub
Private Sub COMPUTE_BAC()

Dim SACT(), SSMP(), SFRM(), SBAC(), SSN, I

Dim dbn, tsmp, tact

tsmp = 0: tact = 0

dbn = APPROOT + "\ARTBAS\RESULTS\WEFF.MDB"

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(dbn)
Set prm_record = prm_database.OpenRecordset("ETAB")

prm_record.MoveFirst

With prm_record

SSN = 0

.MoveFirst

Do Until .EOF

If ![emnc] <> CURMINOR Or Mid(![ekey], 7, 5) <> CURBG Or ![efrm] = 0 Then
   GoTo CONT_READ
   End If
   
If ![esmp] = 0 Then tsmp = tsmp + ![efrm]
If ![esmp] <> 0 Then tsmp = tsmp + ![esmp]

SSN = SSN + 1

ReDim Preserve SACT(1 To SSN), SSMP(1 To SSN), SFRM(1 To SSN), SBAC(1 To SSN)

SACT(SSN) = ![eact]: SSMP(SSN) = ![esmp]: SFRM(SSN) = ![efrm]

If SSMP(SSN) <> 0 Then SBAC(SSN) = SACT(SSN) / SSMP(SSN)
If SSMP(SSN) = 0 Then SBAC(SSN) = SACT(SSN) / SFRM(SSN)

tact = tact + ![eact]

CONT_READ:

.MoveNext

Loop

End With

prm_record.Close
prm_database.Close

dbn = APPROOT + "\ARTBAS\RESULTS\WTOT.MDB"

Set prm_database = OpenDatabase(dbn)
Set prm_record = prm_database.OpenRecordset("ESTAB")

Dim XKEY, NNP, NNS, NNA, XBAC, VAR, STD, COEFF

With prm_record

.Index = "primarykey"

XKEY = "M" + Right(CURMBC, 16)

.Seek "=", XKEY

If .NoMatch = True Then End

.Edit

![bac] = 0: EFLAG = "N"

If tsmp = 0 Then
   EFLAG = "N": CURBAC = 0
   Exit Sub
   End If

If tsmp <> 0 Then ![bac] = tact / tsmp

CURBAC = ![bac]: EFLAG = "Y"

![smpn] = tsmp
![eact] = tact

NNA = ![eact]
NNP = ![popn]
NNS = ![smpn]

XBAC = ![bac]

VAR = 0

For I = 1 To SSN

VAR = VAR + (SBAC(I) - XBAC) ^ 2

Next I

If SSN <= 1 Then
   ![bac_cv] = 0.5102
   GoTo END_UPDATE
   End If

VAR = VAR / (SSN - 1)
STD = VAR ^ 0.5
STD = STD / SSN ^ 0.5

COEFF = (1 - NNS / NNP)

If COEFF < 0 Then
   COEFF = 0
   STD = 0
   End If

COEFF = COEFF ^ 0.5

![bac_cv] = 0

If XBAC <> 0 Then ![bac_cv] = COEFF * STD / XBAC

If ![bac_cv] > 0.5102 Then ![bac_cv] = 0.5102

END_UPDATE:

.Update

End With

prm_record.Close
prm_database.Close

End Sub
Private Sub EFF_SVAR()

If EFLAG <> "Y" Then Exit Sub

Dim NOSI(), SISMP(), SIACT(), SIFRM(), SIBAC(), I, WW()

ReDim NOSI(1 To 10000), SISMP(1 To 10000), SIFRM(1 To 10000), WW(1 To 1000)
ReDim SIACT(1 To 10000), SIBAC(1 To 10000)

For I = 1 To 10000
NOSI(I) = 0: SISMP(I) = 0: SIACT(I) = 0:  SIBAC(I) = 0: SIFRM(I) = 0
Next I

Dim dbn, tsmp, tact, XKEY

dbn = APPROOT + "\ARTBAS\RESULTS\WEFF.MDB"

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(dbn)
Set prm_record = prm_database.OpenRecordset("ETAB")

prm_record.MoveFirst

With prm_record

Dim xsmp, xact, XFR

.MoveFirst

Do Until .EOF

If ![emnc] <> CURMINOR Or Mid(![ekey], 7, 5) <> CURBG Then GoTo CONT_READ

xsmp = ![esmp]
xact = ![eact]
XFR = ![efrm]

I = Mid(![ekey], 2, 4)

NOSI(I) = NOSI(I) + 1

If xsmp = 0 Then SISMP(I) = SISMP(I) + XFR
If xsmp <> 0 Then SISMP(I) = SISMP(I) + xsmp

SIACT(I) = SIACT(I) + xact

If xsmp = 0 Then WW(I) = XFR
If xsmp <> 0 Then WW(I) = xsmp

CONT_READ:

.MoveNext

Loop

End With

prm_record.Close
prm_database.Close

'==========================
'Find data
'==========================

Dim VAR

VAR = 0: CURN = 0: CURSN = 0

For I = 1 To 10000

If NOSI(I) = 0 Then GoTo NEXT_I

SIBAC(I) = 0

If SISMP(I) <> 0 Then SIBAC(I) = SIACT(I) / SISMP(I)

CURN = CURN + 1: CURSN = CURSN + WW(I)

VAR = VAR + WW(I) * (CURBAC - SIBAC(I)) ^ 2

NEXT_I:

Next I

CURSTD = -999

If CURN <= 1 Then Exit Sub

VAR = VAR / (CURSN - 1)

CURSTD = (VAR) ^ 0.5: CURSTD = CURSTD / CURSN ^ 0.5

End Sub
Private Sub UPDATE_ESVAR()

If EFLAG <> "Y" Then Exit Sub

Dim dbn

dbn = APPROOT + "\ARTBAS\RESULTS\WTOT.MDB"

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(dbn)
Set prm_record = prm_database.OpenRecordset("ESTAB")

Dim XKEY, STD, COEFF

With prm_record

.Index = "primarykey"

XKEY = "M" + Right(CURMBC, 16)

.Seek "=", XKEY

If .NoMatch = True Then End

.Edit

COEFF = 1 - CURSN / ![FRNO]

If COEFF < 0 Then COEFF = 0

![esites] = CURN
![esmp] = CURSN

![bac_cvs] = 0

If CURSTD = -999 Then ![bac_cvs] = 0.5102

If CURSTD <> -999 Then
   If CURBAC <> 0 Then ![bac_cvs] = (COEFF ^ 0.5) * CURSTD / CURBAC
   End If
   
'If ![bac_cvs] > 0.5102 Then ![bac_cvs] = 0.5102

.Update

End With

prm_record.Close
prm_database.Close

End Sub
Private Sub EFF_TVAR()

If EFLAG <> "Y" Then Exit Sub

Dim NOSI(), SISMP(), SIACT(), SIFRM(), SIBAC(), I

ReDim NOSI(1 To 31), SISMP(1 To 31), SIFRM(1 To 31)
ReDim SIACT(1 To 31), SIBAC(1 To 31)

For I = 1 To 31
NOSI(I) = 0: SISMP(I) = 0: SIACT(I) = 0:  SIBAC(I) = 0
Next I

Dim dbn, tsmp, tact, XKEY

dbn = APPROOT + "\ARTBAS\RESULTS\WEFF.MDB"

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(dbn)
Set prm_record = prm_database.OpenRecordset("ETAB")

prm_record.MoveFirst

With prm_record

Dim xsmp, xact, XFR

.MoveFirst

Do Until .EOF

If ![emnc] <> CURMINOR Or Mid(![ekey], 7, 5) <> CURBG Then GoTo CONT_READ

xsmp = ![esmp]
xact = ![eact]
XFR = ![efrm]

I = Right(RTrim(![ekey]), 2)

NOSI(I) = NOSI(I) + 1

If xsmp = 0 Then SISMP(I) = SISMP(I) + XFR
If xsmp <> 0 Then SISMP(I) = SISMP(I) + xsmp

SIFRM(I) = SIFRM(I) + XFR

SIACT(I) = SIACT(I) + xact

CONT_READ:

.MoveNext

Loop

End With

prm_record.Close
prm_database.Close

'==========================
'Find data
'==========================

Dim VAR

VAR = 0: CURN = 0: CURSN = 0

For I = 1 To 31

If NOSI(I) = 0 Then GoTo NEXT_I

SIBAC(I) = 0

If SISMP(I) <> 0 Then SIBAC(I) = SIACT(I) / SISMP(I)

CURN = CURN + 1: CURSN = CURSN + SISMP(I)

VAR = VAR + (CURBAC - SIBAC(I)) ^ 2

NEXT_I:

Next I

CURSTD = -999

If CURN <= 1 Then Exit Sub

VAR = VAR / (CURN - 1)

CURSTD = (VAR) ^ 0.5: CURSTD = CURSTD / CURN ^ 0.5

End Sub
Private Sub UPDATE_ETVAR()

If EFLAG <> "Y" Then Exit Sub

Dim dbn

dbn = APPROOT + "\ARTBAS\RESULTS\WTOT.MDB"

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(dbn)
Set prm_record = prm_database.OpenRecordset("ESTAB")

Dim XKEY, STD, COEFF

With prm_record

.Index = "primarykey"

XKEY = "M" + Right(CURMBC, 16)

.Seek "=", XKEY

If .NoMatch = True Then End

.Edit

COEFF = 1 - CURN / ![actno]

If COEFF < 0 Then COEFF = 0

![edays] = CURN

![bac_cvt] = 0

If CURSTD = -999 Then ![bac_cvt] = 0.5102

If CURSTD <> -999 Then
   If CURBAC <> 0 Then ![bac_cvt] = COEFF ^ 0.5 * CURSTD / CURBAC
   End If
   
'If ![bac_cvt] > 0.5102 Then ![bac_cvt] = 0.5102

.Update

End With

prm_record.Close
prm_database.Close

End Sub
Private Sub COMPUTE_EFFORT()

If EFLAG <> "Y" Then Exit Sub

Dim dbn

dbn = APPROOT + "\ARTBAS\RESULTS\WTOT.MDB"

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(dbn)
Set prm_record = prm_database.OpenRecordset("ESTAB")

Dim XKEY, STD, COEFF

With prm_record

.Index = "primarykey"

XKEY = "M" + Right(CURMBC, 16)

.Seek "=", XKEY

If .NoMatch = True Then End

Dim cv

.Edit

If ![bac_cvs] = 0.5102 And ![bac_cvt] <> 0.5102 Then
   ![bac_cvsp] = 100: ![bac_cvtp] = 0
   GoTo END_CALC
   End If

If ![bac_cvs] <> 0.5102 And ![bac_cvt] = 0.5102 Then
   ![bac_cvsp] = 0: ![bac_cvtp] = 100
   GoTo END_CALC
   End If

If ![bac_cvs] <> 0.5102 And ![bac_cvt] <> 0.5102 Then

   If ![bac_cvs] <> 0 And ![bac_cvt] <> 0 Then
   ![bac_cvsp] = (100 * ![bac_cvs] / (![bac_cvs] + ![bac_cvt]))
   ![bac_cvtp] = 100 - ![bac_cvsp]
   GoTo END_CALC
   End If
   
   If ![bac_cvs] = 0 And ![bac_cvt] <> 0 Then
   ![bac_cvsp] = 0: ![bac_cvtp] = 100
   GoTo END_CALC
   End If
   
   If ![bac_cvt] = 0 And ![bac_cvs] <> 0 Then
   ![bac_cvtp] = 0: ![bac_cvsp] = 100
   GoTo END_CALC
   End If
   
   End If

END_CALC:

If ![bac_cv] = 0 Then
   ![bac_cvsp] = 0: ![bac_cvtp] = 0
   End If

STD = ![bac_cv] * ![bac]

![bac_low] = ![bac] - 1.96 * STD
![bac_upper] = ![bac] + 1.96 * STD

If ![bac_low] < 0 Then
   ![bac_low] = 0: ![bac_upper] = 2 * ![bac]
   End If
   
If ![bac_cv] = 0.5102 Then
   ![bac_low] = 0: ![bac_upper] = 0
   End If
   
![eff] = ![bac] * ![actno] * ![FRNO]
![eff_low] = ![bac_low] * ![actno] * ![FRNO]
![eff_upper] = ![bac_upper] * ![actno] * ![FRNO]

Dim xvar

![BAC_ACCUR] = 0

If ![smpn] <> 0 And ![popn] <> 0 Then
   
   xvar = Log(![smpn]) / Log(![popn])
   
   If xvar > 1 Then xvar = 1
   
   ![BAC_ACCUR] = 1 - 1.96 * 0.5 / ![popn] ^ 0.5 * (![popn] ^ (1 - xvar) - 1) ^ 0.5
   
   End If

.Update

End With

prm_record.Close
prm_database.Close

End Sub
Private Sub COMPUTE_CPUE()

Dim dbn, tsmp, teff, TNL, VAR, STD, COEFF, xcpue, NNP, NNS, NNN, I

Dim QC(), QT(), QE()

tsmp = 0: teff = 0: TNL = 0

dbn = APPROOT + "\ARTBAS\RESULTS\WLAND.MDB"

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(dbn)
Set prm_record = prm_database.OpenRecordset("LTAB")

prm_record.MoveFirst

With prm_record

TNL = 0: tsmp = 0: teff = 0

.MoveFirst

Do Until .EOF

If ![LMNC] <> CURMINOR Or Mid(![lsbc], 7, 5) <> CURBG Then GoTo CONT_READ

TNL = TNL + 1

ReDim Preserve QC(1 To TNL), QT(1 To TNL), QE(1 To TNL)

tsmp = tsmp + ![ltot]
teff = teff + ![LNOU] * ![LDUR]

QT(TNL) = ![ltot]
QE(TNL) = ![LNOU] * ![LDUR]
QC(TNL) = 0

If QE(TNL) <> 0 Then QC(TNL) = QT(TNL) / QE(TNL)

CONT_READ:

.MoveNext

Loop

End With

prm_record.Close
prm_database.Close

If TNL <= 1 Then
   LFLAG = "N"
   Exit Sub
   End If

dbn = APPROOT + "\ARTBAS\RESULTS\WTOT.MDB"

Set prm_database = OpenDatabase(dbn)
Set prm_record = prm_database.OpenRecordset("ESTAB")

Dim XKEY

With prm_record

.Index = "primarykey"

XKEY = "M" + Right(CURMBC, 16)

.Seek "=", XKEY

If .NoMatch = True Then End

.Edit

![cpue] = 0: LFLAG = "N"

If teff <> 0 Then
   ![cpue] = tsmp / teff
   LFLAG = "Y"
   End If
   
CURCPUE = ![cpue]

![ltot] = tsmp
![leff] = teff
![nland] = TNL

NNP = ![bac] * ![actno] * ![FRNO]

If NNP <= 0 Then Exit Sub

NNS = TNL

![LPOP] = NNP

VAR = 0

xcpue = ![cpue]

For I = 1 To TNL

VAR = VAR + (QC(I) - xcpue) ^ 2

Next I

VAR = VAR / (TNL - 1)

STD = (VAR ^ 0.5) / TNL ^ 0.5

COEFF = 0

If NNP <> 0 Then COEFF = (1 - NNS / NNP)

If COEFF < 0 Then
   COEFF = 0
   STD = 0
   End If
   
COEFF = COEFF ^ 0.5

![cpue_cv] = 0

If xcpue <> 0 Then ![cpue_cv] = STD / xcpue

If ![cpue_cv] > 0.5102 Then ![cpue_cv] = 0.5102

STD = ![cpue_cv] * ![cpue]

![cpue_low] = ![cpue] - 1.96 * STD
![cpue_upper] = ![cpue] + 1.96 * STD

If ![cpue_low] < 0 Then
   ![cpue_low] = 0: ![cpue_upper] = 2 * ![cpue]
   End If

Dim xlog

COEFF = ((2 * NNP - 1) / (6 * NNP - 6) - 0.25) ^ 0.5

xlog = Log(TNL) / Log(NNP)

If xlog > 1 Then xlog = 1

![CPUE_ACCUR] = 1 - 1.96 * (COEFF / (NNP) ^ 0.5) * (NNP ^ (1 - xlog) - 1) ^ 0.5

.Update

End With

prm_record.Close
prm_database.Close
End Sub
Private Sub SETUP_LAND()

Dim fnm, dbn, XXX, xcode, xmnc, xact, xsmp, xfrm, xrec

fnm = APPROOT + "\ARTBAS\LANDINGS\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_LSAMPLES.TXT"

Open fnm For Input As #1

FileCopy APPROOT + "\ARTBAS\STRUS\LSAMPLES.MDB", APPROOT + "\ARTBAS\RESULTS\WLAND.MDB"

dbn = APPROOT + "\ARTBAS\RESULTS\WLAND.MDB"

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(dbn)
Set prm_record = prm_database.OpenRecordset("LTAB")

With prm_record

.Index = "primarykey"

Do Until EOF(1)

Line Input #1, XXX

.AddNew

![LDOC] = CDbl(Left(XXX, 6))
![LDAY] = CDbl(Mid(XXX, 8, 2))
![LNOU] = CDbl(Mid(XXX, 11, 8))
![LDUR] = CDbl(Mid(XXX, 20, 8))
![LSMP] = CDbl(Mid(XXX, 29, 19))
![ltot] = CDbl(Mid(XXX, 49, 19))
![LREC] = Mid(XXX, 69, 15)
![LMNC] = CDbl(Mid(XXX, 85, 4))
![lsbc] = Mid(XXX, 90, 11)

.Update

Loop

Close #1

End With

prm_record.Close
prm_database.Close

End Sub
Private Sub CPUE_TVAR()

If LFLAG <> "Y" Then Exit Sub

Dim NOSI(), SITOT(), SIEFF(), SIFRM(), SICPUE(), I

ReDim NOSI(1 To 31), SITOT(1 To 31), SIEFF(1 To 31)
ReDim SICPUE(1 To 31)

For I = 1 To 31
NOSI(I) = 0: SITOT(I) = 0: SIEFF(I) = 0:  SICPUE(I) = 0
Next I

Dim dbn, tsmp, tact, XKEY

dbn = APPROOT + "\ARTBAS\RESULTS\WLAND.MDB"

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(dbn)
Set prm_record = prm_database.OpenRecordset("LTAB")

prm_record.MoveFirst

With prm_record

Dim xsmp, xact, xeff, xcpue

.MoveFirst

Do Until .EOF

If ![LMNC] <> CURMINOR Or Mid(![lsbc], 7, 5) <> CURBG Then GoTo CONT_READ

xsmp = ![ltot]
xeff = ![LNOU] * ![LDUR]

I = ![LDAY]

NOSI(I) = NOSI(I) + 1
SIEFF(I) = SIEFF(I) + xeff
SITOT(I) = SITOT(I) + xsmp

CONT_READ:

.MoveNext

Loop

End With

prm_record.Close
prm_database.Close

'==========================
'Find data
'==========================

Dim VAR

VAR = 0: CURN = 0: CURSN = 0

For I = 1 To 31

If NOSI(I) = 0 Then GoTo NEXT_I

SICPUE(I) = 0

If SIEFF(I) <> 0 Then SICPUE(I) = SITOT(I) / SIEFF(I)

CURN = CURN + 1

VAR = VAR + (CURCPUE - SICPUE(I)) ^ 2

NEXT_I:

Next I

CURSTD = -999

If CURN <= 1 Then Exit Sub

VAR = VAR / (CURN - 1)

CURSTD = (VAR) ^ 0.5: CURSTD = CURSTD / CURN ^ 0.5

End Sub
Private Sub UPDATE_LTVAR()

If LFLAG <> "Y" Then Exit Sub

If CURN = 0 Then Exit Sub

Dim dbn

dbn = APPROOT + "\ARTBAS\RESULTS\WTOT.MDB"

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(dbn)
Set prm_record = prm_database.OpenRecordset("ESTAB")

Dim XKEY, STD, COEFF, NP, xlog, m

With prm_record

.Index = "primarykey"

XKEY = "M" + Right(CURMBC, 16)

.Seek "=", XKEY

If .NoMatch = True Then End

.Edit

![ldays] = CURN

CURCPUE = ![cpue]

NP = ![actno]: m = CURN

COEFF = (1 - m / NP)

If COEFF < 0 Then COEFF = 0

COEFF = COEFF ^ 0.5

![cpue_cvt] = 0

If CURSTD = -999 Then ![cpue_cvt] = 0.5102

If CURSTD <> -999 Then
   If CURCPUE <> 0 Then ![cpue_cvt] = COEFF * (CURSTD) / CURCPUE
   End If
   
'If ![cpue_cvt] > 0.5102 Then ![cpue_cvt] = 0.5102

.Update

End With

prm_record.Close
prm_database.Close

End Sub
Private Sub CPUE_SVAR()

If LFLAG <> "Y" Then Exit Sub

Dim NOSI(), SITOT(), SIEFF(), SIFRM(), SICPUE(), I

ReDim NOSI(1 To 10000), SITOT(1 To 10000), SIFRM(1 To 10000)
ReDim SICPUE(1 To 10000), SIEFF(1 To 10000)

For I = 1 To 10000
NOSI(I) = 0: SITOT(I) = 0: SIEFF(I) = 0:  SICPUE(I) = 0: SIFRM(I) = 0
Next I

Dim dbn, tsmp, teff, XKEY

dbn = APPROOT + "\ARTBAS\RESULTS\WLAND.MDB"

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(dbn)
Set prm_record = prm_database.OpenRecordset("LTAB")

Dim prm2_database As Database, prm2_record As Recordset

Set prm2_database = OpenDatabase(APPROOT + "\ARTBAS\RESULTS\WFRAME.MDB")
Set prm2_record = prm2_database.OpenRecordset("FRTAB")

prm2_record.Index = "primarykey"

prm_record.MoveFirst

With prm_record

Dim xsmp, xeff, XFR

.MoveFirst

Do Until .EOF

If ![LMNC] <> CURMINOR Or Mid(![lsbc], 7, 5) <> CURBG Then GoTo CONT_READ

prm2_record.Seek "=", ![lsbc]

If prm2_record.NoMatch = True Then End

xsmp = ![ltot]
xeff = ![LNOU] * ![LDUR]
XFR = prm2_record![FNO]

I = Mid(![lsbc], 2, 4)

NOSI(I) = NOSI(I) + 1
SITOT(I) = SITOT(I) + xsmp
SIFRM(I) = XFR
SIEFF(I) = SIEFF(I) + xeff

CONT_READ:

.MoveNext

Loop

End With

prm_record.Close
prm_database.Close

prm2_record.Close
prm2_database.Close

'==========================
'Find data
'==========================

Dim VAR

VAR = 0: CURN = 0: CURSN = 0

For I = 1 To 10000

If NOSI(I) = 0 Then GoTo NEXT_I

SICPUE(I) = 0

If SIEFF(I) <> 0 Then SICPUE(I) = SITOT(I) / SIEFF(I)

CURN = CURN + 1: CURSN = CURSN + NOSI(I)

VAR = VAR + NOSI(I) * (CURCPUE - SICPUE(I)) ^ 2

NEXT_I:

Next I

CURSTD = -999

If CURN <= 1 Then Exit Sub

VAR = VAR / (CURSN - 1)

CURSTD = (VAR) ^ 0.5: CURSTD = CURSTD / CURSN ^ 0.5

End Sub
Private Sub UPDATE_LSVAR()

If LFLAG <> "Y" Then Exit Sub

If CURN = 0 Then Exit Sub

Dim dbn

dbn = APPROOT + "\ARTBAS\RESULTS\WTOT.MDB"

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(dbn)
Set prm_record = prm_database.OpenRecordset("ESTAB")

Dim XKEY, STD, COEFF, NP, xlog, m

With prm_record

.Index = "primarykey"

XKEY = "M" + Right(CURMBC, 16)

.Seek "=", XKEY

If .NoMatch = True Then End

.Edit

![lsites] = CURN

CURCPUE = ![cpue]

NP = Int(![bac] * ![FRNO] * ![actno]): m = CURSN

If NP <= 0 Then Exit Sub

COEFF = (1 - m / NP)

If COEFF < 0 Then COEFF = 0

COEFF = COEFF ^ 0.5

![cpue_cvs] = 0

If CURSTD = -999 Then ![cpue_cvs] = 0.5102

If CURSTD <> -999 Then
   If CURCPUE <> 0 Then ![cpue_cvs] = COEFF * (CURSTD) / CURCPUE
   End If
   
'If ![cpue_cvs] > 0.5102 Then ![cpue_cvs] = 0.5102

.Update

End With

prm_record.Close
prm_database.Close

End Sub
Private Sub COMPUTE_CATCH()

If LFLAG <> "Y" Then Exit Sub

Dim dbn

dbn = APPROOT + "\ARTBAS\RESULTS\WTOT.MDB"

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(dbn)
Set prm_record = prm_database.OpenRecordset("ESTAB")

Dim XKEY, STD, COEFF

With prm_record

.Index = "primarykey"

XKEY = "M" + Right(CURMBC, 16)

.Seek "=", XKEY

If .NoMatch = True Then End

Dim cv

.Edit

If ![cpue_cvs] = 0.5102 And ![cpue_cvt] <> 0.5102 Then
   ![cpue_cvsp] = 100: ![cpue_cvtp] = 0
   GoTo END_CALC
   End If

If ![cpue_cvs] <> 0.5102 And ![cpue_cvt] = 0.5102 Then
   ![cpue_cvsp] = 0: ![cpue_cvtp] = 100
   GoTo END_CALC
   End If

If ![cpue_cvs] <> 0.5102 And ![cpue_cvt] <> 0.5102 Then

   If ![cpue_cvs] <> 0 And ![cpue_cvt] <> 0 Then
   ![cpue_cvsp] = (100 * ![cpue_cvs] / (![cpue_cvs] + ![cpue_cvt]))
   ![cpue_cvtp] = 100 - ![cpue_cvsp]
   GoTo END_CALC
   End If
   
   If ![cpue_cvs] = 0 And ![cpue_cvt] <> 0 Then
   ![cpue_cvsp] = 0: ![cpue_cvtp] = 100
   GoTo END_CALC
   End If
   
   If ![cpue_cvt] = 0 And ![cpue_cvs] <> 0 Then
   ![cpue_cvtp] = 0: ![cpue_cvsp] = 100
   GoTo END_CALC
   End If
   
   End If

END_CALC:

If ![cpue_cv] = 0 Then
   ![cpue_cvsp] = 0: ![cpue_cvtp] = 0
   End If

![catch] = ![eff] * ![cpue]

![catch_cv] = (![cpue_cv] ^ 2 + ![bac_cv] ^ 2) ^ 0.5

If ![catch_cv] > 0.5102 Then ![catch_cv] = 0.5102

STD = ![catch_cv] * ![catch]

![catch_low] = ![catch] - 1.96 * STD
![catch_upper] = ![catch] + 1.96 * STD

If ![catch_low] < 0 Then
   ![catch_low] = 0: ![catch_upper] = 2 * ![catch]
   End If
   
If ![catch_cv] = 0.5102 Then
   ![catch_low] = 0: ![catch_upper] = 0
   End If
   
.Update

End With

prm_record.Close
prm_database.Close

End Sub
Private Sub SETUP_LAND2()

Dim fnm, dbn, XXX, xcode, xmnc, xact, xsmp, xfrm, xrec

fnm = APPROOT + "\ARTBAS\RESULTS\WSPECIES.TXT"

Open fnm For Input As #1

FileCopy APPROOT + "\ARTBAS\STRUS\LSPECIES.MDB", APPROOT + "\ARTBAS\RESULTS\WSPEC.MDB"

dbn = APPROOT + "\ARTBAS\RESULTS\WSPEC.MDB"

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(dbn)
Set prm_record = prm_database.OpenRecordset("STAB")

With prm_record

.Index = "primarykey"

Do Until EOF(1)

Line Input #1, XXX

.AddNew

![skey] = Left(XXX, 13)
![slan] = CDbl(Mid(XXX, 15, 19))
![snof] = CDbl(Mid(XXX, 35, 19))
![spri] = CDbl(Mid(XXX, 55, 19))
![sval] = CDbl(Mid(XXX, 75, 19))
![smnc] = Val(Mid(XXX, 95, 4))
![ssbc] = Mid(XXX, 100, 11)

.Update

Loop

Close #1

End With

prm_record.Close
prm_database.Close

End Sub
Private Sub RAISE_SPECIES()

Dim fnm, dbn, XXX, xcode, xmnc, xact, xsmp, xfrm, xrec, dbn2

dbn = APPROOT + "\ARTBAS\RESULTS\WSPEC.MDB"

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(dbn)
Set prm_record = prm_database.OpenRecordset("STAB")

dbn2 = APPROOT + "\ARTBAS\RESULTS\WLAND.MDB"

Dim prm2_database As Database, prm2_record As Recordset

Set prm2_database = OpenDatabase(dbn2)
Set prm2_record = prm2_database.OpenRecordset("LTAB")

prm2_record.Index = "primarykey"

With prm_record

.MoveFirst

Do Until .EOF

XXX = Val(Mid(![skey], 2, 6))

prm2_record.Seek "=", XXX

If prm2_record.NoMatch = True Then End

.Edit

![slan] = ![slan] * (prm2_record![ltot] / prm2_record![LSMP])
![snof] = ![snof] * (prm2_record![ltot] / prm2_record![LSMP])
![sval] = ![slan] * ![spri]

.Update

.MoveNext

Loop

End With

prm_record.Close
prm_database.Close

prm2_record.Close
prm2_database.Close

End Sub
Private Sub COMPUTE_Q()

Dim NUSP, TESTQ, TESTSP

If LFLAG <> "Y" Then Exit Sub

NUSP = 0: TESTSP = 0

Dim fnm, dbn, XXX, xcode, xmnc, xact, xsmp, xfrm, xrec, dbn2, I

dbn = APPROOT + "\ARTBAS\RESULTS\WSPEC.MDB"

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(dbn)
Set prm_record = prm_database.OpenRecordset("STAB")

dbn2 = APPROOT + "\ARTBAS\RESULTS\WTOT.MDB"

Dim prm2_database As Database, prm2_record As Recordset

Set prm2_database = OpenDatabase(dbn2)
Set prm2_record = prm2_database.OpenRecordset("ESTAB")

prm2_record.Index = "primarykey"

With prm_record

.MoveFirst

Do Until .EOF

If CURMINOR <> ![smnc] Then GoTo NEXT_RECORD
If CURBG <> Mid(![ssbc], 7, 5) Then GoTo NEXT_RECORD

XXX = "M" + Format(![smnc], "0000")
XXX = XXX + "+" + Mid(![ssbc], 7, 5)
XXX = XXX + "+" + Right(RTrim(![skey]), 5)

I = Val(Right(XXX, 4))

prm2_record.Seek "=", XXX

If prm2_record.NoMatch = True Then

    NUSP = NUSP + 1

    prm2_record.AddNew
    
    prm2_record![estdes] = SPEN(I)
    prm2_record![estkey] = XXX
    prm2_record![ltot] = ![slan]
    
    If ![sval] <> 0 Then
    prm2_record![lsmpv] = ![slan]
    prm2_record![Value] = ![sval]
    End If
    
    If ![snof] <> 0 Then
    prm2_record![lsmpf] = ![slan]
    prm2_record![fish] = ![snof]
    End If
        
    prm2_record.Update

    GoTo NEXT_RECORD

    End If

prm2_record.Edit

prm2_record![ltot] = prm2_record![ltot] + ![slan]
    
If ![sval] <> 0 Then
   prm2_record![lsmpv] = prm2_record![lsmpv] + ![slan]
   prm2_record![Value] = prm2_record![Value] + ![sval]
   End If
   
If ![snof] <> 0 Then
   prm2_record![lsmpf] = prm2_record![lsmpf] + ![slan]
   prm2_record![fish] = prm2_record![fish] + ![snof]
   End If
    
prm2_record.Update

NEXT_RECORD:

.MoveNext

Loop

End With

prm_record.Close
prm_database.Close

prm2_record.Close
prm2_database.Close

End Sub
Private Sub CONTINUE_COMP()

Dim I, KEYT(), dbn, NN, XKEY

dbn = APPROOT + "\ARTBAS\RESULTS\WTOT.MDB"

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(dbn)
Set prm_record = prm_database.OpenRecordset("ESTAB")

With prm_record

NN = .RecordCount

ReDim KEYT(1 To NN)

.MoveFirst

I = 0

Do Until .EOF

I = I + 1

KEYT(I) = ![estkey]

.MoveNext

Loop

.Index = "primarykey"

For I = 1 To NN

If Right(KEYT(I), 4) = "0000" Then GoTo NEXT_I

XKEY = Left(KEYT(I), 11) + "+" + "S0000"

.Seek "=", XKEY

If .NoMatch = True Then End

Dim xtot, xcatch, xeffort, xval

xtot = ![ltot]
xcatch = ![catch]
xeffort = ![eff]

.Seek "=", KEYT(I)

If .NoMatch = True Then End

.Edit

![catch] = 0
If xtot <> 0 Then ![catch] = (![ltot] / xtot) * xcatch
![eff] = xeffort
![cpue] = 0

If xeffort <> 0 Then ![cpue] = ![catch] / ![eff]

If ![lsmpv] <> 0 Then ![price] = ![Value] / ![lsmpv]
If ![fish] <> 0 Then ![kgfish] = ![lsmpf] / ![fish]

![Value] = ![catch] * ![price]

![fish] = 0

If ![kgfish] <> 0 Then ![fish] = ![catch] / ![kgfish]

xval = ![Value]

.Update

.Seek "=", XKEY

.Edit

![Value] = ![Value] + xval

![price] = 0

If ![catch] <> 0 Then ![price] = ![Value] / ![catch]

.Update

NEXT_I:

Next I

End With

prm_record.Close
prm_database.Close

End Sub
Private Sub CONTINUE_COMP2()

Dim I, KEYT(), dbn, NN, XKEY

Dim C(), E(), V(), D(), FR()

dbn = APPROOT + "\ARTBAS\RESULTS\WTOT.MDB"

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(dbn)
Set prm_record = prm_database.OpenRecordset("ESTAB")

With prm_record

NN = .RecordCount

ReDim KEYT(1 To NN), C(1 To NN), E(1 To NN), V(1 To NN), D(1 To NN), FR(1 To NN)

.MoveFirst

I = 0

Do Until .EOF

I = I + 1

KEYT(I) = ![estkey]
C(I) = ![catch]
E(I) = ![eff]
V(I) = ![Value]
D(I) = ![estdes]
FR(I) = ![FRNO]

.MoveNext

Loop

.Index = "primarykey"

For I = 1 To NN

XKEY = Left(KEYT(I), 5) + "+B0000+" + Right(RTrim(KEYT(I)), 5)

.Seek "=", XKEY

If .NoMatch = True Then

   .AddNew
   
   ![estdes] = Left(D(I), 30) + " " + msgtab(139)
   ![estkey] = XKEY
   ![catch] = C(I)
   ![eff] = E(I)
   ![Value] = V(I)
   ![FRNO] = FR(I)
      
   ![price] = 0
   If ![catch] <> 0 Then ![price] = ![Value] / ![catch]
   ![cpue] = 0
   If ![eff] <> 0 Then ![cpue] = ![catch] / ![eff]
   
   .Update
   
   GoTo NEXT_I
   
   End If
   
.Edit

![catch] = ![catch] + C(I)
![eff] = ![eff] + E(I)
![Value] = ![Value] + V(I)
![FRNO] = ![FRNO] + FR(I)
   
![price] = 0
If ![catch] <> 0 Then ![price] = ![Value] / ![catch]
![cpue] = 0
If ![eff] <> 0 Then ![cpue] = ![catch] / ![eff]
   
.Update

NEXT_I:

Next I

End With

prm_record.Close
prm_database.Close

End Sub
Private Sub CREATE_LOGS()

Dim dbn, CREC, VALF, NOFF

dbn = APPROOT + "\ARTBAS\RESULTS\WTOT.MDB"

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(dbn)
Set prm_record = prm_database.OpenRecordset("ESTAB")

With prm_record

Dim I, fnm, IX, XKEY, XXX, BGX

For I = 1 To 10000

If Len(RTrim(MNN(I))) = 0 Then GoTo NEXT_I

IX = Format(I, "0000")

fnm = APPROOT + "\ARTBAS\RESULTS\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_MN" + Format(I, "0000") + "_LOG.TXT"

Open fnm For Output As #1

Print #1, Tab(5); frmESTIM.Caption + " : " + msgtab(121)
Print #1, " "

Print #1, Tab(5); msgtab(122)
Print #1, Tab(5); String(80, "=")

'Start loop
'==========

.MoveFirst

Do Until .EOF

If Mid(![estkey], 2, 4) <> IX Then GoTo CONT_READ
If Mid(![estkey], 14, 4) <> "0000" Then GoTo CONT_READ
If Mid(![estkey], 8, 4) = "0000" Then GoTo CONT_READ

BGX = Mid(![estkey], 8, 4)

If ![catch] * ![eff] <> 0 Then XXX = ![estdes] + " " + msgtab(123)

If ![catch] * ![eff] = 0 Then XXX = ![estdes] + " " + msgtab(124)

Print #1, Tab(5); XXX
Print #1, " "

VALF = "Y": NOFF = "Y"

If ![Value] = 0 Then VALF = "N"
If ![fish] = 0 Then NOFF = "N"

If ![ltot] = 0 And ![actno] <> 0 Then Print #1, Tab(5); msgtab(125)
If ![esmp] = 0 And ![actno] <> 0 Then Print #1, Tab(5); msgtab(126)

If ![actno] = 0 Then Print #1, Tab(5); msgtab(128)
If ![FRNO] = 0 Then Print #1, Tab(5); msgtab(127)

If ![catch] * ![eff] = 0 Then
    'Print #1, Tab(5); String(80, ".")
    'Close #1
    CREC = .Bookmark
    GoTo SKIP_REST
    End If

If VALF = "N" And ![actno] <> 0 Then Print #1, Tab(5); msgtab(135) + " : " + msgtab(201)
'If NOFF = "N" And ![actno] <> 0 Then Print #1, Tab(5); msgtab(136) + " : " + msgtab(201)

If ![lsites] = 1 Then Print #1, Tab(5); msgtab(129)
If ![ldays] = 1 Then Print #1, Tab(5); msgtab(130)
If ![esites] = 1 Then Print #1, Tab(5); msgtab(131)
If ![edays] = 1 Then Print #1, Tab(5); msgtab(132)

'If ![BAC_ACCUR] < 0.9 Then Print #1, Tab(5); msgtab(133)
'If ![CPUE_ACCUR] < 0.9 Then Print #1, Tab(5); msgtab(134)

If ![bac_cv] = 0.5102 Then Print #1, Tab(5); msgtab(137)
If ![cpue_cv] = 0.5102 Then Print #1, Tab(5); msgtab(138)

CREC = .Bookmark

Print #1, " "

.MoveFirst

Do Until .EOF

If Mid(![estkey], 2, 4) <> IX Then GoTo CONT_READ2
If Mid(![estkey], 8, 4) <> BGX Then GoTo CONT_READ2
If Mid(![estkey], 14, 4) = "0000" Then GoTo CONT_READ2
If Mid(![estkey], 8, 4) = "0000" Then GoTo CONT_READ2

If ![price] = 0 And VALF = "Y" Then
   Print #1, Tab(5); msgtab(135) + " : " + ![estdes]
   End If

'If ![kgfish] = 0 And NOFF = "Y" Then
   'Print #1, Tab(5); msgtab(136) + " : " + ![estdes]
   'End If

CONT_READ2:

.MoveNext

Loop

SKIP_REST:

Print #1, Tab(5); String(80, "=")

.Bookmark = CREC

CONT_READ:

.MoveNext

Loop

Close #1

NEXT_I:

Next I

End With

prm_record.Close
prm_database.Close

End Sub
Private Sub JOIN_LOGS()

Dim I, fnm, IX, XKEY, XXX

Open APPROOT + "\ARTBAS\RESULTS\WLOG.TXT" For Output As #1

Print #1, Tab(5); frmESTIM.Caption + " : " + msgtab(121)
Print #1, " "

Print #1, Tab(5); msgtab(122)
Print #1, Tab(5); String(80, "=")

Close #1

Open APPROOT + "\ARTBAS\RESULTS\WLOG.TXT" For Append As #1

For I = 1 To 10000

If Len(RTrim(MNN(I))) = 0 Then GoTo NEXT_I

IX = Format(I, "0000")

fnm = APPROOT + "\ARTBAS\RESULTS\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_MN" + Format(I, "0000") + "_LOG.TXT"

If Dir(fnm) = "" Then GoTo NEXT_I:

Open fnm For Input As #2

Line Input #2, XXX
Line Input #2, XXX
Line Input #2, XXX
Line Input #2, XXX

Do Until EOF(2)

Line Input #2, XXX
Print #1, XXX

Loop

Close #2

NEXT_I:

Next I

Close #1

End Sub
Private Sub SPLIT_DATA()

Dim estdes, popn, smpn, BAC_ACCUR, FRNO, actno, cal, eact, _
          esmp, esites, edays, bac, bac_cvs, bac_cvsp, _
          bac_cvt, bac_cvtp, bac_cv, bac_low, bac_upper, eff, _
          eff_low, eff_upper, nland, CPUE_ACCUR, LPOP, ltot, _
          lsmpv, lsmpf, leff, cpue, lsites, ldays, _
          cpue_cvs, cpue_cvsp, cpue_cvt, cpue_cvtp, cpue_cv, cpue_low, _
          cpue_upper, catch, catch_low, catch_upper, catch_cv, Value, _
          price, fish, kgfish

Dim TYPEREC, ZERODATA, FOUND

Dim dbn, CREC

dbn = APPROOT + "\ARTBAS\RESULTS\WTOT.MDB"

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(dbn)
Set prm_record = prm_database.OpenRecordset("ESTAB")

With prm_record

Dim I, fnm, fnm2, fnm3, IX, XKEY, XXX

For I = 1 To 10000

If Len(RTrim(MNN(I))) = 0 Then GoTo NEXT_I

IX = Format(I, "0000"): FOUND = "N"

fnm = "XXX"

'Start loop
'==========

ZERODATA = "N"

.MoveFirst

Do Until .EOF

TYPEREC = 2

If Mid(![estkey], 2, 4) <> IX Then GoTo CONT_READ

FOUND = "Y"

If Mid(![estkey], 14, 4) = "0000" And Mid(![estkey], 8, 4) = "0000" Then
   If ![catch] * ![eff] = 0 Then ZERODATA = "Y"
   End If

If Mid(![estkey], 14, 4) = "0000" And Mid(![estkey], 8, 4) <> "0000" Then
   TYPEREC = 1
   End If

fnm = APPROOT + "\ARTBAS\RESULTS\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_MN" + Format(I, "0000") + "_ESTIM.TXT"

If Dir(fnm) = "" Then Open fnm For Output As #1
     
Write #1, ![estkey]

estdes = ![estdes]: popn = ![popn]: smpn = ![smpn]: BAC_ACCUR = ![BAC_ACCUR]
FRNO = ![FRNO]: actno = ![actno]: cal = ![cal]: eact = ![eact]
esmp = ![esmp]: esites = ![esites]: edays = ![edays]: bac = ![bac]
bac_cvs = ![bac_cvs]: bac_cvsp = ![bac_cvsp]
bac_cvt = ![bac_cvt]: bac_cvtp = ![bac_cvtp]: bac_cv = ![bac_cv]: bac_low = ![bac_low]
bac_upper = ![bac_upper]: eff = ![eff]: eff_low = ![eff_low]: eff_upper = ![eff_upper]
nland = ![nland]: CPUE_ACCUR = ![CPUE_ACCUR]: LPOP = ![LPOP]: ltot = ![ltot]
lsmpv = ![lsmpv]: lsmpf = ![lsmpf]: leff = ![leff]: cpue = ![cpue]: lsites = ![lsites]
ldays = ![ldays]: cpue_cvs = ![cpue_cvs]: cpue_cvsp = ![cpue_cvsp]: cpue_cvt = ![cpue_cvt]
cpue_cvtp = ![cpue_cvtp]: cpue_cv = ![cpue_cv]: cpue_low = ![cpue_low]
cpue_upper = ![cpue_upper]: catch = ![catch]: catch_low = ![catch_low]: catch_upper = ![catch_upper]
catch_cv = ![catch_cv]: Value = ![Value]: price = ![price]: fish = ![fish]: kgfish = ![kgfish]

If TYPEREC = 1 Then

Write #1, estdes, popn, smpn, BAC_ACCUR, FRNO, actno, cal, eact, _
          esmp, esites, edays, bac, bac_cvs, bac_cvsp, _
          bac_cvt, bac_cvtp, bac_cv, bac_low, bac_upper, eff, _
          eff_low, eff_upper, nland, CPUE_ACCUR, LPOP, ltot, _
          lsmpv, lsmpf, leff, cpue, lsites, ldays, _
          cpue_cvs, cpue_cvsp, cpue_cvt, cpue_cvtp, cpue_cv, cpue_low, _
          cpue_upper, catch, catch_low, catch_upper, catch_cv, Value, _
          price, fish, kgfish

         End If

If TYPEREC = 2 Then

Write #1, estdes, eff, cpue, catch, Value, price, kgfish, _
           fish, FRNO

         End If

CONT_READ:

.MoveNext

Loop

Close #1

If FOUND = "N" Then GoTo NEXT_I

If ZERODATA = "Y" And Dir(fnm) <> "" Then Kill fnm

NEXT_I:

Next I

End With

prm_record.Close
prm_database.Close

End Sub
