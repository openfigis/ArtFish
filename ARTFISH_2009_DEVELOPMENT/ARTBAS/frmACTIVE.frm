VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmACTIVE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10710
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   10710
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGUIDE 
      Appearance      =   0  'Flat
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
      Height          =   735
      Left            =   9000
      MousePointer    =   1  'Arrow
      Picture         =   "frmACTIVE.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6240
      Width           =   735
   End
   Begin VB.CommandButton cmdBACK 
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
      Height          =   735
      Left            =   9840
      MousePointer    =   1  'Arrow
      Picture         =   "frmACTIVE.frx":2262
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6240
      Width           =   735
   End
   Begin VB.CommandButton cmdPRINT 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   960
      MousePointer    =   1  'Arrow
      Picture         =   "frmACTIVE.frx":24E4
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6240
      Width           =   735
   End
   Begin VB.CommandButton cmdEND 
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
      Picture         =   "frmACTIVE.frx":2766
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6240
      Width           =   735
   End
   Begin VB.Data dtaGEN 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "ACTAB"
      Top             =   1560
      Width           =   2175
   End
   Begin MSDBGrid.DBGrid DBGGEN 
      Bindings        =   "frmACTIVE.frx":4558
      Height          =   5415
      Left            =   120
      OleObjectBlob   =   "frmACTIVE.frx":4599
      TabIndex        =   0
      Top             =   480
      Width           =   10455
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FFFF&
      Caption         =   " 07"
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
      TabIndex        =   6
      Top             =   6960
      Width           =   255
   End
   Begin VB.Label lblUPD 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   4080
      TabIndex        =   4
      Top             =   6120
      Width           =   5175
   End
End
Attribute VB_Name = "frmACTIVE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' TNJ Note - all the code subroutines in this form are alphabetized by name

Private NS, NBG, NOB
Private SCODE(), SNAME(), SSEQ(), BGSEQ(), BGCODE(), BGNAME()
Private NF, FD(), FN(), FTC(), FTN(), NEDIT, ASO()

Private Sub CHECK_FRNOS()

Dim I, J, K, XXX, yyy, fnm, XKEY

Dim dbn As String

dbn = APPROOT + "\ARTBAS\TABLES\ACTIVE.MDB"

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(dbn)
Set prm_record = prm_database.OpenRecordset("ACTAB")

With prm_record

.Index = "primarykey"

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_FRAME.TXT"

Open fnm For Input As #1

Do Until EOF(1)

Line Input #1, XXX

J = Val(Mid(XXX, 2, 4))

K = ASO(J)

XKEY = "S" + Format(K, "0000") + Mid(XXX, 6, 6)

.Seek "=", XKEY

If .NoMatch = True Then GoTo next_rec

.Edit

![afr] = ![afr] + CDbl(Mid(XXX, 26, 15))

.Update

next_rec:

Loop

Close #1

.MoveFirst

Do Until .EOF

If ![afr] = 0 Then .Delete

.MoveNext

Loop

prm_record.Close
prm_database.Close

End With

End Sub
Private Sub CLEAN_ACTIVE()

Dim fnm, I, J, XXX, vvv

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_ACTIVE.TXT"
      
      
Dim SF(), BF()

ReDim SF(1 To 10000), BF(1 To 10000)

For I = 1 To 10000
SF(I) = 0
Next I

For I = 1 To 10000
BF(I) = 0
Next I

For I = 1 To NS
J = SCODE(I): SF(J) = 1
Next I

For I = 1 To NBG
J = BGCODE(I): BF(J) = 1
Next I

Open APPROOT + "\ARTBAS\TABLES\WORK.TXT" For Output As #2
Open fnm For Input As #1

Do Until EOF(1)

Line Input #1, XXX

I = CDbl(Mid(XXX, 2, 4))

If SF(I) <> 1 Then GoTo CONT_LOOP

I = CDbl(Mid(XXX, 8, 4))

If BF(I) <> 1 Then GoTo CONT_LOOP

Print #2, XXX

CONT_LOOP:

Loop

Close #1
Close #2

FileCopy APPROOT + "\artbas\tables\work.txt", fnm

End Sub
Private Sub cmdBACK_Click()

Dim resp, dbn

resp = MsgBox(msgtab(60), vbOKCancel, " ")

If resp = 2 Then Exit Sub

dbn = APPROOT + "\ARTBAS\TABLES\ACTIVE.MDB"

dtaGEN.DatabaseName = APPROOT + "\ARTBAS\STRUS\ACTAUX.MDB"
dtaGEN.Refresh

If Dir(dbn) <> "" Then Kill dbn

Dim fnm

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_ACTIVE.TXT"

If NOB = 0 Then Kill fnm

If Dir(APPROOT + "\ARTBAS\TABLES\WORK.TXT") <> "" Then Kill APPROOT + "\ARTBAS\TABLES\WORK.TXT"

frmACTIVE.MousePointer = 13
Load frmARTB01
Unload frmACTIVE
frmARTB01.Show

End Sub
Private Sub cmdEND_Click()

frmACTIVE.MousePointer = 13

Call CREATE_FTAB
Call DUMP_ACTIVE

Dim fnm

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_ACTIVE.TXT"

If NOB = 0 Then Kill fnm

If Dir(APPROOT + "\ARTBAS\TABLES\WORK.TXT") <> "" Then Kill APPROOT + "\ARTBAS\TABLES\WORK.TXT"

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_WACTIVE.TXT"

If Dir(fnm) <> "" Then Kill fnm

Load frmARTB01
Unload frmACTIVE
frmARTB01.Show

End Sub

Private Sub cmdGUIDE_Click()

HTYPE = "B0"

HFNM = APPROOT + "\ARTBAS\HELP\" + current_language + "HELP" + HTYPE + ".rtf"

If Dir(HFNM) = "" Then Exit Sub

frmACTIVE.Enabled = False
Load frmGUIDE
frmGUIDE.Show

End Sub

Private Sub cmdPRINT_Click()

Dim dbn

dbn = APPROOT + "\ARTBAS\TABLES\ACTIVE.MDB"

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(dbn)
Set prm_record = prm_database.OpenRecordset("ACTAB")

Printer.FontBold = True
Printer.FontName = "Courier"
Printer.FontName = "Courier New"
Printer.FontSize = 11

Dim I, J, pageno, lineno

pageno = 0

GoSub CHANGE_PAGE

With prm_record

.MoveFirst
.Index = "sortkey"

Do Until .EOF

Printer.Print Tab(5); ![adescr] + " " + Format(![aNO], "#####0.00")

lineno = lineno + 1

If lineno > 55 Then GoSub CHANGE_PAGE

.MoveNext

Loop

End With

prm_record.Close
prm_database.Close

Printer.EndDoc

Exit Sub

'========================
CHANGE_PAGE:

lineno = 0
pageno = pageno + 1
If pageno > 1 Then Printer.NewPage

Printer.Print

Printer.Print Tab(5); frmACTIVE.Caption

Printer.Print

Printer.Print Tab(5); dbgGEN.Columns(1).Caption; _
              Tab(66); dbgGEN.Columns(2).Caption

Printer.Print Tab(5); String(70, "-")

Return

End Sub
Private Sub CREATE_FTAB()

Dim XXX, dbn As String

dbn = APPROOT + "\ARTBAS\TABLES\ACTIVE.MDB"

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(dbn)
Set prm_record = prm_database.OpenRecordset("ACTAB")

Dim I, K

With prm_record

NF = .RecordCount

.MoveFirst
.Index = "primarykey"

For I = 1 To NEDIT

.Seek "=", FTC(I)

If .NoMatch = True Then
   Debug.Print FTC(I)
   End
   End If

.Edit

![aNO] = FTN(I)

.Update

Next I

End With

prm_record.Close
prm_database.Close

End Sub
Private Sub DBGGEN_AfterColEdit(ByVal ColIndex As Integer)

Dim vvv

vvv = dbgGEN.Columns(2).Value

If IsNumeric(vvv) = False Then dbgGEN.Columns(2).Value = 0

If CDbl(dbgGEN.Columns(2).Value) > CURCAL Then dbgGEN.Columns(2).Value = CURCAL

Dim I

NEDIT = NEDIT + 1: I = NEDIT

ReDim Preserve FTC(1 To NEDIT), FTN(1 To NEDIT)

FTC(I) = dbgGEN.Columns(0).Value
FTN(I) = dbgGEN.Columns(2).Value

End Sub
Private Sub DBGGEN_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)

Dim vvv

vvv = dbgGEN.Columns(2).Value

If IsNumeric(vvv) = False Then dbgGEN.Columns(2).Value = 0

End Sub

Private Sub DBGGEN_KeyPress(KeyAscii As Integer)

On Error GoTo EXIT_SUB

Dim I, J, dbn

With dbgGEN

If KeyAscii = vbKeyReturn And .Row <> 0 Then
   I = .Row: J = .Col
   .Col = 1
   .Row = .Row - 1
   .Refresh
   .Row = I: .Col = J
   .Refresh
   End If
   
If KeyAscii = vbKeyReturn And .Row = 0 Then

   I = 0: J = .Col
   .Col = 1
   .Row = 1
   .Refresh
   .Row = I: .Col = J
   .Refresh
   End If
  
End With

EXIT_SUB:

End Sub
Private Sub DBGGEN_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

With dbgGEN

If .Col = 1 Then .Col = 2

End With

End Sub
Private Sub DUMP_ACTIVE()

NOB = 0

Dim fnm

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_ACTIVE.TXT"

Open fnm For Output As #1

Dim XXX, dbn As String

dbn = APPROOT + "\ARTBAS\TABLES\ACTIVE.MDB"

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(dbn)
Set prm_record = prm_database.OpenRecordset("ACTAB")

With prm_record

.MoveFirst
.Index = "sortkey"

Do Until .EOF

Print #1, ![akey] + " " + Left(![SortKey] + Space(12), 12) + " " + Format(![aNO], "#####0.00")

NOB = NOB + ![aNO]

.MoveNext

Loop

End With

Close #1

prm_record.Close
prm_database.Close

dtaGEN.DatabaseName = APPROOT + "\ARTBAS\STRUS\ACTAUX.MDB"
dtaGEN.Refresh

Kill dbn

End Sub
Private Sub Form_Load()

dtaGEN.DatabaseName = APPROOT + "\ARTBAS\STRUS\ACTAUX.MDB"

Set Picture = LoadPicture(APPROOT + "\ARTBAS\PICS_RUNTIME\SCREEN_07.JPG")

NEDIT = 0

frmACTIVE.Caption = monthtab(current_month) + " " + _
                    Format(current_year, "0000") + " - " + msgtab(35)

lblUPD.Visible = False
dtaGEN.Visible = False
dbgGEN.Visible = True

dbgGEN.Columns(1).Caption = msgtab(68)
dbgGEN.Columns(2).Caption = msgtab(69)

cmdBACK.ToolTipText = msgtab(49)
cmdEND.ToolTipText = msgtab(51)
cmdPRINT.ToolTipText = msgtab(52)
cmdBACK.ToolTipText = msgtab(60)
cmdGUIDE.ToolTipText = msgtab(243)

lblUPD.Caption = msgtab(58)

Call LOAD_ASSO
Call LOAD_STRATA
Call LOAD_BG
Call SETUP_ACTIVE

End Sub
Private Sub INSERT_ACTIVE()

Dim XKEY, I, J

Call MOVE_TODB

dtaGEN.DatabaseName = APPROOT + "\ARTBAS\STRUS\ACTAUX.MDB"
dtaGEN.Refresh

Dim dbn As String

dbn = APPROOT + "\ARTBAS\TABLES\ACTIVE.MDB"

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(dbn)
Set prm_record = prm_database.OpenRecordset("ACTAB")

With prm_record

.Index = "primarykey"

For I = 1 To NS
For J = 1 To NBG

XKEY = "S" + Format(SCODE(I), "0000") + "+" + "B" + Format(BGCODE(J), "0000")

prm_record.Seek "=", XKEY

If prm_record.NoMatch = False Then GoTo CONT_LOOP

.AddNew

![akey] = XKEY
![aNO] = 0
![adescr] = SNAME(I) + " " + BGNAME(J)
![SortKey] = SSEQ(I) + BGSEQ(J)

.Update

CONT_LOOP:

Next J
Next I

End With

prm_record.Close
prm_database.Close

Call DUMP_ACTIVE

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
Private Sub LOAD_BG()

Dim I, XXX, fnm

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_BG.TXT"

NBG = 0

Open fnm For Input As #1

Do Until EOF(1)

Line Input #1, XXX

NBG = NBG + 1

ReDim Preserve BGNAME(1 To NBG), BGCODE(1 To NBG), BGSEQ(1 To NBG)

BGCODE(NBG) = CDbl(Mid(XXX, 1, 4))
BGNAME(NBG) = Mid(XXX, 6, 30)
BGSEQ(NBG) = Mid(XXX, 69, 6)

Loop

Close #1

End Sub
Private Sub LOAD_STRATA()

Dim I, XXX, fnm

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_MINOR.TXT"

NS = 0

Open fnm For Input As #1

Do Until EOF(1)

Line Input #1, XXX

NS = NS + 1

ReDim Preserve SNAME(1 To NS), SCODE(1 To NS), SSEQ(1 To NS)

SCODE(NS) = CDbl(Mid(XXX, 1, 4))
SNAME(NS) = Mid(XXX, 6, 30)
SSEQ(NS) = Mid(XXX, 69, 6)

Loop

Close #1

End Sub
Private Sub MOVE_TODB()

Dim sn(), bn(), I, J, ss(), bs()

ReDim sn(1 To 10000), bn(1 To 10000), ss(1 To 10000), bs(1 To 10000)

For I = 1 To NS
sn(SCODE(I)) = SNAME(I)
ss(SCODE(I)) = SSEQ(I)
Next I

For I = 1 To NBG
bn(BGCODE(I)) = BGNAME(I)
bs(BGCODE(I)) = BGSEQ(I)
Next I

dtaGEN.DatabaseName = APPROOT + "\ARTBAS\STRUS\ACTAUX.MDB"
dtaGEN.Refresh

Dim fnm

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_ACTIVE.TXT"

Open fnm For Input As #1

Dim XXX, yyy

Dim dbn As String

FileCopy APPROOT + "\ARTBAS\STRUS\ACTIVE.MDB", APPROOT + "\ARTBAS\TABLES\ACTIVE.MDB"

dbn = APPROOT + "\ARTBAS\TABLES\ACTIVE.MDB"

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(dbn)
Set prm_record = prm_database.OpenRecordset("ACTAB")

With prm_record

.Index = "sortkey"

Do Until EOF(1)

Line Input #1, XXX

.AddNew

![akey] = Left(XXX, 11)
![aNO] = CDbl(Mid(XXX, 26, 15))

If ![aNO] > CURCAL Then ![aNO] = CURCAL

I = CDbl(Mid(XXX, 2, 4)): J = CDbl(Mid(XXX, 8, 4))

![adescr] = sn(I) + " " + bn(J)
![SortKey] = ss(I) + bs(J)

.Update

Loop

Close #1

End With

prm_record.Close
prm_database.Close

Call CHECK_FRNOS

dtaGEN.DatabaseName = APPROOT + "\ARTBAS\TABLES\ACTIVE.MDB"
dtaGEN.Refresh

dbgGEN.Visible = True

End Sub
Private Sub SETUP_ACTIVE()
Dim fnm

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_ACTIVE.TXT"
     
If Dir(fnm) <> "" Then
   Call CLEAN_ACTIVE
   Call INSERT_ACTIVE
   Call MOVE_TODB
   Exit Sub
   End If
     
If Dir(fnm) = "" Then
   Call SETUP_NEW
   Call MOVE_TODB
   Exit Sub
   End If

End Sub
Private Sub SETUP_NEW()

Dim fnm, I, J

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_ACTIVE.TXT"

Open fnm For Output As #1

For I = 1 To NS
For J = 1 To NBG

Print #1, "S" + Format(SCODE(I), "0000") + "+" + "B" + _
                Format(BGCODE(J), "0000") + _
                " " + SSEQ(I) + BGSEQ(J) + "000000000"

Next J
Next I

Close #1

Call MOVE_TODB

End Sub
