VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmFRAME 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10710
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
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
   ScaleHeight     =   7155
   ScaleWidth      =   10710
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBACK 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9840
      MousePointer    =   1  'Arrow
      Picture         =   "frmFRAME.frx":0000
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
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   960
      MousePointer    =   1  'Arrow
      Picture         =   "frmFRAME.frx":0282
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
      Picture         =   "frmFRAME.frx":0504
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
         Charset         =   178
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
      RecordSource    =   "FRTAB"
      Top             =   1560
      Width           =   2175
   End
   Begin MSDBGrid.DBGrid DBGGEN 
      Bindings        =   "frmFRAME.frx":22F6
      Height          =   5415
      Left            =   120
      OleObjectBlob   =   "frmFRAME.frx":2307
      TabIndex        =   0
      Top             =   480
      Width           =   10455
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FFFF&
      Caption         =   " 06"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   5
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
      Left            =   4200
      TabIndex        =   4
      Top             =   6120
      Width           =   5055
   End
End
Attribute VB_Name = "frmFRAME"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private NS, NBG, NOB
Private SCODE(), SNAME(), SSEQ(), BGCODE(), BGNAME(), BGSEQ()
Private NF, FD(), FN(), FTC(), FTD(), FTN(), FTQ(), NEDIT
Private FIND_STRING, WKEY_CODE, WKEY_CHAR
Private Sub cmdBACK_Click()

Dim resp, dbn, fnm

resp = MsgBox(msgtab(60), vbOKCancel, " ")

If resp = 2 Then Exit Sub

dbn = APPROOT + "\ARTBAS\TABLES\FRAME.MDB"

dtaGEN.DatabaseName = APPROOT + "\ARTBAS\STRUS\FRAUX.MDB"
dtaGEN.Refresh

If Dir(dbn) <> "" Then Kill dbn

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_FRAME.TXT"

If NOB = 0 Then Kill fnm

frmFRAME.MousePointer = 13
Load frmTABLES
Unload frmFRAME
frmTABLES.Show

End Sub
Private Sub cmdEND_Click()

Dim fnm

frmFRAME.MousePointer = 13

Call CREATE_FTAB
Call DUMP_FRAME


fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_FRAME.TXT"

If NOB = 0 Then Kill fnm

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_WFRAME.TXT"

If Dir(fnm) <> "" Then Kill fnm

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_WACTIVE.TXT"

Open fnm For Output As #1
Close #1

Load frmTABLES
Unload frmFRAME
frmTABLES.Show

End Sub
Private Sub CREATE_FTAB()

Dim XXX, dbn As String

dbn = APPROOT + "\ARTBAS\TABLES\FRAME.MDB"

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(dbn)
Set prm_record = prm_database.OpenRecordset("FRTAB")

Dim I, K

With prm_record

.MoveFirst
.Index = "primarykey"

For I = 1 To NEDIT

.Seek "=", FTC(I)

If .NoMatch = True Then
   Debug.Print FTC(I)
   End
   End If

.Edit

![FNO] = FTN(I)

.Update

Next I

End With

prm_record.Close
prm_database.Close

End Sub
Private Sub cmdPRINT_Click()

Dim dbn

dbn = APPROOT + "\ARTBAS\TABLES\FRAME.MDB"

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(dbn)
Set prm_record = prm_database.OpenRecordset("FRTAB")

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

Printer.Print Tab(5); ![fdescr] + " " + Format(![FNO], "#####0.00")

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

Printer.Print Tab(5); frmFRAME.Caption

Printer.Print

Printer.Print Tab(5); dbgGEN.Columns(1).Caption; _
              Tab(66); dbgGEN.Columns(2).Caption

Printer.Print Tab(5); String(70, "-")

Return

End Sub
Private Sub DBGGEN_AfterColEdit(ByVal ColIndex As Integer)

Dim vvv

vvv = dbgGEN.Columns(2).Value

If IsNumeric(vvv) = False Then dbgGEN.Columns(2).Value = 0

Dim I

NEDIT = NEDIT + 1: I = NEDIT

ReDim Preserve FTC(1 To NEDIT), FTN(1 To NEDIT)

FTC(I) = dbgGEN.Columns(0).Value
FTN(I) = dbgGEN.Columns(2).Value

End Sub
Private Sub dbgGEN_Click()

FIND_STRING = ""

End Sub
Private Sub DBGGEN_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn And dbgGEN.Col = 1 Then
   dbgGEN.Col = 1
   FIND_STRING = ""
   Exit Sub
   End If

If dbgGEN.Col <> 1 Then
   FIND_STRING = ""
   Exit Sub
   End If
 
WKEY_CODE = KeyCode

If KeyCode = vbKeyPageUp Or _
   KeyCode = vbKeyPageDown Or _
   KeyCode = vbKeyEnd Or _
   KeyCode = vbKeyHome Or _
   KeyCode = vbKeyLeft Or _
   KeyCode = vbKeyUp Or _
   KeyCode = vbKeyRight Or _
   KeyCode = vbKeyDown Then
   
   FIND_STRING = ""
   
   Exit Sub
   
   End If

Call FIND_CHAR_STR

End Sub
Private Sub FIND_CHAR_STR()

Dim CURX, KEYX, I, XXX, yyy, FIND_FLAG, POS_CUR, POS_KEY, DIFF_POS, BMKC, L, WM

KEYX = WKEY_CODE

FIND_STRING = FIND_STRING + UCase(Chr(KEYX))

If Len(FIND_STRING) >= 1 Then GoSub FIND_CHAR_KEY

If POS_KEY = 0 Then
   FIND_STRING = ""
   dbgGEN.Col = 2
   Exit Sub
   End If
   
POS_KEY = POS_KEY - 1

dtaGEN.Recordset.MoveFirst

For I = 1 To POS_KEY

dtaGEN.Recordset.MoveNext

Next I

Exit Sub

FIND_CHAR_KEY:

POS_KEY = 0

FIND_STRING = UCase(FIND_STRING)
FIND_STRING = LTrim(RTrim(FIND_STRING))
L = Len(FIND_STRING)

Dim FCHR

FCHR = Left(FIND_STRING, 1)

For I = 1 To NF
XXX = UCase(FD(I))

If FCHR <> "!" Then
   If Left(XXX, L) = FIND_STRING Then
      POS_KEY = I
      Return
      End If
GoTo NEXT_I2
End If

If FCHR = "!" Then
   WM = InStr(XXX, Right(FIND_STRING, L - 1))
   If WM <> 0 Then
   POS_KEY = I
   Return
   End If
GoTo NEXT_I2
End If

NEXT_I2:

Next I

FIND_STRING = ""

Return

End Sub
Private Sub DBGGEN_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

Exit Sub

With dbgGEN

If .Col = 1 Then .Col = 2

End With

End Sub
Private Sub Form_Load()

dtaGEN.DatabaseName = APPROOT + "\ARTBAS\STRUS\Fraux.mdb"

Set Picture = LoadPicture(APPROOT + "\ARTBAS\PICS_RUNTIME\SCREEN_06.JPG")

NEDIT = 0

frmFRAME.Caption = monthtab(current_month) + " " + _
                    Format(current_year, "0000") + " - " + msgtab(46)

lblUPD.Visible = False
dtaGEN.Visible = False
dbgGEN.Visible = True

dbgGEN.Columns(1).Caption = msgtab(66)
dbgGEN.Columns(2).Caption = msgtab(67)

cmdBACK.ToolTipText = msgtab(49)
cmdEND.ToolTipText = msgtab(51)
cmdPRINT.ToolTipText = msgtab(52)
cmdBACK.ToolTipText = msgtab(60)

lblUPD.Caption = msgtab(58)

Call LOAD_SITES
Call LOAD_BG
Call SETUP_FRAME

End Sub
Private Sub SETUP_FRAME()

Dim fnm

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_FRAME.TXT"
     
If Dir(fnm) <> "" Then
   Call CLEAN_FRAME
   Call INSERT_FRAME
   Call MOVE_TODB
   Exit Sub
   End If
     
If Dir(fnm) = "" Then
   Call SETUP_NEW
   Call MOVE_TODB
   Exit Sub
   End If

End Sub
Private Sub LOAD_SITES()

Dim I, XXX, fnm

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_SITES.TXT"

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
Private Sub SETUP_NEW()

Dim fnm, I, J

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_FRAME.TXT"

Open fnm For Output As #1

For I = 1 To NS
For J = 1 To NBG

Print #1, "S" + Format(SCODE(I), "0000") + _
          "+" + "B" + Format(BGCODE(J), "0000") + _
          " " + SSEQ(I) + BGSEQ(J) + "0000000000"

Next J
Next I

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

dtaGEN.DatabaseName = APPROOT + "\ARTBAS\STRUS\FRAUX.MDB"
dtaGEN.Refresh

Dim fnm

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_FRAME.TXT"

Open fnm For Input As #1

Dim XXX, yyy

Dim dbn As String

FileCopy APPROOT + "\ARTBAS\STRUS\FRAME.MDB", APPROOT + "\ARTBAS\TABLES\FRAME.MDB"

dbn = APPROOT + "\ARTBAS\TABLES\FRAME.MDB"

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(dbn)
Set prm_record = prm_database.OpenRecordset("FRTAB")

With prm_record

.Index = "SORTKEY"

NF = 0

Do Until EOF(1)

Line Input #1, XXX

.AddNew

![fkey] = Left(XXX, 11)
![FNO] = CDbl(Mid(XXX, 26, 15))

I = CDbl(Mid(XXX, 2, 4)): J = CDbl(Mid(XXX, 8, 4))

![SortKey] = ss(I) + bs(J)

![fdescr] = sn(I) + " " + bn(J)

.Update

NF = NF + 1

ReDim Preserve FD(1 To NF)
FD(NF) = sn(I) + " " + bn(J)

Loop

Close #1

End With

prm_record.Close
prm_database.Close

dtaGEN.DatabaseName = APPROOT + "\ARTBAS\TABLES\FRAME.MDB"
dtaGEN.Refresh

dbgGEN.Visible = True

End Sub
Private Sub DUMP_FRAME()

Dim fnm

NOB = 0

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_FRAME.TXT"

Open fnm For Output As #1

Dim XXX, dbn As String

dbn = APPROOT + "\ARTBAS\TABLES\FRAME.MDB"

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(dbn)
Set prm_record = prm_database.OpenRecordset("FRTAB")

With prm_record

.MoveFirst
.Index = "sortkey"

Do Until .EOF

Print #1, ![fkey] + " " + Left(![SortKey] + Space(12), 12) + " " + _
        Format(![FNO], "#####0.00")

NOB = NOB + ![FNO]

.MoveNext

Loop

End With

Close #1

prm_record.Close
prm_database.Close

dtaGEN.DatabaseName = APPROOT + "\ARTBAS\STRUS\FRAUX.MDB"
dtaGEN.Refresh

Kill dbn

End Sub
Private Sub CLEAN_FRAME()

Dim fnm, I, J, XXX, vvv

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_FRAME.TXT"
      
      
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
Private Sub INSERT_FRAME()

Dim XKEY, I, J

Call MOVE_TODB

dtaGEN.DatabaseName = APPROOT + "\ARTBAS\STRUS\FRAUX.MDB"
dtaGEN.Refresh

Dim dbn As String

dbn = APPROOT + "\ARTBAS\TABLES\FRAME.MDB"

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(dbn)
Set prm_record = prm_database.OpenRecordset("FRTAB")

With prm_record

prm_record.Index = "primarykey"

For I = 1 To NS
For J = 1 To NBG

XKEY = "S" + Format(SCODE(I), "0000") + "+" + "B" + Format(BGCODE(J), "0000")

prm_record.Seek "=", XKEY

If prm_record.NoMatch = False Then GoTo CONT_LOOP

.AddNew

![fkey] = XKEY
![FNO] = 0
![fdescr] = SNAME(I) + " " + BGNAME(J)
![SortKey] = SSEQ(I) + BGSEQ(J)

.Update

CONT_LOOP:

Next J
Next I

End With

prm_record.Close
prm_database.Close

Call DUMP_FRAME

End Sub
