VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmDB01 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9195
   BeginProperty Font 
      Name            =   "Arial Narrow"
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
   ScaleHeight     =   3810
   ScaleWidth      =   9195
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ProgressBar pgbFILES 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3240
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2880
      Width           =   8895
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FFFF&
      Caption         =   " 02"
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
      TabIndex        =   0
      Top             =   7560
      Width           =   255
   End
End
Attribute VB_Name = "frmDB01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private SITE_NAME, SITE_MIN, SITE_MAJ
Private LISTMY(), NS, NSTRUS, READY_ERROR, EMPTY_ERROR
Private MY_ERROR, BKP_ERROR, IMP_ERROR, PROCEED_FLAG
Private RESTORE_FLAG
Private ASSOSIMNC(1 To 10000), ASSOSIMND(1 To 10000)
Private MNC(1 To 10000), MND(1 To 10000), SIC(1 To 10000), SID(1 To 10000), BGC(1 To 10000), BGD(1 To 10000)
Private SPEC(1 To 10000), SPED(1 To 10000)
Private LOAD_TABLES_FLAG, LANG_IND
Private YYI, MMI, NYT, YMTAB()
Private Sub Form_Load()

Dim XXX

Open APPROOT + "\ARTBAS\CONTROL\SYSPARM.TXT" For Input As #1

Input #1, language
current_language = language
Close #1

Call MSGLOAD

language = "FRENCH": current_language = language

Call MAIN_ROUTINE

End Sub
Private Sub MAIN_ROUTINE()

Dim XXX

Open APPROOT + "\ARTBAS\CONTROL\SELYM.TXT" For Input As #1

NYT = 0
Do Until EOF(1)

Line Input #1, XXX
NYT = NYT + 1

ReDim Preserve YMTAB(1 To NYT)

YMTAB(NYT) = RTrim(XXX)

Loop

Close #1

frmDB01.Show
frmDB01.Refresh

Open APPROOT + "\ARTBAS\CONTROL\YEARMONTH.TXT" For Input As #1

Input #1, current_year, current_month

Close #1

YYI = current_year: MMI = current_month

pgbFILES.Min = 0: pgbFILES.Max = 13 * NYT
pgbFILES.Visible = True

frmDB01.Refresh

NSTRUS = 0

Dim I, II As Integer

ReDim monthtab(1 To 12)

For I = 1 To 12
monthtab(I) = msgtab(I + 17)
Next I

frmDB01.Caption = msgtab(40) + " - " + monthtab(current_month) + " " + _
                    Format(current_year, "0000")

Dim resp

pgbFILES.Visible = True
pgbFILES.Max = NYT
pgbFILES.Min = 0
pgbFILES.Value = 0

resp = MsgBox(msgtab(277), vbYesNoCancel, " ")
If resp = vbCancel Then End

LANG_IND = "LOCAL"

If resp = vbYes Then LANG_IND = "LOCAL"
If resp = vbNo Then LANG_IND = "INTER"

'====================== LOOP =====================

For II = 1 To NYT

MMI = Val(Mid(YMTAB(II), 8, 2))
YYI = Val(Left(YMTAB(II), 4))

DBN2 = APPROOT + "\ARTBAS\EXPORT\" + "Y" + Format(YYI, "0000") + "M" + _
       Format(MMI, "00") + "_ALLDATA.MDB"

DBN = APPROOT + "\ARTBAS\STRUS\ALLDATA.MDB"
FileCopy DBN, DBN2

'============================================================


If Dir(DBN) = "" Then

   resp = MsgBox(msgtab(281), vbOKOnly, " ")
   GoTo DO_NOTHING

End If

Dim fnm, CREATE_FLAG

CREATE_FLAG = "NO"

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(YYI, "0000") + _
      "M" + Format(MMI, "00") + "_MAJOR.TXT"

If Dir(fnm) <> "" Then CREATE_FLAG = "YES"

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(YYI, "0000") + _
      "M" + Format(MMI, "00") + "_MINOR.TXT"

If Dir(fnm) <> "" Then CREATE_FLAG = "YES"

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(YYI, "0000") + _
      "M" + Format(MMI, "00") + "_ASSOMN.TXT"

If Dir(fnm) <> "" Then CREATE_FLAG = "YES"

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(YYI, "0000") + _
      "M" + Format(MMI, "00") + "_SITES.TXT"

If Dir(fnm) <> "" Then CREATE_FLAG = "YES"

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(YYI, "0000") + _
      "M" + Format(MMI, "00") + "_ASSOSI.TXT"

If Dir(fnm) <> "" Then CREATE_FLAG = "YES"

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(YYI, "0000") + _
      "M" + Format(MMI, "00") + "_BG.TXT"

If Dir(fnm) <> "" Then CREATE_FLAG = "YES"

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(YYI, "0000") + _
      "M" + Format(MMI, "00") + "_SPECIES.TXT"

If Dir(fnm) <> "" Then CREATE_FLAG = "YES"

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(YYI, "0000") + _
      "M" + Format(MMI, "00") + "_FRAME.TXT"

If Dir(fnm) <> "" Then CREATE_FLAG = "YES"

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(YYI, "0000") + _
      "M" + Format(MMI, "00") + "_ACTIVE.TXT"

If Dir(fnm) <> "" Then CREATE_FLAG = "YES"

'====================================================================================

If CREATE_FLAG = "NO" Then
resp = MsgBox(msgtab(279), vbOKOnly, " ")
GoTo DO_NOTHING
End If

Call CRDB_MAJOR_STRATA
Call CRDB_SPECIES
Call CRDB_BOATS_GEARS
Call CRDB_MINOR_STRATA
Call CRDB_SITES
Call CRDB_ASSOCIATIONS
Call CRDB_ACTIVE_DAYS
Call CRDB_FRAME_SURVEY
Call CRDB_EFFORT
Call CRDB_ESTIM_REMARKS
Call CRDB_TRIP_TOTALS
Call CRDB_BY_SPECIES
Call CRDB_RESULTS

pgbFILES.Visible = True
pgbFILES.Value = II

Label2.Caption = DBN2
Label2.Refresh

Next II

resp = MsgBox(msgtab(278) + " " + msgtab(280), vbOKOnly, " ")

Call APPEND_DATABASES

DO_NOTHING:

End

End Sub
Private Sub CRDB_MAJOR_STRATA()

Dim fnm, XXX, xcode, xmnc, xact, xsmp, xfrm, xrec

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(YYI, "0000") + _
      "M" + Format(MMI, "00") + "_MAJOR.TXT"

If Dir(fnm) = "" Then Exit Sub

Open fnm For Input As #1

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(DBN2)
Set prm_record = prm_database.OpenRecordset("A_MAJOR_STRATA")

With prm_record

.Index = "primarykey"

Do Until EOF(1)

Line Input #1, XXX

.AddNew

![COMPOSITE_CODE] = Format(YYI, "0000") + "+" + Format(MMI, "00") + "+" + Left(XXX, 4)
![CODE] = Left(XXX, 4)
![NAME_1] = Mid(XXX, 6, 30)
![NAME_2] = Mid(XXX, 37, 31)
![SORT_SEQ] = Mid(XXX, 69, 6)

.Update

Loop

Close #1

End With

prm_record.Close
prm_database.Close

End Sub
Private Sub CRDB_SPECIES()

Dim fnm, XXX, xcode, xmnc, xact, xsmp, xfrm, xrec

DBN2 = APPROOT + "\ARTBAS\EXPORT\" + "Y" + Format(YYI, "0000") + "M" + Format(MMI, "00") + "_ALLDATA.MDB"

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(YYI, "0000") + _
      "M" + Format(current_month, "00") + "_SPECIES.TXT"

If Dir(fnm) = "" Then Exit Sub

Open fnm For Input As #1

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(DBN2)
Set prm_record = prm_database.OpenRecordset("F_SPECIES")

With prm_record

.Index = "primarykey"

Do Until EOF(1)

Line Input #1, XXX

.AddNew

![COMPOSITE_CODE] = Format(YYI, "0000") + "+" + Format(MMI, "00") + "+" + Left(XXX, 4)
![CODE] = Left(XXX, 4)
![NAME_1] = Mid(XXX, 6, 30)
![NAME_2] = Mid(XXX, 37, 31)
![SORT_SEQ] = Mid(XXX, 69, 6)

.Update

Loop

Close #1

End With

prm_record.Close
prm_database.Close

End Sub
Private Sub CRDB_BOATS_GEARS()

Dim fnm, DBN, DBN2, XXX, xcode, xmnc, xact, xsmp, xfrm, xrec

DBN2 = APPROOT + "\ARTBAS\EXPORT\" + "Y" + Format(YYI, "0000") + "M" + Format(MMI, "00") + "_ALLDATA.MDB"

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(YYI, "0000") + _
      "M" + Format(MMI, "00") + "_BG.TXT"

If Dir(fnm) = "" Then Exit Sub

Open fnm For Input As #1

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(DBN2)
Set prm_record = prm_database.OpenRecordset("E_BOATS_GEARS")

With prm_record

.Index = "primarykey"

Do Until EOF(1)

Line Input #1, XXX

.AddNew

![COMPOSITE_CODE] = Format(YYI, "0000") + "+" + Format(MMI, "00") + "+" + Left(XXX, 4)
![CODE] = Left(XXX, 4)
![NAME_1] = Mid(XXX, 6, 30)
![NAME_2] = Mid(XXX, 37, 31)
![SORT_SEQ] = Mid(XXX, 69, 6)

.Update

Loop

Close #1

End With

prm_record.Close
prm_database.Close

End Sub
Private Sub CRDB_MINOR_STRATA()

Dim fnm, DBN, DBN2, XXX, xcode, xmnc, xact, xsmp, xfrm, xrec

DBN2 = APPROOT + "\ARTBAS\EXPORT\" + "Y" + Format(YYI, "0000") + "M" + Format(MMI, "00") + "_ALLDATA.MDB"

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(YYI, "0000") + _
      "M" + Format(MMI, "00") + "_MINOR.TXT"

If Dir(fnm) = "" Then Exit Sub

Open fnm For Input As #1

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(DBN2)
Set prm_record = prm_database.OpenRecordset("B_MINOR_STRATA")

With prm_record

.Index = "primarykey"

Do Until EOF(1)

Line Input #1, XXX

.AddNew

![COMPOSITE_CODE] = Format(YYI, "0000") + "+" + Format(MMI, "00") + "+" + Left(XXX, 4)
![CODE] = Left(XXX, 4)
![NAME_1] = Mid(XXX, 6, 30)
![NAME_2] = Mid(XXX, 37, 31)
![SORT_SEQ] = Mid(XXX, 69, 6)

.Update

Loop

Close #1

End With

prm_record.Close
prm_database.Close

End Sub
Private Sub CRDB_SITES()

Dim fnm, DBN, DBN2, XXX, xcode, xmnc, xact, xsmp, xfrm, xrec

DBN2 = APPROOT + "\ARTBAS\EXPORT\" + "Y" + Format(YYI, "0000") + "M" + Format(MMI, "00") + "_ALLDATA.MDB"

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(YYI, "0000") + _
      "M" + Format(MMI, "00") + "_SITES.TXT"

If Dir(fnm) = "" Then Exit Sub

Open fnm For Input As #1

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(DBN2)
Set prm_record = prm_database.OpenRecordset("C_SITES")

With prm_record

.Index = "primarykey"

Do Until EOF(1)

Line Input #1, XXX

.AddNew

![COMPOSITE_CODE] = Format(YYI, "0000") + "+" + Format(MMI, "00") + "+" + Left(XXX, 4)
![CODE] = Left(XXX, 4)
![NAME_1] = Mid(XXX, 6, 30)
![NAME_2] = Mid(XXX, 37, 31)
![SORT_SEQ] = Mid(XXX, 69, 6)

.Update

Loop

Close #1

End With

prm_record.Close
prm_database.Close
End Sub
Private Sub CRDB_ASSOCIATIONS()

Dim SI_TABC(1 To 10000), SI_TABN(1 To 10000), SI_MNC(1 To 10000), SI_MND(1 To 10000), SI_MAJC(1 To 10000), SI_MAJN(1 To 10000)
Dim I, J, K, L, m, N, XXX, YYY, NOREC, fnm, DBN2, MN_MAJC(1 To 10000), MN_MAJN(1 To 10000)

Dim SITC(1 To 10000), SITD(1 To 10000), NSTRC(1 To 10000), NSTRD(1 To 10000), JSTRC(1 To 10000), JSTRD(1 To 10000)

For I = 1 To 10000
SITC(I) = 0: SITD(I) = Space(30)
NSTRC(I) = 0: NSTRD(I) = Space(30)
JSTRC(I) = 0: JSTRD(I) = Space(30)
Next I

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(YYI, "0000") + _
      "M" + Format(MMI, "00") + "_SITES.TXT"
If Dir(fnm) = "" Then Exit Sub

Open fnm For Input As #1

Do Until EOF(1)
Line Input #1, XXX
L = Val(Left(XXX, 4)): SITC(L) = L: SITD(L) = Mid(XXX, 6, 30)
If LANG_IND = "INTER" Then SITD(L) = Mid(XXX, 37, 30)

Loop

Close #1

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(YYI, "0000") + _
      "M" + Format(MMI, "00") + "_MINOR.TXT"
If Dir(fnm) = "" Then Exit Sub

Open fnm For Input As #1

Do Until EOF(1)
Line Input #1, XXX
L = Val(Left(XXX, 4)): NSTRC(L) = L: NSTRD(L) = Mid(XXX, 6, 30)
If LANG_IND = "INTER" Then NSTRD(L) = Mid(XXX, 37, 30)

Loop

Close #1

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(YYI, "0000") + _
      "M" + Format(MMI, "00") + "_MAJOR.TXT"
If Dir(fnm) = "" Then Exit Sub

Open fnm For Input As #1

Do Until EOF(1)
Line Input #1, XXX
L = Val(Left(XXX, 4)): JSTRC(L) = L: JSTRD(L) = Mid(XXX, 6, 30)
If LANG_IND = "INTER" Then JSTRD(L) = Mid(XXX, 37, 30)

Loop

Close #1

'================================================================================================================
For I = 1 To 10000
SI_TABC(I) = 0: SI_TABN(I) = Space(30): SI_MNC(I) = 0: SI_MAJC(I) = 0: SI_MND(I) = Space(30): SI_MAJN(I) = Space(30)
MN_MAJC(I) = 0: MN_MAJN(I) = Space(30)
Next I

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(YYI, "0000") + _
      "M" + Format(MMI, "00") + "_ASSOSI.TXT"

If Dir(fnm) = "" Then Exit Sub

Open fnm For Input As #1

Do Until EOF(1)

Line Input #1, XXX
XXX = LTrim(RTrim(XXX))
NOREC = Val(Right(XXX, 4))

For J = 1 To NOREC
Line Input #1, YYY
YYY = LTrim(YYY)

K = Val(Left(YYY, 4))
SI_TABC(K) = K: SI_MNC(K) = Left(XXX, 4)

Next J

Loop

Close #1
'--------------------------------------------------------------------------------------------
fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(YYI, "0000") + _
      "M" + Format(MMI, "00") + "_ASSOMN.TXT"

If Dir(fnm) = "" Then Exit Sub

Open fnm For Input As #1

Do Until EOF(1)

Line Input #1, XXX
XXX = LTrim(RTrim(XXX))
NOREC = Val(Right(XXX, 4))

For J = 1 To NOREC
Line Input #1, YYY
YYY = LTrim(YYY)

L = Val(Left(YYY, 4))
MN_MAJC(L) = Left(XXX, 4)
Next J

Loop

Close #1

For J = 1 To 10000
If SI_TABC(J) = 0 Then GoTo CONT_J

K = SI_MNC(J)
SI_MAJC(J) = MN_MAJC(K)
CONT_J:

Next J
'--------------------------------------------------------------------------------------------
DBN2 = APPROOT + "\ARTBAS\EXPORT\" + "Y" + Format(YYI, "0000") + "M" + Format(MMI, "00") + "_ALLDATA.MDB"

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(DBN2)
Set prm_record = prm_database.OpenRecordset("D_ASSOCIATIONS")

With prm_record

.Index = "primarykey"

For K = 1 To 10000

If SI_TABC(K) = 0 Then GoTo CONT_K

ASSOSIMNC(K) = Val(SI_MNC(K))

m = ASSOSIMNC(K)

.AddNew

![COMPOSITE_CODE] = Format(YYI, "0000") + "+" + Format(MMI, "00")
![COMPOSITE_CODE] = ![COMPOSITE_CODE] + "+" + Format(SI_TABC(K), "0000")
![COMPOSITE_CODE] = ![COMPOSITE_CODE] + "+" + SI_MNC(K)
![COMPOSITE_CODE] = ![COMPOSITE_CODE] + "+" + SI_MAJC(K)

![SITE_CODE] = Format(SI_TABC(K), "0000")
![SITE_NAME] = SITD(K)
![MINOR_CODE] = SI_MNC(K)
![MINOR_NAME] = NSTRD(m)
![Major_Code] = SI_MAJC(K)

m = SI_MAJC(K)

![Major_Name] = JSTRD(m)


.Update

CONT_K:

Next K

Close #1

End With

prm_record.Close
prm_database.Close

End Sub
Private Sub CRDB_ACTIVE_DAYS()

Dim fnm, DBN, DBN2, XXX, xcode, xmnc, xact, xsmp, xfrm, xrec, K, I

Dim MINOR_NAME(1 To 10000), BG_NAME(1 To 10000)

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(YYI, "0000") + _
      "M" + Format(MMI, "00") + "_MINOR.TXT"

If Dir(fnm) = "" Then Exit Sub

For I = 1 To 10000
MINOR_NAME(I) = Space(30)
Next I

Open fnm For Input As #1

Do Until EOF(1)

Line Input #1, XXX
K = Val(Left(XXX, 4))
MINOR_NAME(K) = Mid(XXX, 6, 30)
If LANG_IND = "INTER" Then MINOR_NAME(K) = Mid(XXX, 37, 30)
Loop

Close #1
'============================================================================================

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(YYI, "0000") + _
      "M" + Format(MMI, "00") + "_BG.TXT"

If Dir(fnm) = "" Then Exit Sub

For I = 1 To 10000
BG_NAME(I) = Space(30)
Next I

Open fnm For Input As #1

Do Until EOF(1)

Line Input #1, XXX
K = Val(Left(XXX, 4))
BG_NAME(K) = Mid(XXX, 6, 30)
If LANG_IND = "INTER" Then BG_NAME(K) = Mid(XXX, 37, 30)
Loop

Close #1
'============================================================================================
DBN2 = APPROOT + "\ARTBAS\EXPORT\" + "Y" + Format(YYI, "0000") + "M" + Format(MMI, "00") + "_ALLDATA.MDB"

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(YYI, "0000") + _
      "M" + Format(MMI, "00") + "_ACTIVE.TXT"

If Dir(fnm) = "" Then Exit Sub

Open fnm For Input As #1

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(DBN2)
Set prm_record = prm_database.OpenRecordset("H_ACTIVE_DAYS")

With prm_record

.Index = "primarykey"

Do Until EOF(1)

Line Input #1, XXX

.AddNew

![COMPOSITE_CODE] = Format(YYI, "0000") + "+" + Format(MMI, "00")
![COMPOSITE_CODE] = ![COMPOSITE_CODE] + "+" + Mid(XXX, 2, 4) + "+" + Mid(XXX, 8, 4)


![MINOR_STRATUM_CODE] = Mid(XXX, 2, 4)
![Minor_Stratum_Name] = MINOR_NAME(Val(Mid(XXX, 2, 4)))

![BOAT_GEAR_CODE] = Mid(XXX, 8, 4)
![BOAT_GEAR_NAME] = BG_NAME(Val(Mid(XXX, 8, 4)))

![ACTIVE_DAYS] = Val(Right(XXX, 5))

.Update

Loop

Close #1

End With

prm_record.Close
prm_database.Close

End Sub
Private Sub CRDB_FRAME_SURVEY()

Dim fnm, DBN, DBN2, XXX, xcode, xmnc, xact, xsmp, xfrm, xrec, K, I

Dim MINC(1 To 10000), MINN(1 To 10000)

For I = 1 To 10000
MINC(I) = 0: MINN(I) = Space(30)
Next I

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(YYI, "0000") + _
      "M" + Format(MMI, "00") + "_MINOR.TXT"

If Dir(fnm) = "" Then Exit Sub

Open fnm For Input As #1

Do Until EOF(1)

Line Input #1, XXX
K = Val(Left(XXX, 4))
MINN(K) = Mid(XXX, 6, 30)
If LANG_IND = "INTER" Then MINN(K) = Mid(XXX, 37, 30)
Loop

Close #1

'-----------------------------------------------------------------
Dim BG_NAME(1 To 10000)

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(YYI, "0000") + _
      "M" + Format(MMI, "00") + "_BG.TXT"

If Dir(fnm) = "" Then Exit Sub

For I = 1 To 10000
BG_NAME(I) = Space(30)
Next I

Open fnm For Input As #1

Do Until EOF(1)

Line Input #1, XXX
K = Val(Left(XXX, 4))
BG_NAME(K) = Mid(XXX, 6, 30)
If LANG_IND = "INTER" Then BG_NAME(K) = Mid(XXX, 37, 30)
Loop

Close #1
'============================================================================================
Dim LS_NAME(1 To 10000)

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(YYI, "0000") + _
      "M" + Format(MMI, "00") + "_SITES.TXT"

If Dir(fnm) = "" Then Exit Sub

For I = 1 To 10000
LS_NAME(I) = Space(30)
Next I

Open fnm For Input As #1

Do Until EOF(1)

Line Input #1, XXX
K = Val(Left(XXX, 4))
LS_NAME(K) = Mid(XXX, 6, 30)
If LANG_IND = "INTER" Then LS_NAME(K) = Mid(XXX, 37, 30)
Loop

Close #1
'============================================================================================
DBN2 = APPROOT + "\ARTBAS\EXPORT\" + "Y" + Format(YYI, "0000") + "M" + Format(MMI, "00") + "_ALLDATA.MDB"

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(YYI, "0000") + _
      "M" + Format(MMI, "00") + "_FRAME.TXT"

If Dir(fnm) = "" Then Exit Sub

Open fnm For Input As #1

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(DBN2)
Set prm_record = prm_database.OpenRecordset("G_FRAME_SURVEY")

With prm_record

.Index = "primarykey"

Do Until EOF(1)

Line Input #1, XXX

.AddNew

![COMPOSITE_CODE] = Format(YYI, "0000") + "+" + Format(MMI, "00") + "+"

K = Val(Mid(XXX, 2, 4))

![COMPOSITE_CODE] = ![COMPOSITE_CODE] + Format(ASSOSIMNC(K), "0000") + "+"
![COMPOSITE_CODE] = ![COMPOSITE_CODE] + Mid(XXX, 2, 4) + "+"
![COMPOSITE_CODE] = ![COMPOSITE_CODE] + Mid(XXX, 8, 4)

![MINOR_STRATUM_CODE] = Format(ASSOSIMNC(K), "0000")
![Minor_Stratum_Name] = MINN(ASSOSIMNC(K))

![BOAT_GEAR_CODE] = Mid(XXX, 8, 4)
![BOAT_GEAR_NAME] = BG_NAME(Val(Mid(XXX, 8, 4)))

![SITE_CODE] = Mid(XXX, 2, 4)
![SITE_NAME] = LS_NAME(Val(Mid(XXX, 2, 4)))

![NO_UNITS] = Val(Mid(XXX + Space(20), 26, 15))

.Update

Loop

Close #1

End With

prm_record.Close
prm_database.Close

End Sub
Private Sub CRDB_EFFORT()

Call LOAD_ALL_TABLES

If LOAD_TABLES_FLAG <> "OK" Then Exit Sub

Dim fnm, DBN, DBN2, XXX, xcode, xmnc, xact, xsmp, xfrm, xrec, I, J, K, L, m, N, NCODE

DBN2 = APPROOT + "\ARTBAS\EXPORT\" + "Y" + Format(YYI, "0000") + "M" + Format(MMI, "00") + "_ALLDATA.MDB"

fnm = APPROOT + "\ARTBAS\EFFORT\Y" + Format(YYI, "0000") + _
      "M" + Format(MMI, "00") + "_ESAMPLES.TXT"

If Dir(fnm) = "" Then Exit Sub

Open fnm For Input As #1

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(DBN2)
Set prm_record = prm_database.OpenRecordset("I_BOAT_ACTIVITIES")

With prm_record

.Index = "primarykey"

Do Until EOF(1)

Line Input #1, XXX

.AddNew

![COMPOSITE_CODE] = Format(YYI, "0000") + "+" + Format(MMI, "00") + "+"
![COMPOSITE_CODE] = ![COMPOSITE_CODE] + Mid(XXX, 17, 4) + "+"
![COMPOSITE_CODE] = ![COMPOSITE_CODE] + Mid(XXX, 2, 4) + "+"
![COMPOSITE_CODE] = ![COMPOSITE_CODE] + Mid(XXX, 8, 4) + "+"
![COMPOSITE_CODE] = ![COMPOSITE_CODE] + Mid(XXX, 14, 2)

![MINOR_CODE] = Mid(XXX, 17, 4): ![MINOR_NAME] = MND(Val(Mid(XXX, 17, 4)))
![SITE_CODE] = Mid(XXX, 2, 4): ![SITE_NAME] = SID(Val(Mid(XXX, 2, 4)))
![GEAR_CODE] = Mid(XXX, 8, 4): ![GEAR_NAME] = BGD(Val(Mid(XXX, 8, 4)))

![Day] = Mid(XXX, 14, 2)

![ACTIVE_BOATS] = CDbl(Mid(XXX, 22, 10))
![SAMPLED_BOATS] = CDbl(Mid(XXX, 33, 10))
![FRAME_BOATS] = CDbl(Mid(XXX, 44, 10))
![RECORDER] = Mid(XXX + Space(15), 55, 15)

If ![ACTIVE_BOATS] <> 0 And ![SAMPLED_BOATS] = 0 Then
![SAMPLED_BOATS] = ![FRAME_BOATS]
End If

.Update

Loop

Close #1

End With

prm_record.Close
prm_database.Close

End Sub
Private Sub CRDB_ESTIM_REMARKS()

Call LOAD_ALL_TABLES

If LOAD_TABLES_FLAG <> "OK" Then Exit Sub

Dim fnm, DBN, DBN2, XXX, xcode, xmnc, xact, xsmp, xfrm, xrec, I, J, K, L, m, N, NCODE

DBN2 = APPROOT + "\ARTBAS\EXPORT\" + "Y" + Format(YYI, "0000") + "M" + Format(MMI, "00") + "_ALLDATA.MDB"

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(DBN2)
Set prm_record = prm_database.OpenRecordset("N_ESTIMATION_REMARKS")

For I = 1 To 10000

If MNC(I) = 0 Then GoTo NEXT_I

fnm = APPROOT + "\ARTBAS\RESULTS\Y" + Format(YYI, "0000") + _
      "M" + Format(MMI, "00") + "_MN" + Format(I, "0000") + "_LOG.TXT"

If Dir(fnm) = "" Then GoTo NEXT_I

Open fnm For Input As #1: m = 0

With prm_record

.Index = "primarykey"

Do Until EOF(1)

Line Input #1, XXX: m = m + 1

.AddNew

![COMPOSITE_CODE] = Format(YYI, "0000") + "+" + Format(MMI, "00") + _
                    "+" + Format(I, "0000") + "+" + Format(m, "0000")

![ESTIMATION_REMARK] = LTrim(RTrim(XXX))


.Update

Loop

Close #1

End With

NEXT_I:

Next I

prm_record.Close
prm_database.Close

End Sub
Private Sub CRDB_TRIP_TOTALS()

Call LOAD_ALL_TABLES

If LOAD_TABLES_FLAG <> "OK" Then Exit Sub

Dim fnm, DBN, DBN2, XXX, xcode, xmnc, xact, xsmp, xfrm, xrec, I, J, K, L, m, N, NCODE

DBN2 = APPROOT + "\ARTBAS\EXPORT\" + "Y" + Format(YYI, "0000") + "M" + Format(MMI, "00") + "_ALLDATA.MDB"

fnm = APPROOT + "\ARTBAS\LANDINGS\Y" + Format(YYI, "0000") + _
      "M" + Format(MMI, "00") + "_LSAMPLES.TXT"

If Dir(fnm) = "" Then Exit Sub

Open fnm For Input As #1

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(DBN2)
Set prm_record = prm_database.OpenRecordset("J_TRIP_TOTALS")

With prm_record

.Index = "primarykey"

Do Until EOF(1)

Line Input #1, XXX

.AddNew

![COMPOSITE_CODE] = Format(YYI, "0000") + "+" + Format(MMI, "00") + "+"
![COMPOSITE_CODE] = ![COMPOSITE_CODE] + Mid(XXX, 85, 4) + "+"
![COMPOSITE_CODE] = ![COMPOSITE_CODE] + Mid(XXX, 97, 4) + "+"
![COMPOSITE_CODE] = ![COMPOSITE_CODE] + Mid(XXX, 91, 4) + "+"


![COMPOSITE_CODE] = ![COMPOSITE_CODE] + Mid(XXX, 1, 6)

![MINOR_CODE] = Mid(XXX, 85, 4): ![MINOR_NAME] = MND(Val(Mid(XXX, 85, 4)))
![SITE_CODE] = Mid(XXX, 91, 4): ![SITE_NAME] = SID(Val(Mid(XXX, 91, 4)))
![GEAR_CODE] = Mid(XXX, 97, 4): ![GEAR_NAME] = BGD(Val(Mid(XXX, 97, 4)))

![Day] = Mid(XXX, 8, 2)
![SAMPLE_NO] = Mid(XXX, 1, 6)
![RECORDER] = Mid(XXX, 69, 15)
![NO_UNITS] = CDbl(Mid(XXX, 11, 8))
![Duration] = CDbl(Mid(XXX, 20, 8))
![RAISING_FACTOR] = CDbl(Mid(XXX, 49, 19))
![SAMPLE_TOT] = CDbl(Mid(XXX, 29, 19))

.Update

Loop

Close #1

End With

prm_record.Close
prm_database.Close

End Sub
Private Sub CRDB_BY_SPECIES()

Call LOAD_ALL_TABLES

If LOAD_TABLES_FLAG <> "OK" Then Exit Sub

Dim fnm, DBN, DBN2, XXX, xcode, xmnc, xact, xsmp, xfrm, xrec, I, J, K, L, m, N, NCODE

DBN2 = APPROOT + "\ARTBAS\EXPORT\" + "Y" + Format(YYI, "0000") + "M" + Format(MMI, "00") + "_ALLDATA.MDB"

fnm = APPROOT + "\ARTBAS\LANDINGS\Y" + Format(YYI, "0000") + _
      "M" + Format(MMI, "00") + "_LSPECIES.TXT"

If Dir(fnm) = "" Then Exit Sub

Open fnm For Input As #1

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(DBN2)
Set prm_record = prm_database.OpenRecordset("K_BY_SPECIES")

With prm_record

.Index = "primarykey"

Do Until EOF(1)

Line Input #1, XXX

.AddNew

![COMPOSITE_CODE] = Format(YYI, "0000") + "+" + Format(MMI, "00") + "+"
![COMPOSITE_CODE] = ![COMPOSITE_CODE] + Mid(XXX, 95, 4) + "+"
![COMPOSITE_CODE] = ![COMPOSITE_CODE] + Mid(XXX, 107, 4) + "+"
![COMPOSITE_CODE] = ![COMPOSITE_CODE] + Mid(XXX, 101, 4) + "+"

![COMPOSITE_CODE] = ![COMPOSITE_CODE] + Mid(XXX, 2, 6) + "+"
![COMPOSITE_CODE] = ![COMPOSITE_CODE] + Mid(XXX, 10, 4)

![MINOR_CODE] = Mid(XXX, 95, 4): ![MINOR_NAME] = MND(Val(Mid(XXX, 95, 4)))
![SITE_CODE] = Mid(XXX, 101, 4): ![SITE_NAME] = SID(Val(Mid(XXX, 101, 4)))
![GEAR_CODE] = Mid(XXX, 107, 4): ![GEAR_NAME] = BGD(Val(Mid(XXX, 107, 4)))
![SPECIES_CODE] = Mid(XXX, 10, 4): ![SPECIES_NAME] = SPED(Val(Mid(XXX, 10, 4)))

![SAMPLE_NO] = Mid(XXX, 2, 6)

![catch] = CDbl(Mid(XXX, 15, 19))
![NO_FISH] = CDbl(Mid(XXX, 35, 19))
![price] = CDbl(Mid(XXX, 55, 19))
![Value] = CDbl(Mid(XXX, 75, 19))

.Update

Loop

Close #1

End With

prm_record.Close
prm_database.Close

End Sub
Private Sub CRDB_RESULTS()

Call LOAD_ALL_TABLES

If LOAD_TABLES_FLAG <> "OK" Then Exit Sub

Dim fnm, DBN, DBN2, XXX, xcode, xmnc, xact, xsmp, xfrm, xrec, I, J, K, L, m, N, NCODE, GUIDE_CODE

Dim estdes, popn, smpn, BAC_ACCUR, FRNO, actno, cal, eact, _
          esmp, esites, edays, BAC, bac_cvs, bac_cvsp, _
          bac_cvt, bac_cvtp, bac_cv, bac_low, bac_upper, eff, _
          eff_low, eff_upper, nland, CPUE_ACCUR, LPOP, ltot, _
          lsmpv, lsmpf, leff, cpue, lsites, ldays, _
          cpue_cvs, cpue_cvsp, cpue_cvt, cpue_cvtp, cpue_cv, cpue_low, _
          cpue_upper, catch, catch_low, catch_upper, catch_cv, Value, _
          price, fish, kgfish

DBN2 = APPROOT + "\ARTBAS\EXPORT\" + "Y" + Format(YYI, "0000") + "M" + Format(MMI, "00") + "_ALLDATA.MDB"

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(DBN2)
Set prm_record = prm_database.OpenRecordset("L_RESULTS")

For I = 1 To 10000

If MNC(I) = 0 Then GoTo NEXT_I

fnm = APPROOT + "\ARTBAS\RESULTS\Y" + Format(YYI, "0000") + _
      "M" + Format(MMI, "00") + "_MN" + Format(I, "0000") + "_ESTIM.TXT"

If Dir(fnm) = "" Then GoTo NEXT_I

Open fnm For Input As #1: m = 0

With prm_record

.Index = "primarykey"

Do Until EOF(1)

Input #1, GUIDE_CODE
GUIDE_CODE = LTrim(RTrim(GUIDE_CODE))

If Mid(GUIDE_CODE, 8, 4) <> "0000" And Mid(GUIDE_CODE, 14, 4) = "0000" Then GoTo READ_TRIP
If Mid(GUIDE_CODE, 8, 4) <> "0000" And Mid(GUIDE_CODE, 14, 4) <> "0000" Then GoTo READ_SPECIES
If Mid(GUIDE_CODE, 8, 4) = "0000" And Mid(GUIDE_CODE, 14, 4) <> "0000" Then GoTo READ_SPECIES
If Mid(GUIDE_CODE, 8, 4) = "0000" And Mid(GUIDE_CODE, 14, 4) = "0000" Then GoTo READ_SPECIES

READ_TRIP:

Input #1, estdes, popn, smpn, BAC_ACCUR, FRNO, actno, cal, eact, _
          esmp, esites, edays, BAC, bac_cvs, bac_cvsp, _
          bac_cvt, bac_cvtp, bac_cv, bac_low, bac_upper, eff, _
          eff_low, eff_upper, nland, CPUE_ACCUR, LPOP, ltot, _
          lsmpv, lsmpf, leff, cpue, lsites, ldays, _
          cpue_cvs, cpue_cvsp, cpue_cvt, cpue_cvtp, cpue_cv, cpue_low, _
          cpue_upper, catch, catch_low, catch_upper, catch_cv, Value, _
          price, fish, kgfish

If Mid(GUIDE_CODE, 8, 4) = "0000" Or catch + eff = 0 Then GoTo CONT_READ

.AddNew

![COMPOSITE_CODE] = Format(YYI, "0000") + "+" + Format(MMI, "00") + "+"
![COMPOSITE_CODE] = ![COMPOSITE_CODE] + Format(I, "0000") + "+"
![COMPOSITE_CODE] = ![COMPOSITE_CODE] + Mid(GUIDE_CODE, 8, 4) + "+"
![COMPOSITE_CODE] = ![COMPOSITE_CODE] + Mid(GUIDE_CODE, 14, 4)

![MINOR_CODE] = Format(I, "0000"): ![MINOR_NAME] = MND(I)
![GEAR_CODE] = Mid(GUIDE_CODE, 8, 4): ![GEAR_NAME] = BGD(Val(Mid(GUIDE_CODE, 8, 4)))
![SPECIES_CODE] = "0000": ![SPECIES_NAME] = msgtab(201)


![catch] = catch
![effort] = eff
![cpue] = cpue
'If kgfish <> 0 Then ![aver_weight] = kgfish
'If fish <> 0 Then ![NO_FISH] = fish
'If price <> 0 Then ![price] = price
'If Value <> 0 Then ![Value] = Value
![BAC] = BAC
![ACCUR_BAC] = 100 * BAC_ACCUR
![ACCUR_CPUE] = 100 * CPUE_ACCUR

.Update

GoTo CONT_READ

READ_SPECIES:

Input #1, estdes, eff, cpue, catch, Value, price, kgfish, _
           fish, FRNO

If Mid(GUIDE_CODE, 8, 4) = "0000" Or catch + eff = 0 Then GoTo CONT_READ

.AddNew

![COMPOSITE_CODE] = Format(YYI, "0000") + "+" + Format(MMI, "00") + "+"
![COMPOSITE_CODE] = ![COMPOSITE_CODE] + Format(I, "0000") + "+"
![COMPOSITE_CODE] = ![COMPOSITE_CODE] + Mid(GUIDE_CODE, 8, 4) + "+"
![COMPOSITE_CODE] = ![COMPOSITE_CODE] + Mid(GUIDE_CODE, 14, 4)

![MINOR_CODE] = Format(I, "0000"): ![MINOR_NAME] = MND(I)
![GEAR_CODE] = Mid(GUIDE_CODE, 8, 4): ![GEAR_NAME] = BGD(Val(Mid(GUIDE_CODE, 8, 4)))
![SPECIES_CODE] = Mid(GUIDE_CODE, 14, 4): ![SPECIES_NAME] = SPED(Val(Mid(GUIDE_CODE, 14, 4)))


![catch] = catch
![effort] = eff
![cpue] = cpue
If kgfish <> 0 Then ![aver_weight] = kgfish
If fish <> 0 Then ![NO_FISH] = fish
If price <> 0 Then ![price] = price
If Value <> 0 Then ![Value] = Value

.Update

CONT_READ:

Loop

Close #1

End With

NEXT_I:

Next I

prm_record.Close
prm_database.Close

End Sub
Private Sub LOAD_ALL_TABLES()

Dim I, fnm, XXX

LOAD_TABLES_FLAG = "OK"

For I = 1 To 10000
MNC(I) = 0: MND(I) = Space(30)
SIC(I) = 0: SID(I) = Space(30)
BGC(I) = 0: BGD(I) = Space(30)
SPEC(I) = 0: SPED(I) = Space(30)
Next I

'--------------- Load species -----------------------------------------------
fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(YYI, "0000") + _
      "M" + Format(MMI, "00") + "_SPECIES.TXT"

If Dir(fnm) = "" Then
LOAD_TABLES_FLAG = "NO"
Exit Sub
End If

Open fnm For Input As #1

Do Until EOF(1)

Line Input #1, XXX
SPEC(Val(Left(XXX, 4))) = Val(Left(XXX, 4)): SPED(Val(Left(XXX, 4))) = Mid(XXX, 6, 30)
If LANG_IND = "INTER" Then SPED(Val(Left(XXX, 4))) = Mid(XXX, 37, 30)
Loop

Close #1
'--------------- Load Minor Strata ------------------------------------------

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(YYI, "0000") + _
      "M" + Format(MMI, "00") + "_MINOR.TXT"

If Dir(fnm) = "" Then
LOAD_TABLES_FLAG = "NO"
Exit Sub
End If

Open fnm For Input As #1

Do Until EOF(1)

Line Input #1, XXX
MNC(Val(Left(XXX, 4))) = Val(Left(XXX, 4)): MND(Val(Left(XXX, 4))) = Mid(XXX, 6, 30)
If LANG_IND = "INTER" Then MND(Val(Left(XXX, 4))) = Mid(XXX, 37, 30)
Loop

Close #1

'-------------- Load Sites ---------------------------------------------------

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(YYI, "0000") + _
      "M" + Format(MMI, "00") + "_SITES.TXT"

If Dir(fnm) = "" Then
LOAD_TABLES_FLAG = "NO"
Exit Sub
End If

Open fnm For Input As #1

Do Until EOF(1)

Line Input #1, XXX
SIC(Val(Left(XXX, 4))) = Val(Left(XXX, 4)): SID(Val(Left(XXX, 4))) = Mid(XXX, 6, 30)
If LANG_IND = "INTER" Then SID(Val(Left(XXX, 4))) = Mid(XXX, 37, 30)
Loop

Close #1

'-------------- Load Boat/Gear  ---------------------------------------------------

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(YYI, "0000") + _
      "M" + Format(MMI, "00") + "_BG.TXT"

If Dir(fnm) = "" Then
LOAD_TABLES_FLAG = "NO"
Exit Sub
End If

Open fnm For Input As #1

Do Until EOF(1)

Line Input #1, XXX
BGC(Val(Left(XXX, 4))) = Val(Left(XXX, 4)): BGD(Val(Left(XXX, 4))) = Mid(XXX, 6, 30)
If LANG_IND = "INTER" Then BGD(Val(Left(XXX, 4))) = Mid(XXX, 37, 30)
Loop

Close #1
'------------------------------------------------------------------------------------------------

End Sub
Private Sub APPEND_DATABASES()

Dim II

Label2.Visible = True
pgbFILES.Visible = True

Dim resp

resp = MsgBox(msgtab(282), vbYesNo, " ")

If resp = vbNo Then End

Label2.Visible = True

pgbFILES.Min = 0
pgbFILES.Max = NYT
pgbFILES.Value = 0

Dim APPI

APPI = 0

Dim YYX, MMX, EMPTYDBN, WDB, DBN, DBN2

DBN2 = APPROOT + "\ARTBAS\EXPORT\" + "CUMMULATIVE_ALLDATA.MDB"
DBN = APPROOT + "\ARTBAS\STRUS\ALLDATA.MDB"

FileCopy DBN, DBN2

'====================== LOOP =====================

For II = 1 To NYT

MMI = Val(Mid(YMTAB(II), 8, 2))
YYI = Val(Left(YMTAB(II), 4))

'============================================================

pgbFILES.Value = II

Label2.Caption = Format(MMI, "00") + "/" + Format(YYI, "0000")
Label2.Refresh

Call CUM_MAJOR_STRATA
Call CUM_SPECIES
Call CUM_BOATS_GEARS
Call CUM_MINOR_STRATA
Call CUM_SITES
Call CUM_ASSOCIATIONS
Call CUM_ACTIVE_DAYS
Call CUM_FRAME_SURVEY
Call CUM_EFFORT
Call CUM_ESTIM_REMARKS
Call CUM_TRIP_TOTALS
Call CUM_BY_SPECIES
Call CUM_RESULTS

Next II

END_SUB:

resp = MsgBox(msgtab(278) + " " + msgtab(284), vbOKOnly, " ")

End

End Sub
Private Sub CUM_MAJOR_STRATA()

Dim fnm, DBN, DBN2, XXX, xcode, xmnc, xact, xsmp, xfrm, xrec

DBN2 = APPROOT + "\ARTBAS\EXPORT\" + "CUMMULATIVE_ALLDATA.MDB"

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(YYI, "0000") + _
      "M" + Format(MMI, "00") + "_MAJOR.TXT"

If Dir(fnm) = "" Then Exit Sub

Label2.Caption = fnm
Label2.Refresh

Open fnm For Input As #1

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(DBN2)
Set prm_record = prm_database.OpenRecordset("A_MAJOR_STRATA")

With prm_record

.Index = "primarykey"

Do Until EOF(1)

Line Input #1, XXX

.AddNew

![COMPOSITE_CODE] = Format(YYI, "0000") + "+" + Format(MMI, "00") + "+" + Left(XXX, 4)
![CODE] = Left(XXX, 4)
![NAME_1] = Mid(XXX, 6, 30)
![NAME_2] = Mid(XXX, 37, 31)
![SORT_SEQ] = Mid(XXX, 69, 6)

.Update

Loop

Close #1

End With

prm_record.Close
prm_database.Close

End Sub
Private Sub CUM_SPECIES()

Dim fnm, DBN, DBN2, XXX, xcode, xmnc, xact, xsmp, xfrm, xrec

DBN2 = APPROOT + "\ARTBAS\EXPORT\" + "CUMMULATIVE_ALLDATA.MDB"

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(YYI, "0000") + _
      "M" + Format(MMI, "00") + "_SPECIES.TXT"

If Dir(fnm) = "" Then Exit Sub

Label2.Caption = fnm
Label2.Refresh

Open fnm For Input As #1

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(DBN2)
Set prm_record = prm_database.OpenRecordset("F_SPECIES")

With prm_record

.Index = "primarykey"

Do Until EOF(1)

Line Input #1, XXX

.AddNew

![COMPOSITE_CODE] = Format(YYI, "0000") + "+" + Format(MMI, "00") + "+" + Left(XXX, 4)
![CODE] = Left(XXX, 4)
![NAME_1] = Mid(XXX, 6, 30)
![NAME_2] = Mid(XXX, 37, 31)
![SORT_SEQ] = Mid(XXX, 69, 6)

.Update

Loop

Close #1

End With

prm_record.Close
prm_database.Close

End Sub
Private Sub CUM_BOATS_GEARS()

Dim fnm, DBN, DBN2, XXX, xcode, xmnc, xact, xsmp, xfrm, xrec

DBN2 = APPROOT + "\ARTBAS\EXPORT\" + "CUMMULATIVE_ALLDATA.MDB"

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(YYI, "0000") + _
      "M" + Format(MMI, "00") + "_BG.TXT"

If Dir(fnm) = "" Then Exit Sub

Label2.Caption = fnm
Label2.Refresh


Open fnm For Input As #1

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(DBN2)
Set prm_record = prm_database.OpenRecordset("E_BOATS_GEARS")

With prm_record

.Index = "primarykey"

Do Until EOF(1)

Line Input #1, XXX

.AddNew

![COMPOSITE_CODE] = Format(YYI, "0000") + "+" + Format(MMI, "00") + "+" + Left(XXX, 4)
![CODE] = Left(XXX, 4)
![NAME_1] = Mid(XXX, 6, 30)
![NAME_2] = Mid(XXX, 37, 31)
![SORT_SEQ] = Mid(XXX, 69, 6)

.Update

Loop

Close #1

End With

prm_record.Close
prm_database.Close

End Sub
Private Sub CUM_MINOR_STRATA()

Dim fnm, DBN, DBN2, XXX, xcode, xmnc, xact, xsmp, xfrm, xrec

DBN2 = APPROOT + "\ARTBAS\EXPORT\" + "CUMMULATIVE_ALLDATA.MDB"

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(YYI, "0000") + _
      "M" + Format(MMI, "00") + "_MINOR.TXT"

If Dir(fnm) = "" Then Exit Sub

Label2.Caption = fnm
Label2.Refresh


Open fnm For Input As #1

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(DBN2)
Set prm_record = prm_database.OpenRecordset("B_MINOR_STRATA")

With prm_record

.Index = "primarykey"

Do Until EOF(1)

Line Input #1, XXX

.AddNew

![COMPOSITE_CODE] = Format(YYI, "0000") + "+" + Format(MMI, "00") + "+" + Left(XXX, 4)
![CODE] = Left(XXX, 4)
![NAME_1] = Mid(XXX, 6, 30)
![NAME_2] = Mid(XXX, 37, 31)
![SORT_SEQ] = Mid(XXX, 69, 6)

.Update

Loop

Close #1

End With

prm_record.Close
prm_database.Close

End Sub
Private Sub CUM_SITES()

Dim fnm, DBN, DBN2, XXX, xcode, xmnc, xact, xsmp, xfrm, xrec

DBN2 = APPROOT + "\ARTBAS\EXPORT\" + "CUMMULATIVE_ALLDATA.MDB"

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(YYI, "0000") + _
      "M" + Format(MMI, "00") + "_SITES.TXT"

If Dir(fnm) = "" Then Exit Sub

Label2.Caption = fnm
Label2.Refresh


Open fnm For Input As #1

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(DBN2)
Set prm_record = prm_database.OpenRecordset("C_SITES")

With prm_record

.Index = "primarykey"

Do Until EOF(1)

Line Input #1, XXX

.AddNew

![COMPOSITE_CODE] = Format(YYI, "0000") + "+" + Format(MMI, "00") + "+" + Left(XXX, 4)
![CODE] = Left(XXX, 4)
![NAME_1] = Mid(XXX, 6, 30)
![NAME_2] = Mid(XXX, 37, 31)
![SORT_SEQ] = Mid(XXX, 69, 6)

.Update

Loop

Close #1

End With

prm_record.Close
prm_database.Close
End Sub
Private Sub CUM_ASSOCIATIONS()

Dim SI_TABC(1 To 10000), SI_TABN(1 To 10000), SI_MNC(1 To 10000), SI_MND(1 To 10000), SI_MAJC(1 To 10000), SI_MAJN(1 To 10000)
Dim I, J, K, L, m, N, XXX, YYY, NOREC, fnm, DBN2, MN_MAJC(1 To 10000), MN_MAJN(1 To 10000)

Dim SITC(1 To 10000), SITD(1 To 10000), NSTRC(1 To 10000), NSTRD(1 To 10000), JSTRC(1 To 10000), JSTRD(1 To 10000)

For I = 1 To 10000
SITC(I) = 0: SITD(I) = Space(30)
NSTRC(I) = 0: NSTRD(I) = Space(30)
JSTRC(I) = 0: JSTRD(I) = Space(30)
Next I

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(YYI, "0000") + _
      "M" + Format(MMI, "00") + "_SITES.TXT"
If Dir(fnm) = "" Then Exit Sub

Label2.Caption = fnm
Label2.Refresh


Open fnm For Input As #1

Do Until EOF(1)
Line Input #1, XXX
L = Val(Left(XXX, 4)): SITC(L) = L: SITD(L) = Mid(XXX, 6, 30)
If LANG_IND = "INTER" Then SITD(L) = Mid(XXX, 37, 30)

Loop

Close #1

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(YYI, "0000") + _
      "M" + Format(MMI, "00") + "_MINOR.TXT"
If Dir(fnm) = "" Then Exit Sub

Open fnm For Input As #1

Do Until EOF(1)
Line Input #1, XXX
L = Val(Left(XXX, 4)): NSTRC(L) = L: NSTRD(L) = Mid(XXX, 6, 30)
If LANG_IND = "INTER" Then NSTRD(L) = Mid(XXX, 37, 30)

Loop

Close #1

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(YYI, "0000") + _
      "M" + Format(MMI, "00") + "_MAJOR.TXT"
If Dir(fnm) = "" Then Exit Sub

Open fnm For Input As #1

Do Until EOF(1)
Line Input #1, XXX
L = Val(Left(XXX, 4)): JSTRC(L) = L: JSTRD(L) = Mid(XXX, 6, 30)
If LANG_IND = "INTER" Then JSTRD(L) = Mid(XXX, 37, 30)

Loop

Close #1

'================================================================================================================
For I = 1 To 10000
SI_TABC(I) = 0: SI_TABN(I) = Space(30): SI_MNC(I) = 0: SI_MAJC(I) = 0: SI_MND(I) = Space(30): SI_MAJN(I) = Space(30)
MN_MAJC(I) = 0: MN_MAJN(I) = Space(30)
Next I

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(YYI, "0000") + _
      "M" + Format(MMI, "00") + "_ASSOSI.TXT"

If Dir(fnm) = "" Then Exit Sub

Open fnm For Input As #1

Do Until EOF(1)

Line Input #1, XXX
XXX = LTrim(RTrim(XXX))
NOREC = Val(Right(XXX, 4))

For J = 1 To NOREC
Line Input #1, YYY
YYY = LTrim(YYY)

K = Val(Left(YYY, 4))
SI_TABC(K) = K: SI_MNC(K) = Left(XXX, 4)

Next J

Loop

Close #1
'--------------------------------------------------------------------------------------------
fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(YYI, "0000") + _
      "M" + Format(MMI, "00") + "_ASSOMN.TXT"

If Dir(fnm) = "" Then Exit Sub

Open fnm For Input As #1

Do Until EOF(1)

Line Input #1, XXX
XXX = LTrim(RTrim(XXX))
NOREC = Val(Right(XXX, 4))

For J = 1 To NOREC
Line Input #1, YYY
YYY = LTrim(YYY)

L = Val(Left(YYY, 4))
MN_MAJC(L) = Left(XXX, 4)
Next J

Loop

Close #1

For J = 1 To 10000
If SI_TABC(J) = 0 Then GoTo CONT_J

K = SI_MNC(J)
SI_MAJC(J) = MN_MAJC(K)
CONT_J:

Next J
'--------------------------------------------------------------------------------------------
DBN2 = APPROOT + "\ARTBAS\EXPORT\" + "CUMMULATIVE_ALLDATA.MDB"

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(DBN2)
Set prm_record = prm_database.OpenRecordset("D_ASSOCIATIONS")

With prm_record

.Index = "primarykey"

For K = 1 To 10000

If SI_TABC(K) = 0 Then GoTo CONT_K

ASSOSIMNC(K) = Val(SI_MNC(K))

m = ASSOSIMNC(K)

.AddNew

![COMPOSITE_CODE] = Format(YYI, "0000") + "+" + Format(MMI, "00")
![COMPOSITE_CODE] = ![COMPOSITE_CODE] + "+" + Format(SI_TABC(K), "0000")
![COMPOSITE_CODE] = ![COMPOSITE_CODE] + "+" + SI_MNC(K)
![COMPOSITE_CODE] = ![COMPOSITE_CODE] + "+" + SI_MAJC(K)

![SITE_CODE] = Format(SI_TABC(K), "0000")
![SITE_NAME] = SITD(K)
![MINOR_CODE] = SI_MNC(K)
![MINOR_NAME] = NSTRD(m)
![Major_Code] = SI_MAJC(K)

m = SI_MAJC(K)

![Major_Name] = JSTRD(m)


.Update

CONT_K:

Next K

Close #1

End With

prm_record.Close
prm_database.Close

End Sub
Private Sub CUM_ACTIVE_DAYS()

Dim fnm, DBN, DBN2, XXX, xcode, xmnc, xact, xsmp, xfrm, xrec, K, I

Dim MINOR_NAME(1 To 10000), BG_NAME(1 To 10000)

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(YYI, "0000") + _
      "M" + Format(MMI, "00") + "_MINOR.TXT"

If Dir(fnm) = "" Then Exit Sub

Label2.Caption = fnm
Label2.Refresh


Label2.Caption = fnm

For I = 1 To 10000
MINOR_NAME(I) = Space(30)
Next I

Open fnm For Input As #1

Do Until EOF(1)

Line Input #1, XXX
K = Val(Left(XXX, 4))
MINOR_NAME(K) = Mid(XXX, 6, 30)
If LANG_IND = "INTER" Then MINOR_NAME(K) = Mid(XXX, 37, 30)
Loop

Close #1
'============================================================================================

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(YYI, "0000") + _
      "M" + Format(MMI, "00") + "_BG.TXT"

If Dir(fnm) = "" Then Exit Sub

Label2.Caption = fnm
Label2.Refresh

For I = 1 To 10000
BG_NAME(I) = Space(30)
Next I

Open fnm For Input As #1

Do Until EOF(1)

Line Input #1, XXX
K = Val(Left(XXX, 4))
BG_NAME(K) = Mid(XXX, 6, 30)
If LANG_IND = "INTER" Then BG_NAME(K) = Mid(XXX, 37, 30)
Loop

Close #1
'============================================================================================
DBN2 = APPROOT + "\ARTBAS\EXPORT\" + "CUMMULATIVE_ALLDATA.MDB"

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(YYI, "0000") + _
      "M" + Format(MMI, "00") + "_ACTIVE.TXT"

If Dir(fnm) = "" Then Exit Sub

Open fnm For Input As #1

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(DBN2)
Set prm_record = prm_database.OpenRecordset("H_ACTIVE_DAYS")

With prm_record

.Index = "primarykey"

Do Until EOF(1)

Line Input #1, XXX

.AddNew

![COMPOSITE_CODE] = Format(YYI, "0000") + "+" + Format(MMI, "00")
![COMPOSITE_CODE] = ![COMPOSITE_CODE] + "+" + Mid(XXX, 2, 4) + "+" + Mid(XXX, 8, 4)


![MINOR_STRATUM_CODE] = Mid(XXX, 2, 4)
![Minor_Stratum_Name] = MINOR_NAME(Val(Mid(XXX, 2, 4)))

![BOAT_GEAR_CODE] = Mid(XXX, 8, 4)
![BOAT_GEAR_NAME] = BG_NAME(Val(Mid(XXX, 8, 4)))

![ACTIVE_DAYS] = Val(Right(XXX, 5))

.Update

Loop

Close #1

End With

prm_record.Close
prm_database.Close

End Sub
Private Sub CUM_FRAME_SURVEY()

Dim fnm, DBN, DBN2, XXX, xcode, xmnc, xact, xsmp, xfrm, xrec, K, I

Dim MINC(1 To 10000), MINN(1 To 10000)

For I = 1 To 10000
MINC(I) = 0: MINN(I) = Space(30)
Next I

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(YYI, "0000") + _
      "M" + Format(MMI, "00") + "_MINOR.TXT"

If Dir(fnm) = "" Then Exit Sub

Label2.Caption = fnm
Label2.Refresh


Open fnm For Input As #1

Do Until EOF(1)

Line Input #1, XXX
K = Val(Left(XXX, 4))
MINN(K) = Mid(XXX, 6, 30)
If LANG_IND = "INTER" Then MINN(K) = Mid(XXX, 37, 30)
Loop

Close #1

'-----------------------------------------------------------------
Dim BG_NAME(1 To 10000)

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(YYI, "0000") + _
      "M" + Format(MMI, "00") + "_BG.TXT"

If Dir(fnm) = "" Then Exit Sub

For I = 1 To 10000
BG_NAME(I) = Space(30)
Next I

Open fnm For Input As #1

Do Until EOF(1)

Line Input #1, XXX
K = Val(Left(XXX, 4))
BG_NAME(K) = Mid(XXX, 6, 30)
If LANG_IND = "INTER" Then BG_NAME(K) = Mid(XXX, 37, 30)
Loop

Close #1
'============================================================================================
Dim LS_NAME(1 To 10000)

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(YYI, "0000") + _
      "M" + Format(MMI, "00") + "_SITES.TXT"

If Dir(fnm) = "" Then Exit Sub

For I = 1 To 10000
LS_NAME(I) = Space(30)
Next I

Open fnm For Input As #1

Do Until EOF(1)

Line Input #1, XXX
K = Val(Left(XXX, 4))
LS_NAME(K) = Mid(XXX, 6, 30)
If LANG_IND = "INTER" Then LS_NAME(K) = Mid(XXX, 37, 30)
Loop

Close #1
'============================================================================================
DBN2 = APPROOT + "\ARTBAS\EXPORT\" + "CUMMULATIVE_ALLDATA.MDB"

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(YYI, "0000") + _
      "M" + Format(MMI, "00") + "_FRAME.TXT"

If Dir(fnm) = "" Then Exit Sub

Open fnm For Input As #1

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(DBN2)
Set prm_record = prm_database.OpenRecordset("G_FRAME_SURVEY")

With prm_record

.Index = "primarykey"

Do Until EOF(1)

Line Input #1, XXX

.AddNew

![COMPOSITE_CODE] = Format(YYI, "0000") + "+" + Format(MMI, "00") + "+"

K = Val(Mid(XXX, 2, 4))

![COMPOSITE_CODE] = ![COMPOSITE_CODE] + Format(ASSOSIMNC(K), "0000") + "+"
![COMPOSITE_CODE] = ![COMPOSITE_CODE] + Mid(XXX, 2, 4) + "+"
![COMPOSITE_CODE] = ![COMPOSITE_CODE] + Mid(XXX, 8, 4)

![MINOR_STRATUM_CODE] = Format(ASSOSIMNC(K), "0000")
![Minor_Stratum_Name] = MINN(ASSOSIMNC(K))

![BOAT_GEAR_CODE] = Mid(XXX, 8, 4)
![BOAT_GEAR_NAME] = BG_NAME(Val(Mid(XXX, 8, 4)))

![SITE_CODE] = Mid(XXX, 2, 4)
![SITE_NAME] = LS_NAME(Val(Mid(XXX, 2, 4)))

![NO_UNITS] = Val(Mid(XXX + Space(20), 26, 15))

.Update

Loop

Close #1

End With

prm_record.Close
prm_database.Close

End Sub
Private Sub CUM_EFFORT()

Call LOAD_ALL_TABLES

If LOAD_TABLES_FLAG <> "OK" Then Exit Sub

Dim fnm, DBN, DBN2, XXX, xcode, xmnc, xact, xsmp, xfrm, xrec, I, J, K, L, m, N, NCODE

DBN2 = APPROOT + "\ARTBAS\EXPORT\" + "CUMMULATIVE_ALLDATA.MDB"

fnm = APPROOT + "\ARTBAS\EFFORT\Y" + Format(YYI, "0000") + _
      "M" + Format(MMI, "00") + "_ESAMPLES.TXT"

If Dir(fnm) = "" Then Exit Sub

Label2.Caption = fnm
Label2.Refresh

Open fnm For Input As #1

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(DBN2)
Set prm_record = prm_database.OpenRecordset("I_BOAT_ACTIVITIES")

With prm_record

.Index = "primarykey"

Do Until EOF(1)

Line Input #1, XXX

.AddNew

![COMPOSITE_CODE] = Format(YYI, "0000") + "+" + Format(MMI, "00") + "+"
![COMPOSITE_CODE] = ![COMPOSITE_CODE] + Mid(XXX, 17, 4) + "+"
![COMPOSITE_CODE] = ![COMPOSITE_CODE] + Mid(XXX, 2, 4) + "+"
![COMPOSITE_CODE] = ![COMPOSITE_CODE] + Mid(XXX, 8, 4) + "+"
![COMPOSITE_CODE] = ![COMPOSITE_CODE] + Mid(XXX, 14, 2)

![MINOR_CODE] = Mid(XXX, 17, 4): ![MINOR_NAME] = MND(Val(Mid(XXX, 17, 4)))
![SITE_CODE] = Mid(XXX, 2, 4): ![SITE_NAME] = SID(Val(Mid(XXX, 2, 4)))
![GEAR_CODE] = Mid(XXX, 8, 4): ![GEAR_NAME] = BGD(Val(Mid(XXX, 8, 4)))

![Day] = Mid(XXX, 14, 2)

![ACTIVE_BOATS] = CDbl(Mid(XXX, 22, 10))
![SAMPLED_BOATS] = CDbl(Mid(XXX, 33, 10))
![FRAME_BOATS] = CDbl(Mid(XXX, 44, 10))
![RECORDER] = Mid(XXX + Space(15), 55, 15)

If ![ACTIVE_BOATS] <> 0 And ![SAMPLED_BOATS] = 0 Then
![SAMPLED_BOATS] = ![FRAME_BOATS]
End If

.Update

Loop

Close #1

End With

prm_record.Close
prm_database.Close

End Sub
Private Sub CUM_ESTIM_REMARKS()

Call LOAD_ALL_TABLES

If LOAD_TABLES_FLAG <> "OK" Then Exit Sub

Dim fnm, DBN, DBN2, XXX, xcode, xmnc, xact, xsmp, xfrm, xrec, I, J, K, L, m, N, NCODE

DBN2 = APPROOT + "\ARTBAS\EXPORT\" + "CUMMULATIVE_ALLDATA.MDB"

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(DBN2)
Set prm_record = prm_database.OpenRecordset("N_ESTIMATION_REMARKS")

For I = 1 To 10000

If MNC(I) = 0 Then GoTo NEXT_I

fnm = APPROOT + "\ARTBAS\RESULTS\Y" + Format(YYI, "0000") + _
      "M" + Format(MMI, "00") + "_MN" + Format(I, "0000") + "_LOG.TXT"

If Dir(fnm) = "" Then GoTo NEXT_I

Label2.Caption = fnm
Label2.Refresh


Open fnm For Input As #1: m = 0

With prm_record

.Index = "primarykey"

Do Until EOF(1)

Line Input #1, XXX: m = m + 1

.AddNew

![COMPOSITE_CODE] = Format(YYI, "0000") + "+" + Format(MMI, "00") + _
                    "+" + Format(I, "0000") + "+" + Format(m, "0000")

![ESTIMATION_REMARK] = LTrim(RTrim(XXX))


.Update

Loop

Close #1

End With

NEXT_I:

Next I

prm_record.Close
prm_database.Close

End Sub
Private Sub CUM_TRIP_TOTALS()

Call LOAD_ALL_TABLES

If LOAD_TABLES_FLAG <> "OK" Then Exit Sub

Dim fnm, DBN, DBN2, XXX, xcode, xmnc, xact, xsmp, xfrm, xrec, I, J, K, L, m, N, NCODE

DBN2 = APPROOT + "\ARTBAS\EXPORT\" + "CUMMULATIVE_ALLDATA.MDB"

fnm = APPROOT + "\ARTBAS\LANDINGS\Y" + Format(YYI, "0000") + _
      "M" + Format(MMI, "00") + "_LSAMPLES.TXT"

If Dir(fnm) = "" Then Exit Sub

Label2.Caption = fnm
Label2.Refresh


Open fnm For Input As #1

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(DBN2)
Set prm_record = prm_database.OpenRecordset("J_TRIP_TOTALS")

With prm_record

.Index = "primarykey"

Do Until EOF(1)

Line Input #1, XXX

.AddNew

![COMPOSITE_CODE] = Format(YYI, "0000") + "+" + Format(MMI, "00") + "+"
![COMPOSITE_CODE] = ![COMPOSITE_CODE] + Mid(XXX, 85, 4) + "+"
![COMPOSITE_CODE] = ![COMPOSITE_CODE] + Mid(XXX, 97, 4) + "+"
![COMPOSITE_CODE] = ![COMPOSITE_CODE] + Mid(XXX, 91, 4) + "+"


![COMPOSITE_CODE] = ![COMPOSITE_CODE] + Mid(XXX, 1, 6)

![MINOR_CODE] = Mid(XXX, 85, 4): ![MINOR_NAME] = MND(Val(Mid(XXX, 85, 4)))
![SITE_CODE] = Mid(XXX, 91, 4): ![SITE_NAME] = SID(Val(Mid(XXX, 91, 4)))
![GEAR_CODE] = Mid(XXX, 97, 4): ![GEAR_NAME] = BGD(Val(Mid(XXX, 97, 4)))

![Day] = Mid(XXX, 8, 2)
![SAMPLE_NO] = Mid(XXX, 1, 6)
![RECORDER] = Mid(XXX, 69, 15)
![NO_UNITS] = CDbl(Mid(XXX, 11, 8))
![Duration] = CDbl(Mid(XXX, 20, 8))
![RAISING_FACTOR] = CDbl(Mid(XXX, 49, 19))
![SAMPLE_TOT] = CDbl(Mid(XXX, 29, 19))

.Update

Loop

Close #1

End With

prm_record.Close
prm_database.Close

End Sub
Private Sub CUM_BY_SPECIES()

Call LOAD_ALL_TABLES

If LOAD_TABLES_FLAG <> "OK" Then Exit Sub

Dim fnm, DBN, DBN2, XXX, xcode, xmnc, xact, xsmp, xfrm, xrec, I, J, K, L, m, N, NCODE

DBN2 = APPROOT + "\ARTBAS\EXPORT\" + "CUMMULATIVE_ALLDATA.MDB"

fnm = APPROOT + "\ARTBAS\LANDINGS\Y" + Format(YYI, "0000") + _
      "M" + Format(MMI, "00") + "_LSPECIES.TXT"

If Dir(fnm) = "" Then Exit Sub

Label2.Caption = fnm
Label2.Refresh


Open fnm For Input As #1

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(DBN2)
Set prm_record = prm_database.OpenRecordset("K_BY_SPECIES")

With prm_record

.Index = "primarykey"

Do Until EOF(1)

Line Input #1, XXX

.AddNew

![COMPOSITE_CODE] = Format(YYI, "0000") + "+" + Format(MMI, "00") + "+"
![COMPOSITE_CODE] = ![COMPOSITE_CODE] + Mid(XXX, 95, 4) + "+"
![COMPOSITE_CODE] = ![COMPOSITE_CODE] + Mid(XXX, 107, 4) + "+"
![COMPOSITE_CODE] = ![COMPOSITE_CODE] + Mid(XXX, 101, 4) + "+"

![COMPOSITE_CODE] = ![COMPOSITE_CODE] + Mid(XXX, 2, 6) + "+"
![COMPOSITE_CODE] = ![COMPOSITE_CODE] + Mid(XXX, 10, 4)

![MINOR_CODE] = Mid(XXX, 95, 4): ![MINOR_NAME] = MND(Val(Mid(XXX, 95, 4)))
![SITE_CODE] = Mid(XXX, 101, 4): ![SITE_NAME] = SID(Val(Mid(XXX, 101, 4)))
![GEAR_CODE] = Mid(XXX, 107, 4): ![GEAR_NAME] = BGD(Val(Mid(XXX, 107, 4)))
![SPECIES_CODE] = Mid(XXX, 10, 4): ![SPECIES_NAME] = SPED(Val(Mid(XXX, 10, 4)))

![SAMPLE_NO] = Mid(XXX, 2, 6)

![catch] = CDbl(Mid(XXX, 15, 19))
![NO_FISH] = CDbl(Mid(XXX, 35, 19))
![price] = CDbl(Mid(XXX, 55, 19))
![Value] = CDbl(Mid(XXX, 75, 19))

.Update

Loop

Close #1

End With

prm_record.Close
prm_database.Close

End Sub
Private Sub CUM_RESULTS()

Call LOAD_ALL_TABLES

If LOAD_TABLES_FLAG <> "OK" Then Exit Sub

Dim fnm, DBN, DBN2, XXX, xcode, xmnc, xact, xsmp, xfrm, xrec, I, J, K, L, m, N, NCODE, GUIDE_CODE

Dim estdes, popn, smpn, BAC_ACCUR, FRNO, actno, cal, eact, _
          esmp, esites, edays, BAC, bac_cvs, bac_cvsp, _
          bac_cvt, bac_cvtp, bac_cv, bac_low, bac_upper, eff, _
          eff_low, eff_upper, nland, CPUE_ACCUR, LPOP, ltot, _
          lsmpv, lsmpf, leff, cpue, lsites, ldays, _
          cpue_cvs, cpue_cvsp, cpue_cvt, cpue_cvtp, cpue_cv, cpue_low, _
          cpue_upper, catch, catch_low, catch_upper, catch_cv, Value, _
          price, fish, kgfish

DBN2 = APPROOT + "\ARTBAS\EXPORT\" + "CUMMULATIVE_ALLDATA.MDB"

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(DBN2)
Set prm_record = prm_database.OpenRecordset("L_RESULTS")

For I = 1 To 10000

If MNC(I) = 0 Then GoTo NEXT_I

fnm = APPROOT + "\ARTBAS\RESULTS\Y" + Format(YYI, "0000") + _
      "M" + Format(MMI, "00") + "_MN" + Format(I, "0000") + "_ESTIM.TXT"

If Dir(fnm) = "" Then GoTo NEXT_I

Label2.Caption = fnm
Label2.Refresh


Open fnm For Input As #1: m = 0

With prm_record

.Index = "primarykey"

Do Until EOF(1)

Input #1, GUIDE_CODE
GUIDE_CODE = LTrim(RTrim(GUIDE_CODE))

If Mid(GUIDE_CODE, 8, 4) <> "0000" And Mid(GUIDE_CODE, 14, 4) = "0000" Then GoTo READ_TRIP
If Mid(GUIDE_CODE, 8, 4) <> "0000" And Mid(GUIDE_CODE, 14, 4) <> "0000" Then GoTo READ_SPECIES
If Mid(GUIDE_CODE, 8, 4) = "0000" And Mid(GUIDE_CODE, 14, 4) <> "0000" Then GoTo READ_SPECIES
If Mid(GUIDE_CODE, 8, 4) = "0000" And Mid(GUIDE_CODE, 14, 4) = "0000" Then GoTo READ_SPECIES

READ_TRIP:

Input #1, estdes, popn, smpn, BAC_ACCUR, FRNO, actno, cal, eact, _
          esmp, esites, edays, BAC, bac_cvs, bac_cvsp, _
          bac_cvt, bac_cvtp, bac_cv, bac_low, bac_upper, eff, _
          eff_low, eff_upper, nland, CPUE_ACCUR, LPOP, ltot, _
          lsmpv, lsmpf, leff, cpue, lsites, ldays, _
          cpue_cvs, cpue_cvsp, cpue_cvt, cpue_cvtp, cpue_cv, cpue_low, _
          cpue_upper, catch, catch_low, catch_upper, catch_cv, Value, _
          price, fish, kgfish

If Mid(GUIDE_CODE, 8, 4) = "0000" Or catch + eff = 0 Then GoTo CONT_READ

.AddNew

![COMPOSITE_CODE] = Format(YYI, "0000") + "+" + Format(MMI, "00") + "+"
![COMPOSITE_CODE] = ![COMPOSITE_CODE] + Format(I, "0000") + "+"
![COMPOSITE_CODE] = ![COMPOSITE_CODE] + Mid(GUIDE_CODE, 8, 4) + "+"
![COMPOSITE_CODE] = ![COMPOSITE_CODE] + Mid(GUIDE_CODE, 14, 4)

![MINOR_CODE] = Format(I, "0000"): ![MINOR_NAME] = MND(I)
![GEAR_CODE] = Mid(GUIDE_CODE, 8, 4): ![GEAR_NAME] = BGD(Val(Mid(GUIDE_CODE, 8, 4)))
![SPECIES_CODE] = "0000": ![SPECIES_NAME] = msgtab(201)

![catch] = catch
![effort] = eff
![cpue] = cpue
'If kgfish <> 0 Then ![aver_weight] = kgfish
'If fish <> 0 Then ![NO_FISH] = fish
'If price <> 0 Then ![price] = price
'If Value <> 0 Then ![Value] = Value
![BAC] = BAC
![ACCUR_BAC] = 100 * BAC_ACCUR
![ACCUR_CPUE] = 100 * CPUE_ACCUR

.Update

GoTo CONT_READ

READ_SPECIES:

Input #1, estdes, eff, cpue, catch, Value, price, kgfish, _
           fish, FRNO

If Mid(GUIDE_CODE, 8, 4) = "0000" Or catch + eff = 0 Then GoTo CONT_READ

.AddNew

![COMPOSITE_CODE] = Format(YYI, "0000") + "+" + Format(MMI, "00") + "+"
![COMPOSITE_CODE] = ![COMPOSITE_CODE] + Format(I, "0000") + "+"
![COMPOSITE_CODE] = ![COMPOSITE_CODE] + Mid(GUIDE_CODE, 8, 4) + "+"
![COMPOSITE_CODE] = ![COMPOSITE_CODE] + Mid(GUIDE_CODE, 14, 4)

![MINOR_CODE] = Format(I, "0000"): ![MINOR_NAME] = MND(I)
![GEAR_CODE] = Mid(GUIDE_CODE, 8, 4): ![GEAR_NAME] = BGD(Val(Mid(GUIDE_CODE, 8, 4)))
![SPECIES_CODE] = Mid(GUIDE_CODE, 14, 4): ![SPECIES_NAME] = SPED(Val(Mid(GUIDE_CODE, 14, 4)))


![catch] = catch
![effort] = eff
![cpue] = cpue
If kgfish <> 0 Then ![aver_weight] = kgfish
If fish <> 0 Then ![NO_FISH] = fish
If price <> 0 Then ![price] = price
If Value <> 0 Then ![Value] = Value

.Update

CONT_READ:

Loop

Close #1

End With

NEXT_I:

Next I

prm_record.Close
prm_database.Close

End Sub
