VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmARTS00 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7515
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   10860
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
   ForeColor       =   &H00008000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   3  'I-Beam
   Moveable        =   0   'False
   ScaleHeight     =   7515
   ScaleWidth      =   10860
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdEXIT 
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
      Left            =   9960
      MousePointer    =   1  'Arrow
      Picture         =   "ARTS00.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   6480
      Width           =   735
   End
   Begin VB.CommandButton cmdENTER 
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
      Left            =   7440
      MousePointer    =   1  'Arrow
      Picture         =   "ARTS00.frx":0282
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   6480
      Width           =   735
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
      Height          =   735
      Left            =   8280
      MousePointer    =   1  'Arrow
      Picture         =   "ARTS00.frx":0504
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6480
      Width           =   735
   End
   Begin VB.CommandButton cmdSERVICES 
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
      Left            =   9120
      MousePointer    =   1  'Arrow
      Picture         =   "ARTS00.frx":2766
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   6480
      Width           =   735
   End
   Begin VB.CheckBox chkLOCAL 
      Caption         =   "Check1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   10320
      TabIndex        =   16
      Top             =   6240
      Width           =   135
   End
   Begin VB.CheckBox chkSPANISH 
      Caption         =   "Check1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   9480
      TabIndex        =   15
      Top             =   6240
      Width           =   135
   End
   Begin VB.CheckBox chkFRENCH 
      Caption         =   "Check1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   8640
      TabIndex        =   14
      Top             =   6240
      Width           =   135
   End
   Begin VB.CheckBox chkENGLISH 
      Caption         =   "Check1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7800
      TabIndex        =   13
      Top             =   6240
      Value           =   1  'Checked
      Width           =   135
   End
   Begin VB.CommandButton cmdLOCAL 
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
      Height          =   495
      Left            =   9960
      MousePointer    =   1  'Arrow
      Picture         =   "ARTS00.frx":4558
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5760
      Width           =   735
   End
   Begin VB.CommandButton cmdSPANISH 
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
      Height          =   495
      Left            =   9120
      MousePointer    =   1  'Arrow
      Picture         =   "ARTS00.frx":47DA
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5760
      Width           =   735
   End
   Begin VB.CommandButton cmdFRENCH 
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
      Height          =   495
      Left            =   8280
      MousePointer    =   1  'Arrow
      Picture         =   "ARTS00.frx":4A5C
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5760
      Width           =   735
   End
   Begin VB.CommandButton cmdENGLISH 
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
      Height          =   495
      Left            =   7440
      MousePointer    =   1  'Arrow
      Picture         =   "ARTS00.frx":4CDE
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5760
      Width           =   735
   End
   Begin VB.OptionButton optEXP 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Option1"
      Height          =   375
      Left            =   3240
      TabIndex        =   8
      Top             =   6840
      Width           =   3975
   End
   Begin VB.OptionButton optIMP 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Option1"
      Height          =   375
      Left            =   3240
      TabIndex        =   7
      Top             =   6480
      Width           =   3975
   End
   Begin ComctlLib.ProgressBar pgbCR 
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   4200
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.ListBox lstYEAR 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   1710
      Left            =   120
      TabIndex        =   5
      Top             =   5520
      Width           =   2415
   End
   Begin VB.ListBox lstDISP 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1335
      Left            =   120
      TabIndex        =   3
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Label lblQUIT 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3240
      TabIndex        =   22
      Top             =   7200
      Width           =   3975
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FFFF&
      Caption         =   "01"
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
      Left            =   120
      TabIndex        =   21
      Top             =   7200
      Width           =   255
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Select an ARTSER database"
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   4920
      Width           =   2415
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Height          =   720
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "              1"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   2415
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "     123456789012"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   2160
      Width           =   2415
   End
End
Attribute VB_Name = "frmARTS00"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private CTAB(), ADDFISH, ADDVALUE
Private MJC(), MJN(), MNC(), MNN(), ASSO(), BGC(), BGN(), SPEC(), SPEN()
Private MJN2(), MNN2(), BGN2(), SPEN2()
Private YTAB(), YTAB2(), NY, CURMN, CURMJ
Private NSOUT, NSOUS, READY_ERROR, EMPTY_ERROR, OLDY, COPYY

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias _
   "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, _
   ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) _
   As Long
   
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias _
   "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
   ByVal lpReserved As Long, lpType As Long, _
   ByVal lpData As String, lpcbData As Long) As Long
               'Note that if you declare the lpData parameter as String,
               'you must pass it ByVal.
                                                                                   
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Const REG_SZ As Long = 1
Const KEY_ALL_ACCESS = &H3F
Const HKEY_LOCAL_MACHINE = &H80000002



Private Sub chkENGLISH_Click()
chkENGLISH.Value = 1
End Sub
Private Sub chkFRENCH_Click()
chkFRENCH.Value = 1
End Sub
Private Sub chkLOCAL_Click()
chkLOCAL.Value = 1
End Sub
Private Sub chkSPANISH_Click()
chkSPANISH.Value = 1
End Sub
Private Sub cmdENTER_Click()

frmARTS00.MousePointer = 13
cmdENTER.MousePointer = 13

pgbCR.Visible = True
pgbCR.Max = 2

Call CREATE_CURRENT_TOT
pgbCR.Value = 1

Call CREATE_CURRENT_SPECIES
pgbCR.Value = 2

frmARTS00.MousePointer = 1
cmdENTER.MousePointer = 1

' TNJ - I am making the progress bar invisible for now to make sure that
' when we hide this form, it will not be confusing for the users when
' the form frmARTS00 is visible again
pgbCR.Visible = False

Load frmSEL
' Unload frmARTS00
frmARTS00.Hide

frmSEL.Show

End Sub
Private Sub cmdEXIT_Click()

cmdEXIT.MousePointer = 13
frmARTS00.MousePointer = 13

Call write_parms

End

End Sub
Private Sub cmdGUIDE_Click()

optIMP.Visible = False
optEXP.Visible = False

HTYPE = "10"

HFNM = APPROOT + "\ARTS\HELP\" + current_language + "HELP" + HTYPE + ".rtf"

If Dir(HFNM) = "" Then Exit Sub

frmARTS00.Enabled = False
Load frmGUIDE
frmGUIDE.Show

End Sub
Private Sub cmdSERVICES_Click()

optIMP.Visible = True: optIMP.Value = False
optEXP.Visible = True: optEXP.Value = False
lblQUIT.Visible = True
   
End Sub
Private Sub Form_Load()

COMPVAL = "YES": ADDVALUE = "YES"

Dim XXX, LLL

XXX = CurDir()

LLL = InStr(XXX, "\ARTS")

' TNJ Note: If LLL is 0 (zero), the directory being executed from
' is not "\ARTS"
If LLL = 0 Then
   MsgBox "The directory where the ArtSer executable is being run is not '\ARTS' !!!"
   MsgBox "The directory is: " + XXX
   End
End If

APPROOT = GetShortName(Left(XXX, LLL - 1))

' Open "C:\ARTSER_CURRENT_FOLDER.TXT" For Output As #1
' Print #1, APPROOT
' Close #1

Open APPROOT + "\ARTS\CONTROL\COMPUTE.TXT" For Output As #1
Print #1, "NOCOMPUTE"
Close #1

Set Picture = LoadPicture(APPROOT + "\ARTS\PICS_RUNTIME\SCREEN_01.JPG")

Dim frun

OLDY = 9999

pgbCR.Visible = False
Label3.Visible = False
Label4.Visible = False
Label5.Visible = False
Label6.Visible = False
lstDISP.Visible = False
lstYEAR.Visible = False
optIMP.Visible = False
optEXP.Visible = False
lblQUIT.Visible = False

cmdENTER.Enabled = False
current_month = 999

frmARTS00.MousePointer = 1

Call read_parms

language = current_language

If language = "E" Then
    chkENGLISH.Visible = True
    chkENGLISH.Value = 1
    chkFRENCH.Visible = False
    chkSPANISH.Visible = False
    chkLOCAL.Visible = False
End If

If language = "F" Then
    chkENGLISH.Visible = False
    chkFRENCH.Value = 1
    chkFRENCH.Visible = True
    chkSPANISH.Visible = False
    chkLOCAL.Visible = False
End If

If language = "S" Then
    chkENGLISH.Visible = False
    chkSPANISH.Value = 1
    chkFRENCH.Visible = False
    chkSPANISH.Visible = True
    chkLOCAL.Visible = False
End If

If language = "L" Then
    chkENGLISH.Visible = False
    chkLOCAL.Value = 1
    chkFRENCH.Visible = False
    chkSPANISH.Visible = False
    chkLOCAL.Visible = True
End If

Call MSGLOAD

frmARTS00.Caption = msgtab(4)

cmdENGLISH.ToolTipText = msgtab_e
cmdFRENCH.ToolTipText = msgtab_f
cmdSPANISH.ToolTipText = msgtab_s
cmdLOCAL.ToolTipText = msgtab_l
cmdENTER.ToolTipText = msgtab(2)
cmdEXIT.ToolTipText = msgtab(3)
cmdGUIDE.ToolTipText = msgtab(6)
cmdSERVICES.ToolTipText = msgtab(30)
optIMP.Caption = msgtab(35)
optEXP.Caption = msgtab(36)

Label5.Caption = msgtab(5)
Label6.Caption = msgtab(7)
lblQUIT.Caption = msgtab(124)

' Call APPROOT_WRITE

Call CHECK_ARTBASIC
Call CHECK_ARTSER
Call LOCATE_EXCEL

End Sub
Private Sub CHECK_ARTBASIC()

frmARTS00.MousePointer = 13

Call TRANSFER_FROM_TRANSFER

frmARTS00.MousePointer = 1

NY = 0

If Dir(APPROOT + "\ARTBAS\RESULTS\*.*") = "" Then Exit Sub

Label3.Visible = True
Label4.Visible = True
Label5.Visible = True

lstDISP.Visible = True

Dim I, ix, fnm, J, jx, XXX

For I = 1990 To 2020

ix = Format(I, "0000"): XXX = ""

For J = 1 To 12

jx = Format(J, "00")

fnm = APPROOT + "\ARTBAS\RESULTS\Y" + ix + "M" + jx + "_ESTIM.TXT"

If Dir(fnm) = "" Then
   XXX = XXX + "-"
   GoTo CONT_J
   End If

XXX = XXX + "X"

CONT_J:

Next J

If XXX = "------------" Then GoTo CONT_I

NY = NY + 1

ReDim Preserve YTAB(1 To NY)

YTAB(NY) = I

lstDISP.AddItem ix + " " + XXX

CONT_I:

Next I

End Sub
Private Sub CHECK_ARTSER()

Dim YYNN

YYNN = 0

If Dir(APPROOT + "\ARTS\DATA\*.*") = "" Then Exit Sub

Label6.Visible = True
lstYEAR.Visible = True

Dim I, ix, fnm

For I = 1990 To 2020

ix = Format(I, "0000")

fnm = APPROOT + "\ARTS\DATA\A" + ix + ".TXT"

If Dir(fnm) = "" Then GoTo CONT_I

YYNN = YYNN + 1

ReDim Preserve YTAB2(1 To YYNN)

YTAB2(YYNN) = I

lstYEAR.AddItem ix

CONT_I:

Next I

End Sub
Private Sub read_parms()

Dim textline As String * 80, ll As Integer

Open APPROOT + "\ARTS\CONTROL\SYSPARM.TXT" For Input As #1

Input #1, current_language

Input #1, current_year

Close #1

End Sub
Private Sub cmdLOCAL_Click()
language = "L": current_language = language
chkENGLISH.Visible = False
chkFRENCH.Visible = False
chkSPANISH.Visible = False
chkLOCAL.Visible = True
chkLOCAL.Value = 1
Call MSGLOAD
cmdGUIDE.ToolTipText = msgtab(6)
cmdENTER.ToolTipText = msgtab(2)
cmdEXIT.ToolTipText = msgtab(3)
frmARTS00.Caption = msgtab(4)
Label5.Caption = msgtab(5)
Label6.Caption = msgtab(7)
cmdSERVICES.ToolTipText = msgtab(30)
optIMP.Caption = msgtab(35)
optEXP.Caption = msgtab(36)

End Sub
Private Sub cmdENGLISH_Click()
language = "E": current_language = language
chkENGLISH.Visible = True
chkENGLISH.Value = 1
chkFRENCH.Visible = False
chkSPANISH.Visible = False
chkLOCAL.Visible = False
Call MSGLOAD
cmdGUIDE.ToolTipText = msgtab(6)
cmdENTER.ToolTipText = msgtab(2)
cmdEXIT.ToolTipText = msgtab(3)
frmARTS00.Caption = msgtab(4)
Label5.Caption = msgtab(5)
Label6.Caption = msgtab(7)
cmdSERVICES.ToolTipText = msgtab(30)
optIMP.Caption = msgtab(35)
optEXP.Caption = msgtab(36)

End Sub
Private Sub cmdFRENCH_Click()
language = "F": current_language = language
chkENGLISH.Visible = False
chkFRENCH.Visible = True
chkFRENCH.Value = 1
chkSPANISH.Visible = False
chkLOCAL.Visible = False
Call MSGLOAD
cmdGUIDE.ToolTipText = msgtab(6)
cmdENTER.ToolTipText = msgtab(2)
cmdEXIT.ToolTipText = msgtab(3)
frmARTS00.Caption = msgtab(4)
Label5.Caption = msgtab(5)
Label6.Caption = msgtab(7)
cmdSERVICES.ToolTipText = msgtab(30)
optIMP.Caption = msgtab(35)
optEXP.Caption = msgtab(36)

End Sub
Private Sub cmdSPANISH_Click()
language = "S": current_language = language
chkENGLISH.Visible = False
chkFRENCH.Visible = False
chkSPANISH.Visible = True
chkSPANISH.Value = 1
chkLOCAL.Visible = False
Call MSGLOAD
cmdGUIDE.ToolTipText = msgtab(6)
cmdENTER.ToolTipText = msgtab(2)
cmdEXIT.ToolTipText = msgtab(3)
frmARTS00.Caption = msgtab(4)
Label5.Caption = msgtab(5)
Label6.Caption = msgtab(7)
cmdSERVICES.ToolTipText = msgtab(30)
optIMP.Caption = msgtab(35)
optEXP.Caption = msgtab(36)

End Sub

Private Sub lblQUIT_Click()

lblQUIT.Visible = False
optIMP.Visible = False
optEXP.Visible = False

End Sub
Private Sub lstDISP_Click()

Dim J

ReDim VALID_MONTHS(1 To 12)

For J = 1 To 12
VALID_MONTHS(J) = "-"
Next J

CURY = YTAB(lstDISP.ListIndex + 1)

Dim fnm, resp

fnm = APPROOT + "\ARTS\DATA\A" + Format(CURY, "0000") + ".TXT"

If Dir(fnm) = "" Then GoTo NO_CHECK

resp = MsgBox(msgtab(8), vbCritical + vbOKCancel, " ")

If resp <> 1 Then
   lstDISP.Clear
   Call CHECK_ARTBASIC
   Exit Sub
   End If

NO_CHECK:

lstDISP.Enabled = False
lstYEAR.Enabled = False

pgbCR.Visible = True

pgbCR.Min = 1: pgbCR.Max = 30

lstDISP.MousePointer = 13
frmARTS00.MousePointer = 13

Call SETUP_TABLES

pgbCR.Value = 1

Call CREATE_LOOP
Call DUMP_TOT
Call DUMP_SPECIES2

Call INSPECT_DB
pgbCR.Value = 26

Call FINISH_MAJOR
pgbCR.Value = 27

Call FINISH_MINOR
pgbCR.Value = 28

Call FINISH_BG
pgbCR.Value = 29

Call FINISH_SPECIES
pgbCR.Value = 30

lstDISP.MousePointer = 1
frmARTS00.MousePointer = 1

lstDISP.Enabled = True
pgbCR.Visible = False
lstYEAR.Enabled = True

lstDISP.Clear
lstYEAR.Clear

Call CHECK_ARTBASIC
Call CHECK_ARTSER

End Sub
Private Sub DUMP_TOT()

Dim TW(), TF()
Dim TC(), TE(), TU(), TP(), TV(), I, ix, FFC, FFE, FFP, FFU, FFV, RR, cc, PP
Dim FFW, FFN

ReDim TC(1 To 13), TE(1 To 13), TU(1 To 13), TP(1 To 13), TV(1 To 13)
ReDim TW(1 To 13), TF(1 To 13)

Dim DBN, fnm

DBN = APPROOT + "\ARTS\DATA\A" + Format(CURY, "0000") + ".MDB"

If Dir(DBN) = "" Then Exit Sub

fnm = APPROOT + "\ARTS\DATA\A" + Format(CURY, "0000") + ".TXT"
Open fnm For Output As #2

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(DBN)
Set prm_record = prm_database.OpenRecordset("ASITAB")

With prm_record

.Index = "primarykey"
.MoveFirst

Do Until .EOF

TC(13) = 0: TE(13) = 0: TV(13) = 0: TF(13) = 0

For I = 1 To 12

ix = Format(I, "00")

FFC = "C" + ix: TC(I) = .Fields(FFC)
FFE = "E" + ix: TE(I) = .Fields(FFE)
FFU = "U" + ix: TU(I) = .Fields(FFU)
FFP = "P" + ix: TP(I) = .Fields(FFP)
FFV = "V" + ix: TV(I) = .Fields(FFV)
FFW = "W" + ix: TW(I) = .Fields(FFW)
FFN = "F" + ix: TF(I) = .Fields(FFN)

TC(13) = TC(13) + TC(I)
TE(13) = TE(13) + TE(I)
TV(13) = TV(13) + TV(I)
TF(13) = TF(13) + TF(I)

Next I

If TE(13) <> 0 Then TU(13) = TC(13) / TE(13)
If TC(13) <> 0 Then TP(13) = TV(13) / TC(13)
If TF(13) <> 0 Then TW(13) = TC(13) / TF(13)

RR = ![RANK]: cc = ![CUM]: PP = ![PER]

ADDFISH = "NO"

Write #2, ![akey], TC(1), TC(2), TC(3), TC(4), TC(5), TC(6), TC(7), TC(8), TC(9), TC(10), TC(11), TC(12), TC(13), _
                   TE(1), TE(2), TE(3), TE(4), TE(5), TE(6), TE(7), TE(8), TE(9), TE(10), TE(11), TE(12), TE(13), _
                   TU(1), TU(2), TU(3), TU(4), TU(5), TU(6), TU(7), TU(8), TU(9), TU(10), TU(11), TU(12), TU(13), _
                   TP(1), TP(2), TP(3), TP(4), TP(5), TP(6), TP(7), TP(8), TP(9), TP(10), TP(11), TP(12), TP(13), _
                   TV(1), TV(2), TV(3), TV(4), TV(5), TV(6), TV(7), TV(8), TV(9), TV(10), TV(11), TV(12), TV(13), _
                   TW(1), TW(2), TW(3), TW(4), TW(5), TW(6), TW(7), TW(8), TW(9), TW(10), TW(11), TW(12), TW(13), _
                   TF(1), TF(2), TF(3), TF(4), TF(5), TF(6), TF(7), TF(8), TF(9), TF(10), TF(11), TF(12), TF(13), _
                   RR, cc, PP, ADDFISH, ADDVALUE
.MoveNext

Loop

End With

Close #2

prm_record.Close
prm_database.Close
                   
Kill DBN
                   
End Sub
Private Sub DUMP_SPECIES2()

Dim TW(), TF()
Dim TC(), TE(), TU(), TP(), TV(), I, ix, FFC, FFE, FFP, FFU, FFV, RR, cc, PP
Dim FFW, FFN
Dim TCP(1 To 13), TCF(1 To 13)

ReDim TC(1 To 13), TE(1 To 13), TU(1 To 13), TP(1 To 13), TV(1 To 13)
ReDim TW(1 To 13), TF(1 To 13)

Dim DBN, fnm

DBN = APPROOT + "\ARTS\DATA\S" + Format(CURY, "0000") + ".MDB"

If Dir(DBN) = "" Then Exit Sub

fnm = APPROOT + "\ARTS\DATA\S" + Format(CURY, "0000") + ".TXT"
Open fnm For Output As #2

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(DBN)
Set prm_record = prm_database.OpenRecordset("ASITAB")

With prm_record

.Index = "primarykey"
.MoveFirst

Do Until .EOF

TC(13) = 0: TE(13) = 0: TV(13) = 0: TF(13) = 0
TCP(13) = 0: TCF(13) = 0

For I = 1 To 12

ix = Format(I, "00")

FFC = "C" + ix: TC(I) = .Fields(FFC)
FFE = "E" + ix: TE(I) = .Fields(FFE)
FFU = "U" + ix: TU(I) = .Fields(FFU)
FFP = "P" + ix: TP(I) = .Fields(FFP)
FFV = "V" + ix: TV(I) = .Fields(FFV)
FFW = "W" + ix: TW(I) = .Fields(FFW)
FFN = "F" + ix: TF(I) = .Fields(FFN)

TC(13) = TC(13) + TC(I)
TE(13) = TE(13) + TE(I)

If TP(I) <> 0 Then
   TV(13) = TV(13) + TV(I): TCP(13) = TCP(13) + TC(I)
   End If
   
TF(13) = TF(13) + TF(I)
If TF(I) <> 0 Then TCF(13) = TCF(13) + TC(I)
 
If TP(I) = 0 And TC(I) <> 0 Then COMPVAL = "NO"
 
Next I

If TE(13) <> 0 Then TU(13) = TC(13) / TE(13)
If TCP(13) <> 0 Then TP(13) = TV(13) / TCP(13)
If TCF(13) <> 0 And TCF(13) = TC(13) Then TW(13) = TCF(13) / TF(13)

ADDFISH = "YES"

If TCF(13) <> TC(13) Then
   ADDFISH = "NO"
   TW(13) = 0: TF(13) = 0
   End If

ADDVALUE = "YES"

If TCP(13) <> TC(13) Then
   ADDVALUE = "NO"
   TP(13) = 0: TV(13) = 0
   End If

If TV(13) = 0 Then TP(13) = 0
If TF(13) = 0 Then TW(13) = 0

RR = ![RANK]: cc = ![CUM]: PP = ![PER]

Write #2, ![akey], TC(1), TC(2), TC(3), TC(4), TC(5), TC(6), TC(7), TC(8), TC(9), TC(10), TC(11), TC(12), TC(13), _
                   TE(1), TE(2), TE(3), TE(4), TE(5), TE(6), TE(7), TE(8), TE(9), TE(10), TE(11), TE(12), TE(13), _
                   TU(1), TU(2), TU(3), TU(4), TU(5), TU(6), TU(7), TU(8), TU(9), TU(10), TU(11), TU(12), TU(13), _
                   TP(1), TP(2), TP(3), TP(4), TP(5), TP(6), TP(7), TP(8), TP(9), TP(10), TP(11), TP(12), TP(13), _
                   TV(1), TV(2), TV(3), TV(4), TV(5), TV(6), TV(7), TV(8), TV(9), TV(10), TV(11), TV(12), TV(13), _
                   TW(1), TW(2), TW(3), TW(4), TW(5), TW(6), TW(7), TW(8), TW(9), TW(10), TW(11), TW(12), TW(13), _
                   TF(1), TF(2), TF(3), TF(4), TF(5), TF(6), TF(7), TF(8), TF(9), TF(10), TF(11), TF(12), TF(13), _
                   RR, cc, PP, ADDFISH, ADDVALUE

.MoveNext

Loop

End With

Close #2

Open APPROOT + "\ARTS\CONTROL\COMPVAL.TXT" For Output As #1
Print #1, COMPVAL
Close #1

prm_record.Close
prm_database.Close
                  
Kill DBN
                  
End Sub
Private Sub SETUP_TABLES()

Dim I, ix, J, jx, fnm

ix = Format(CURY, "0000")

ReDim MJC(1 To 10000), MJN(1 To 10000), MJN2(1 To 1000)
ReDim MNC(1 To 10000), MNN(1 To 10000), MNN2(1 To 1000)
ReDim BGC(1 To 10000), BGN(1 To 10000), BGN2(1 To 1000)
ReDim SPEC(1 To 10000), SPEN(1 To 10000), SPEN2(1 To 1000)
ReDim ASSO(1 To 10000)

For J = 1 To 10000
ASSO(J) = 0
Next J

For J = 1 To 12

CURM = J

Call DUMP_UNITS

jx = Format(J, "00")

fnm = APPROOT + "\ARTBAS\RESULTS\Y" + ix + "M" + jx + "_ESTIM.TXT"

If Dir(fnm) = "" Then GoTo NEXT_J:

Call LOAD_MAJOR
Call LOAD_MINOR
Call LOAD_ASSO

Call LOAD_BG
Call LOAD_SPECIES

NEXT_J:

Next J

Call DUMP_MAJOR
Call DUMP_MINOR
Call DUMP_BG
Call DUMP_SPECIES

End Sub
Private Sub LOAD_MAJOR()

Dim fnm, I, XXX

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(CURY, "0000") + "M" + Format(CURM, "00") + "_MAJOR.TXT"

If Dir(fnm) = "" Then Exit Sub

Open fnm For Input As #1

Do Until EOF(1)

Line Input #1, XXX

I = Val(Left(XXX, 4))

MJC(I) = "Y": MJN(I) = Mid(XXX, 6, 30): MJN2(I) = Mid(XXX, 37, 30)
If Left(MJN2(I), 3) = "..." Or Left(MJN2(I), 2) = "   " Then MJN2(I) = "[FAO ?] " + MJN(I)

Loop

Close #1

End Sub
Private Sub LOAD_MINOR()

Dim fnm, I, XXX

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(CURY, "0000") + "M" + Format(CURM, "00") + _
      "_MINOR.TXT"

If Dir(fnm) = "" Then Exit Sub

Open fnm For Input As #1

Do Until EOF(1)

Line Input #1, XXX

I = Val(Left(XXX, 4))

MNC(I) = "Y": MNN(I) = Mid(XXX, 6, 30): MNN2(I) = Mid(XXX, 37, 30)
If Left(MNN2(I), 3) = "..." Or Left(MNN2(I), 2) = "   " Then MNN2(I) = "[FAO ?] " + MNN(I)

Loop

Close #1

End Sub
Private Sub LOAD_ASSO()

Dim fnm, I, XXX, K, J, YYY, m

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(CURY, "0000") + "M" + Format(CURM, "00") + _
      "_ASSOMN.TXT"

If Dir(fnm) = "" Then Exit Sub

Open fnm For Input As #1

Do Until EOF(1)

Line Input #1, XXX

K = Val(Mid(XXX, 37, 4))

For J = 1 To K

Line Input #1, YYY

m = Val(Mid(YYY, 6, 4))

ASSO(m) = Val(Left(XXX, 4))

Next J

Loop

Close #1

End Sub
Private Sub LOAD_BG()

Dim fnm, I, XXX

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(CURY, "0000") + "M" + Format(CURM, "00") + _
      "_BG.TXT"

If Dir(fnm) = "" Then Exit Sub

Open fnm For Input As #1

Do Until EOF(1)

Line Input #1, XXX

I = Val(Left(XXX, 4))

BGC(I) = "Y": BGN(I) = Mid(XXX, 6, 30): BGN2(I) = Mid(XXX, 37, 30)
If Left(BGN2(I), 3) = "..." Or Left(BGN2(I), 2) = "   " Then BGN2(I) = "[FAO ?] " + BGN(I)

Loop

Close #1

End Sub
Private Sub LOAD_SPECIES()

Dim fnm, I, XXX

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(CURY, "0000") + "M" + Format(CURM, "00") + _
      "_SPECIES.TXT"

If Dir(fnm) = "" Then Exit Sub

Open fnm For Input As #1

Do Until EOF(1)

Line Input #1, XXX

I = Val(Left(XXX, 4))

SPEC(I) = "Y": SPEN(I) = Mid(XXX, 6, 30): SPEN2(I) = Mid(XXX, 37, 30)
If Left(SPEN2(I), 3) = "..." Or Left(SPEN2(I), 2) = "   " Then SPEN2(I) = "[FAO ?] " + SPEN(I)

Loop

Close #1

End Sub
Private Sub DUMP_MAJOR()

Dim fnm, I, XXX

fnm = APPROOT + "\ARTS\TABLES\Y" + Format(CURY, "0000") + "_MAJOR.TXT"

Open fnm For Output As #1

For I = 1 To 10000

If MJC(I) <> "Y" Then GoTo NEXT_I

Print #1, Format(I, "0000") + " " + MJN(I) + " " + MJN2(I)

NEXT_I:

Next I

Close #1

End Sub
Private Sub DUMP_MINOR()

Dim fnm, I, XXX

fnm = APPROOT + "\ARTS\TABLES\Y" + Format(CURY, "0000") + "_MINOR.TXT"

Open fnm For Output As #1

For I = 1 To 10000

If MNC(I) <> "Y" Or ASSO(I) = 0 Then GoTo NEXT_I

Print #1, Format(I, "0000") + " " + MNN(I) + " " + Format(ASSO(I), "0000") + " " + _
          MJN(ASSO(I)) + " " + MJN2(I) + " " + MNN2(I)

NEXT_I:

Next I

Close #1

End Sub
Private Sub DUMP_BG()

Dim fnm, I, XXX

fnm = APPROOT + "\ARTS\TABLES\Y" + Format(CURY, "0000") + "_BG.TXT"

Open fnm For Output As #1

For I = 1 To 10000

If BGC(I) <> "Y" Then GoTo NEXT_I

Print #1, Format(I, "0000") + " " + BGN(I) + " " + BGN2(I)

NEXT_I:

Next I

Close #1

End Sub
Private Sub DUMP_SPECIES()

Dim fnm, I, XXX

fnm = APPROOT + "\ARTS\TABLES\Y" + Format(CURY, "0000") + "_SPECIES.TXT"

Open fnm For Output As #1

For I = 1 To 10000

If SPEC(I) <> "Y" Then GoTo NEXT_I

Print #1, Format(I, "0000") + " " + SPEN(I) + " " + SPEN2(I)

NEXT_I:

Next I

Close #1

End Sub
Private Sub DUMP_UNITS()

Dim fnm1, fnm2, I, XXX

fnm1 = APPROOT + "\ARTBAS\TABLES\Y" + Format(CURY, "0000") + _
       "M" + Format(CURM, "00") + "_UNITS.TXT"
fnm2 = APPROOT + "\ARTS\TABLES\Y" + Format(CURY, "0000") + "_UNITS.TXT"

If Dir(fnm1) <> "" Then FileCopy fnm1, fnm2

End Sub
Private Sub UPDATE_TOT()

Dim I, fnm, ix, XKEY, XXX, FFF, FFC, FFE, FFU, FFP, FFV, FFW, FFN

fnm = APPROOT + "\ARTBAS\RESULTS\Y" + Format(CURY, "0000") + _
      "M" + Format(CURM, "00") + "_MN" + Format(CURMN, "0000") + "_ESTIM.TXT"

If Dir(fnm) = "" Then Exit Sub

VALID_MONTHS(CURM) = "X"

Dim estdes, popn, smpn, BAC_ACCUR, FRNO, actno, cal, eact, _
          esmp, esites, edays, bac, bac_cvs, bac_cvsp, _
          bac_cvt, bac_cvtp, bac_cv, bac_low, bac_upper, eff, _
          eff_low, eff_upper, nland, CPUE_ACCUR, LPOP, ltot, _
          lsmpv, lsmpf, leff, cpue, lsites, ldays, _
          cpue_cvs, cpue_cvsp, cpue_cvt, cpue_cvtp, cpue_cv, cpue_low, _
          cpue_upper, catch, catch_low, catch_upper, catch_cv, Value, _
          price, fish, kgfish


Dim DBN, CREC

DBN = APPROOT + "\ARTS\DATA\A" + Format(CURY, "0000") + ".MDB"

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(DBN)
Set prm_record = prm_database.OpenRecordset("ASITAB")

With prm_record

.Index = "primarykey"

Dim TYPEREC

Open fnm For Input As #1

Do Until EOF(1)

Input #1, XKEY

TYPEREC = 2

If Mid(XKEY, 14, 4) = "0000" And Mid(XKEY, 8, 4) <> "0000" Then TYPEREC = 1

If TYPEREC <> 1 Then
   Input #1, estdes, eff, cpue, catch, Value, price, kgfish, fish, FRNO
   GoTo READ_NEXT
   End If
   
Input #1, estdes, popn, smpn, BAC_ACCUR, FRNO, actno, cal, eact, _
          esmp, esites, edays, bac, bac_cvs, bac_cvsp, _
          bac_cvt, bac_cvtp, bac_cv, bac_low, bac_upper, eff, _
          eff_low, eff_upper, nland, CPUE_ACCUR, LPOP, ltot, _
          lsmpv, lsmpf, leff, cpue, lsites, ldays, _
          cpue_cvs, cpue_cvsp, cpue_cvt, cpue_cvtp, cpue_cv, cpue_low, _
          cpue_upper, catch, catch_low, catch_upper, catch_cv, Value, _
          price, fish, kgfish

If catch + eff = 0 Then GoTo READ_NEXT
         
XKEY = "J" + Format(CURMJ, "0000") + "+" + XKEY
          
.Seek "=", XKEY

If .NoMatch = True Then GoTo ADD_NEW
If .NoMatch = False Then GoTo UPDATE_REC

ADD_NEW:

.AddNew

![akey] = XKEY

For I = 1 To 13

ix = Format(I, "00")

FFC = "C" + ix: .Fields(FFC) = 0
FFE = "E" + ix: .Fields(FFE) = 0
FFU = "U" + ix: .Fields(FFU) = 0
FFP = "P" + ix: .Fields(FFP) = 0
FFV = "V" + ix: .Fields(FFV) = 0
FFW = "W" + ix: .Fields(FFW) = 0
FFN = "F" + ix: .Fields(FFN) = 0

Next I

![PER] = 0: ![CUM] = 0: ![RANK] = 0: ![ADDFISH] = ADDFISH: ![ADDVALUES] = ADDVALUE

FFC = "C" + Format(CURM, "00"): .Fields(FFC) = catch
FFE = "E" + Format(CURM, "00"): .Fields(FFE) = eff
FFV = "V" + Format(CURM, "00"): .Fields(FFV) = Value
FFU = "U" + Format(CURM, "00")
FFP = "P" + Format(CURM, "00")
FFW = "W" + Format(CURM, "00")
FFN = "F" + Format(CURM, "00"): .Fields(FFN) = fish

![ADDFISH] = "NO": ![ADDVALUES] = "YES"

If .Fields(FFE) <> 0 Then .Fields(FFU) = .Fields(FFC) / .Fields(FFE)
If .Fields(FFC) <> 0 Then .Fields(FFP) = .Fields(FFV) / .Fields(FFC)
If .Fields(FFN) <> 0 Then .Fields(FFW) = .Fields(FFC) / .Fields(FFN)

.Update

GoTo READ_NEXT

UPDATE_REC:

.Edit
          
![PER] = 0: ![CUM] = 0: ![RANK] = 0: ![ADDFISH] = ADDFISH: ![ADDVALUES] = ADDVALUE

FFC = "C" + Format(CURM, "00"): .Fields(FFC) = catch
FFE = "E" + Format(CURM, "00"): .Fields(FFE) = eff
FFV = "V" + Format(CURM, "00"): .Fields(FFV) = Value
FFU = "U" + Format(CURM, "00")
FFP = "P" + Format(CURM, "00")
FFW = "W" + Format(CURM, "00")
FFN = "F" + Format(CURM, "00"): .Fields(FFN) = fish

If .Fields(FFE) <> 0 Then .Fields(FFU) = .Fields(FFC) / .Fields(FFE)
If .Fields(FFC) <> 0 Then .Fields(FFP) = .Fields(FFV) / .Fields(FFC)
If .Fields(FFN) <> 0 Then .Fields(FFW) = .Fields(FFC) / .Fields(FFN)
       
![ADDFISH] = "NO": ![ADDVALUES] = "YES"
       
.Update

READ_NEXT:

Loop

Close #1

End With

prm_record.Close
prm_database.Close

End Sub
Private Sub CREATE_LOOP()

Dim I, ix, J, jx, fnm, DBN, K

ix = Format(CURY, "0000")

DBN = APPROOT + "\ARTS\DATA\A" + ix + ".MDB"
FileCopy APPROOT + "\ARTS\STRUS\ARTS.MDB", DBN

DBN = APPROOT + "\ARTS\DATA\S" + ix + ".MDB"
FileCopy APPROOT + "\ARTS\STRUS\ARTS.MDB", DBN

For J = 1 To 12

pgbCR.Value = J + 1

CURM = J

For K = 1 To 10000

If MNC(K) <> "Y" Then GoTo NEXT_K

CURMN = K: CURMJ = ASSO(K)

Call UPDATE_TOT

NEXT_K:

Next K

Next J

Dim XXX

fnm = APPROOT + "\ARTS\TABLES\Y" + Format(CURY, "0000") + "_MONTHS.TXT"

Open fnm For Output As #1

XXX = ""

For J = 1 To 12
XXX = XXX + VALID_MONTHS(J)
Next J

Print #1, Format(CURY, "0000") + " " + XXX

Close #1

For J = 1 To 12

pgbCR.Value = 13 + J

CURM = J

For K = 1 To 10000

If MNC(K) <> "Y" Then GoTo NEXT_K2
CURMN = K: CURMJ = ASSO(K)

Call UPDATE_SPECIES

NEXT_K2:

Next K

Next J

End Sub
Private Sub UPDATE_SPECIES()

Dim I, fnm, ix, XKEY, XXX, FFF, FFC, FFE, FFU, FFP, FFV, FFW, FFN

fnm = APPROOT + "\ARTBAS\RESULTS\Y" + Format(CURY, "0000") + _
      "M" + Format(CURM, "00") + "_MN" + Format(CURMN, "0000") + "_ESTIM.TXT"

If Dir(fnm) = "" Then Exit Sub

Dim estdes, popn, smpn, BAC_ACCUR, FRNO, actno, cal, eact, _
          esmp, esites, edays, bac, bac_cvs, bac_cvsp, _
          bac_cvt, bac_cvtp, bac_cv, bac_low, bac_upper, eff, _
          eff_low, eff_upper, nland, CPUE_ACCUR, LPOP, ltot, _
          lsmpv, lsmpf, leff, cpue, lsites, ldays, _
          cpue_cvs, cpue_cvsp, cpue_cvt, cpue_cvtp, cpue_cv, cpue_low, _
          cpue_upper, catch, catch_low, catch_upper, catch_cv, Value, _
          price, fish, kgfish

Dim DBN, CREC

DBN = APPROOT + "\ARTS\DATA\S" + Format(CURY, "0000") + ".MDB"

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(DBN)
Set prm_record = prm_database.OpenRecordset("ASITAB")

With prm_record

.Index = "primarykey"

Dim TYPEREC

Open fnm For Input As #1

Do Until EOF(1)

Input #1, XKEY

If Mid(XKEY, 14, 4) = "0000" And Mid(XKEY, 8, 4) <> "0000" Then
   
Input #1, estdes, popn, smpn, BAC_ACCUR, FRNO, actno, cal, eact, _
          esmp, esites, edays, bac, bac_cvs, bac_cvsp, _
          bac_cvt, bac_cvtp, bac_cv, bac_low, bac_upper, eff, _
          eff_low, eff_upper, nland, CPUE_ACCUR, LPOP, ltot, _
          lsmpv, lsmpf, leff, cpue, lsites, ldays, _
          cpue_cvs, cpue_cvsp, cpue_cvt, cpue_cvtp, cpue_cv, cpue_low, _
          cpue_upper, catch, catch_low, catch_upper, catch_cv, Value, _
          price, fish, kgfish

   GoTo READ_NEXT
   End If

If Mid(XKEY, 14, 4) <> "0000" And Mid(XKEY, 8, 4) = "0000" Then
   Input #1, estdes, eff, cpue, catch, Value, price, kgfish, fish, FRNO
   
   GoTo READ_NEXT
   End If

If Mid(XKEY, 14, 4) = "0000" And Mid(XKEY, 8, 4) = "0000" Then
   Input #1, estdes, eff, cpue, catch, Value, price, kgfish, fish, FRNO

   GoTo READ_NEXT
   End If

Input #1, estdes, eff, cpue, catch, Value, price, kgfish, fish, FRNO
               
If catch + eff = 0 Then GoTo READ_NEXT
         
XKEY = "J" + Format(CURMJ, "0000") + "+" + XKEY
          
.Seek "=", XKEY

If .NoMatch = True Then GoTo ADD_NEW
If .NoMatch = False Then GoTo UPDATE_REC

ADD_NEW:

.AddNew

![akey] = XKEY

For I = 1 To 13

ix = Format(I, "00")

FFC = "C" + ix: .Fields(FFC) = 0
FFE = "E" + ix: .Fields(FFE) = 0
FFU = "U" + ix: .Fields(FFU) = 0
FFP = "P" + ix: .Fields(FFP) = 0
FFV = "V" + ix: .Fields(FFV) = 0
FFW = "W" + ix: .Fields(FFW) = 0
FFN = "F" + ix: .Fields(FFN) = 0
          
Next I

![PER] = 0: ![CUM] = 0: ![RANK] = 0: ![ADDFISH] = ADDFISH: ![ADDVALUES] = ADDVALUE

FFC = "C" + Format(CURM, "00"): .Fields(FFC) = catch
FFE = "E" + Format(CURM, "00"): .Fields(FFE) = eff
FFV = "V" + Format(CURM, "00"): .Fields(FFV) = Value
FFU = "U" + Format(CURM, "00")
FFP = "P" + Format(CURM, "00")
FFW = "W" + Format(CURM, "00")
FFN = "F" + Format(CURM, "00"): .Fields(FFN) = fish

If .Fields(FFE) <> 0 Then .Fields(FFU) = .Fields(FFC) / .Fields(FFE)
If .Fields(FFC) <> 0 Then .Fields(FFP) = .Fields(FFV) / .Fields(FFC)
If .Fields(FFN) <> 0 Then .Fields(FFW) = .Fields(FFC) / .Fields(FFN)
                 
![ADDFISH] = "NO": ![ADDVALUES] = "YES"
                 
.Update

GoTo READ_NEXT

UPDATE_REC:

.Edit
          
![PER] = 0: ![CUM] = 0: ![RANK] = 0: ![ADDFISH] = ADDFISH: ![ADDVALUES] = ADDVALUE

FFC = "C" + Format(CURM, "00"): .Fields(FFC) = catch
FFE = "E" + Format(CURM, "00"): .Fields(FFE) = eff
FFV = "V" + Format(CURM, "00"): .Fields(FFV) = Value
FFU = "U" + Format(CURM, "00")
FFP = "P" + Format(CURM, "00")
FFW = "W" + Format(CURM, "00")
FFN = "F" + Format(CURM, "00"): .Fields(FFN) = fish

If .Fields(FFE) <> 0 Then .Fields(FFU) = .Fields(FFC) / .Fields(FFE)
If .Fields(FFC) <> 0 Then .Fields(FFP) = .Fields(FFV) / .Fields(FFC)
If .Fields(FFN) <> 0 Then .Fields(FFW) = .Fields(FFC) / .Fields(FFN)

![ADDFISH] = "NO": ![ADDVALUES] = "YES"

.Update

READ_NEXT:

Loop

Close #1

End With

prm_record.Close
prm_database.Close

End Sub
Private Sub lstYEAR_Click()

CURY = YTAB2(lstYEAR.ListIndex + 1)

If OLDY = CURY Then
   OLDY = 9999
   cmdENTER.Enabled = False
   lstDISP.Clear
   lstYEAR.Clear
   Call CHECK_ARTBASIC
   Call CHECK_ARTSER
   Exit Sub
   End If
   
lstDISP.Visible = False

OLDY = CURY
   
Label3.Visible = False
Label4.Visible = False
Label5.Visible = False
pgbCR.Visible = False
cmdENTER.Enabled = True

optEXP.Visible = False
optIMP.Visible = False

End Sub
Private Sub INSPECT_DB()

NSOUS = 0

Dim TW(), TF()
ReDim TW(1 To 13), TF(1 To 13)

Dim TC(), TE(), TU(), TP(), TV(), I, ix, FFC, FFE, FFP, FFU, FFV, RR, cc, PP
Dim FFW, FFN

ReDim TC(1 To 13), TE(1 To 13), TU(1 To 13), TP(1 To 13), TV(1 To 13)

Dim fnm, XKEY

fnm = APPROOT + "\ARTS\DATA\S" + Format(CURY, "0000") + ".TXT"

If Dir(fnm) = "" Then Exit Sub

Open fnm For Input As #2

Do Until EOF(2)

Input #2, XKEY, TC(1), TC(2), TC(3), TC(4), TC(5), TC(6), TC(7), TC(8), TC(9), TC(10), TC(11), TC(12), TC(13), _
                   TE(1), TE(2), TE(3), TE(4), TE(5), TE(6), TE(7), TE(8), TE(9), TE(10), TE(11), TE(12), TE(13), _
                   TU(1), TU(2), TU(3), TU(4), TU(5), TU(6), TU(7), TU(8), TU(9), TU(10), TU(11), TU(12), TU(13), _
                   TP(1), TP(2), TP(3), TP(4), TP(5), TP(6), TP(7), TP(8), TP(9), TP(10), TP(11), TP(12), TP(13), _
                   TV(1), TV(2), TV(3), TV(4), TV(5), TV(6), TV(7), TV(8), TV(9), TV(10), TV(11), TV(12), TV(13), _
                   TW(1), TW(2), TW(3), TW(4), TW(5), TW(6), TW(7), TW(8), TW(9), TW(10), TW(11), TW(12), TW(13), _
                   TF(1), TF(2), TF(3), TF(4), TF(5), TF(6), TF(7), TF(8), TF(9), TF(10), TF(11), TF(12), TF(13), _
                   RR, cc, PP, ADDFISH, ADDVALUE

NSOUS = NSOUS + 1

I = Val(Mid(XKEY, 2, 4)): MJC(I) = "X"
I = Val(Mid(XKEY, 8, 4)): MNC(I) = "X"
I = Val(Mid(XKEY, 14, 4)): BGC(I) = "X"
I = Val(Mid(XKEY, 20, 4)): SPEC(I) = "X"

Loop

Close #2

End Sub
Private Sub FINISH_MAJOR()

Dim fnm, wnm, XXX, I

fnm = APPROOT + "\ARTS\TABLES\Y" + Format(CURY, "0000") + "_MAJOR.TXT"
wnm = APPROOT + "\ARTS\TABLES\WORK.TXT"

If Dir(fnm) = "" Then Exit Sub

Open fnm For Input As #1
Open wnm For Output As #2

Do Until EOF(1)

Line Input #1, XXX

I = Val(Left(XXX, 4))

If MJC(I) <> "X" Then GoTo NOT_WRITE

Print #2, XXX

NOT_WRITE:

Loop

Close #1
Close #2

FileCopy wnm, fnm

Kill wnm

End Sub
Private Sub FINISH_MINOR()

Dim fnm, wnm, XXX, I

fnm = APPROOT + "\ARTS\TABLES\Y" + Format(CURY, "0000") + "_MINOR.TXT"
wnm = APPROOT + "\ARTS\TABLES\WORK.TXT"

If Dir(fnm) = "" Then Exit Sub

Open fnm For Input As #1
Open wnm For Output As #2

Do Until EOF(1)

Line Input #1, XXX

I = Val(Left(XXX, 4))

If MNC(I) <> "X" Then GoTo NOT_WRITE

Print #2, XXX

NOT_WRITE:

Loop

Close #1
Close #2

FileCopy wnm, fnm

Kill wnm

End Sub
Private Sub FINISH_BG()

Dim fnm, wnm, XXX, I

fnm = APPROOT + "\ARTS\TABLES\Y" + Format(CURY, "0000") + "_BG.TXT"
wnm = APPROOT + "\ARTS\TABLES\WORK.TXT"

If Dir(fnm) = "" Then Exit Sub

Open fnm For Input As #1
Open wnm For Output As #2

Do Until EOF(1)

Line Input #1, XXX

I = Val(Left(XXX, 4))

If BGC(I) <> "X" Then GoTo NOT_WRITE

Print #2, XXX

NOT_WRITE:

Loop

Close #1
Close #2

FileCopy wnm, fnm

Kill wnm

End Sub
Private Sub FINISH_SPECIES()

Dim fnm, wnm, XXX, I

fnm = APPROOT + "\ARTS\TABLES\Y" + Format(CURY, "0000") + "_SPECIES.TXT"
wnm = APPROOT + "\ARTS\TABLES\WORK.TXT"

If Dir(fnm) = "" Then Exit Sub

Open fnm For Input As #1
Open wnm For Output As #2

Do Until EOF(1)

Line Input #1, XXX

I = Val(Left(XXX, 4))

If SPEC(I) <> "X" Then GoTo NOT_WRITE

Print #2, XXX

NOT_WRITE:

Loop

Close #1
Close #2

FileCopy wnm, fnm

Kill wnm

End Sub
Private Sub CREATE_CURRENT_TOT()

Dim fnm, wnm, XXX, NN

Dim TW(), TF()
ReDim TW(1 To 13), TF(1 To 13)

Dim TC(), TE(), TU(), TP(), TV(), I, ix, FFC, FFE, FFP, FFU, FFV, FFW, FFN, RR, cc, PP

ReDim TC(1 To 13), TE(1 To 13), TU(1 To 13), TP(1 To 13), TV(1 To 13)

Dim DBN, XKEY

DBN = APPROOT + "\ARTS\WORK\CT" + Format(CURY, "0000") + ".MDB"

FileCopy APPROOT + "\ARTS\STRUS\ARTS.MDB", DBN

fnm = APPROOT + "\ARTS\DATA\A" + Format(CURY, "0000") + ".TXT"

Open fnm For Input As #2

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(DBN)
Set prm_record = prm_database.OpenRecordset("ASITAB")

With prm_record

.Index = "primarykey"

Do Until EOF(2)

Input #2, XKEY, TC(1), TC(2), TC(3), TC(4), TC(5), TC(6), TC(7), TC(8), TC(9), TC(10), TC(11), TC(12), TC(13), _
                   TE(1), TE(2), TE(3), TE(4), TE(5), TE(6), TE(7), TE(8), TE(9), TE(10), TE(11), TE(12), TE(13), _
                   TU(1), TU(2), TU(3), TU(4), TU(5), TU(6), TU(7), TU(8), TU(9), TU(10), TU(11), TU(12), TU(13), _
                   TP(1), TP(2), TP(3), TP(4), TP(5), TP(6), TP(7), TP(8), TP(9), TP(10), TP(11), TP(12), TP(13), _
                   TV(1), TV(2), TV(3), TV(4), TV(5), TV(6), TV(7), TV(8), TV(9), TV(10), TV(11), TV(12), TV(13), _
                   TW(1), TW(2), TW(3), TW(4), TW(5), TW(6), TW(7), TW(8), TW(9), TW(10), TW(11), TW(12), TW(13), _
                   TF(1), TF(2), TF(3), TF(4), TF(5), TF(6), TF(7), TF(8), TF(9), TF(10), TF(11), TF(12), TF(13), _
                   RR, cc, PP, ADDFISH, ADDVALUE

.AddNew

![akey] = XKEY

For I = 1 To 13

ix = Format(I, "00")

FFC = "C" + ix:  .Fields(FFC) = TC(I)
FFE = "E" + ix:  .Fields(FFE) = TE(I)
FFU = "U" + ix:  .Fields(FFU) = TU(I)
FFP = "P" + ix:  .Fields(FFP) = TP(I)
FFV = "V" + ix:  .Fields(FFV) = TV(I)
FFW = "W" + ix:  .Fields(FFW) = TW(I)
FFN = "F" + ix:  .Fields(FFN) = TF(I)

Next I

![RANK] = RR: ![PER] = PP: ![CUM] = cc: ![ADDFISH] = ADDFISH: ![ADDVALUES] = ADDVALUE

.Update

Loop

End With

Close #2

prm_record.Close
prm_database.Close
                   
End Sub
Private Sub CREATE_CURRENT_SPECIES()

Dim fnm, wnm, XXX, NN

Dim TW(), TF()
ReDim TW(1 To 13), TF(1 To 13)

Dim TC(), TE(), TU(), TP(), TV(), I, ix, FFC, FFE, FFP, FFU, FFV, RR, cc, PP
Dim FFW, FFN

ReDim TC(1 To 13), TE(1 To 13), TU(1 To 13), TP(1 To 13), TV(1 To 13)

Dim DBN, XKEY

DBN = APPROOT + "\ARTS\WORK\CS" + Format(CURY, "0000") + ".MDB"

FileCopy APPROOT + "\ARTS\STRUS\ARTS.MDB", DBN

fnm = APPROOT + "\ARTS\DATA\S" + Format(CURY, "0000") + ".TXT"

Open fnm For Input As #2

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(DBN)
Set prm_record = prm_database.OpenRecordset("ASITAB")

With prm_record

.Index = "primarykey"

Do Until EOF(2)

Input #2, XKEY, TC(1), TC(2), TC(3), TC(4), TC(5), TC(6), TC(7), TC(8), TC(9), TC(10), TC(11), TC(12), TC(13), _
                   TE(1), TE(2), TE(3), TE(4), TE(5), TE(6), TE(7), TE(8), TE(9), TE(10), TE(11), TE(12), TE(13), _
                   TU(1), TU(2), TU(3), TU(4), TU(5), TU(6), TU(7), TU(8), TU(9), TU(10), TU(11), TU(12), TU(13), _
                   TP(1), TP(2), TP(3), TP(4), TP(5), TP(6), TP(7), TP(8), TP(9), TP(10), TP(11), TP(12), TP(13), _
                   TV(1), TV(2), TV(3), TV(4), TV(5), TV(6), TV(7), TV(8), TV(9), TV(10), TV(11), TV(12), TV(13), _
                   TW(1), TW(2), TW(3), TW(4), TW(5), TW(6), TW(7), TW(8), TW(9), TW(10), TW(11), TW(12), TW(13), _
                   TF(1), TF(2), TF(3), TF(4), TF(5), TF(6), TF(7), TF(8), TF(9), TF(10), TF(11), TF(12), TF(13), _
                   RR, cc, PP, ADDFISH, ADDVALUE
                  
.AddNew

![akey] = XKEY

For I = 1 To 13

ix = Format(I, "00")

FFC = "C" + ix:  .Fields(FFC) = TC(I)
FFE = "E" + ix:  .Fields(FFE) = TE(I)
FFU = "U" + ix:  .Fields(FFU) = TU(I)
FFP = "P" + ix:  .Fields(FFP) = TP(I)
FFV = "V" + ix:  .Fields(FFV) = TV(I)
FFW = "W" + ix:  .Fields(FFW) = TW(I)
FFN = "F" + ix:  .Fields(FFN) = TF(I)

Next I

![RANK] = RR: ![PER] = PP: ![CUM] = cc: ![ADDFISH] = ADDFISH: ![ADDVALUES] = ADDVALUE

.Update

Loop

End With

Close #2

prm_record.Close
prm_database.Close

End Sub
Private Sub optEXP_Click()

On Error GoTo EXP_ERR

Dim FNMBAT, XXX, fnm1, fnm2, resp

FNMBAT = APPROOT + "\ARTS\WORK\EXPORT.BAT"

Open FNMBAT For Output As #1

Print #1, "ECHO OFF"

fnm1 = APPROOT + "\ARTS\DATA\*.*"
fnm2 = APPROOT + "\ARTS\ARTSER_DATA_AND_TABLES\DATA"

Print #1, " COPY " + fnm1 + " " + fnm2

fnm1 = APPROOT + "\ARTS\TABLES\*.*"
fnm2 = APPROOT + "\ARTS\ARTSER_DATA_AND_TABLES\TABLES"

Print #1, " COPY " + fnm1 + " " + fnm2

Close #1

RUN_CODE = Shell(FNMBAT, 4)

Call lblQUIT_Click

resp = MsgBox(msgtab(122), vbOKOnly, "  ")

Exit Sub

EXP_ERR:

Call lblQUIT_Click

resp = MsgBox(msgtab(122), vbOKOnly, "  ")

Exit Sub

End Sub
Private Sub TRANSFER_FROM_TRANSFER()

Dim fnm, fnm1, fnm2

fnm = APPROOT + "\ARTBAS\TRANSFER\Y" + "*.*"

If Dir(fnm) = "" Then Exit Sub

fnm = Dir(fnm)
fnm1 = APPROOT + "\ARTBAS\TRANSFER\" + fnm
fnm2 = APPROOT + "\ARTBAS\RESULTS\" + fnm

FileCopy fnm1, fnm2

fnm = "?"

Do Until fnm = ""

fnm = Dir

fnm1 = APPROOT + "\ARTBAS\TRANSFER\" + fnm
fnm2 = APPROOT + "\ARTBAS\RESULTS\" + fnm

If fnm <> "" Then FileCopy fnm1, fnm2

Loop

End Sub
Private Sub optIMP_Click()

If Dir(APPROOT + "\ARTS\ARTSER_DATA_AND_TABLES\DATA\*.*") = "" Then GoTo IMP_ERR
If Dir(APPROOT + "\ARTS\ARTSER_DATA_AND_TABLES\TABLES\*.*") = "" Then GoTo IMP_ERR

Dim FNMBAT, XXX, fnm1, fnm2, resp

FNMBAT = APPROOT + "\ARTS\WORK\IMPORT.BAT"

Open FNMBAT For Output As #1

Print #1, "ECHO OFF"

fnm1 = APPROOT + "\ARTS\ARTSER_DATA_AND_TABLES\DATA\*.*"
fnm2 = APPROOT + "\ARTS\DATA"

Print #1, " COPY " + fnm1 + " " + fnm2

fnm1 = APPROOT + "\ARTS\ARTSER_DATA_AND_TABLES\TABLES\*.*"
fnm2 = APPROOT + "\ARTS\TABLES"

Print #1, " COPY " + fnm1 + " " + fnm2

Close #1

RUN_CODE = Shell(FNMBAT, 4)

Call lblQUIT_Click

resp = MsgBox(msgtab(122), vbOKOnly, "  ")

' TNJ - This is really strange but I am leaving the code in for now
Unload frmARTS00
Load frmARTS00
frmARTS00.Show

Exit Sub

IMP_ERR:

Call lblQUIT_Click

resp = MsgBox(msgtab(125), vbOKOnly, "  ")

Exit Sub

End Sub
Private Sub LOCATE_EXCEL()

' VVVVVVVV Commented out by TNJ 24/1/08 VVVVVVVV

'
' New revisions - 15/6/09 by TNJ
' Since the user requires Local Admin rights with my 'read the registry' method,
' I reinstated the Constantine code

Dim PATH_EXCEL_GENERIC, PATH_EXCEL, EXCEL_FOUND, I, IX1, IX2, PTH1, PTH2

PATH_EXCEL_GENERIC = "C:\PROGRAM FILES\MICROSOFT OFFICE\OFFICE"
PATH_EXCEL = PATH_EXCEL_GENERIC

EXCEL_FOUND = "NO"

If Dir(PATH_EXCEL) <> "" Then
   EXCEL_FOUND = "OK"
   GoTo EXCEL_OK
   End If
   
For I = 1 To 50

    IX1 = Format(I, "#0"): IX2 = Format(I, "00")

    PTH1 = PATH_EXCEL_GENERIC + IX1 + "\EXCEL.EXE": PTH2 = PATH_EXCEL_GENERIC + IX2 + "\EXCEL.EXE"

    If Dir(PTH1) <> "" Then
       EXCEL_FOUND = "OK"
       PATH_EXCEL = PATH_EXCEL_GENERIC + IX1
       GoTo EXCEL_OK
       End If
   
    If Dir(PTH2) <> "" Then
       EXCEL_FOUND = "OK"
       PATH_EXCEL = PATH_EXCEL_GENERIC + IX2
       GoTo EXCEL_OK
       End If
   
Next I

If EXCEL_FOUND = "NO" Then
    MsgBox ("!!!!!!!!!!!!!! Important - The Excel application was not located by ArtSer !!!!!!!!!!!!!")
    MsgBox ("Batch file 'GENERAL.BAT' in the EXCEL_REPORTS subdirectory must be edited manually to access the Excel application")
    MsgBox ("Note: The path for the Excel application can NOT have spaces.  You must use 'short' format")
    End If

EXCEL_OK:

Open APPROOT + "\ARTS\EXCEL_REPORTS\GENERAL.BAT" For Output As #1

' We write the Excel call with passing the parameter _AR: that gives the
' AppRoot (Application Root) to tell where we are executing the application
Print #1, "ECHO OFF"
Print #1, GetShortName(LTrim(PATH_EXCEL)) & "\EXCEL.EXE /_AR:" _
    & APPROOT & " " _
    & APPROOT & "\ARTS\EXCEL_REPORTS\ARTFISH_WORK.XLS"

Close #1

' VVVVVVVV Added by TNJ 24/1/08 VVVVVVVV
' Determines the Excel Application location by interrogating the registry
' Note!!!  This query of the registry does NOT work if the user doesn't have
'    Local Administration privileges on his PC

' So this was a great idea but reality slaps us in the face and we backtrack
' to the previous way that it was done by Constantine

' Dim hKey As Long
' Dim RetVal As Long
' Dim sProgId As String
' Dim sCLSID As String
' Dim sPath As String

'    sProgId = "Excel.Application"

   'First, get the clsid from the progid from the registry key:
   'HKEY_LOCAL_MACHINE\Software\Classes\<PROGID>\CLSID
'    RetVal = RegOpenKeyEx(HKEY_LOCAL_MACHINE, "Software\Classes\" & _
'       sProgId & "\CLSID", 0&, KEY_ALL_ACCESS, hKey)
'    If RetVal = 0 Then
'       Dim N As Long
'       RetVal = RegQueryValueEx(hKey, "", 0&, REG_SZ, "", N)
'       sCLSID = Space(N)
'       RetVal = RegQueryValueEx(hKey, "", 0&, REG_SZ, sCLSID, N)
'       sCLSID = Left(sCLSID, N - 1)  'drop null-terminator
'       RegCloseKey hKey
'    End If
   
   'Now that we have the CLSID, locate the server path at
   'HKEY_LOCAL_MACHINE\Software\Classes\CLSID\
   '     {xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxx}\LocalServer32

'     RetVal = RegOpenKeyEx(HKEY_LOCAL_MACHINE, _
'         "Software\Classes\CLSID\" & sCLSID & "\LocalServer32", 0&, _
'       KEY_ALL_ACCESS, hKey)
'    If RetVal = 0 Then
'       RetVal = RegQueryValueEx(hKey, "", 0&, REG_SZ, "", N)
'       sPath = Space(N)

'       RetVal = RegQueryValueEx(hKey, "", 0&, REG_SZ, sPath, N)
'       sPath = Left(sPath, N - 1)
      ' MsgBox sPath
'       RegCloseKey hKey
'    End If

   ' Now - only take first part of the path
'    N = InStr(sPath, "EXCEL.EXE")
'    sPath = Left(sPath, N - 2)

'    Open APPROOT + "\ARTS\EXCEL_REPORTS\GENERAL.BAT" For Output As #1

'    Print #1, "ECHO OFF"
'    Print #1, "CD\"
'    Print #1, "CD " & sPath
'    Print #1, "EXCEL.EXE C:\ARTFISH_WORK.XLS"

'    Close #1

' ^^^^^^^^ Added by TNJ 24/1/08 ^^^^^^^^

End Sub


