VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmARTB01 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7800
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   10995
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
   ScaleHeight     =   7800
   ScaleWidth      =   10995
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGO 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   4920
      Picture         =   "frmARTB01.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   1080
      Width           =   735
   End
   Begin VB.CommandButton cmdREDO 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   4920
      Picture         =   "frmARTB01.frx":010A
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   5400
      Width           =   735
   End
   Begin VB.OptionButton optDB 
      BackColor       =   &H008080FF&
      Caption         =   "DATABASE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6840
      TabIndex        =   29
      Top             =   4320
      Width           =   3855
   End
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
      Left            =   10080
      MousePointer    =   1  'Arrow
      Picture         =   "frmARTB01.frx":038C
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   5520
      Width           =   735
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   3855
      Left            =   6720
      TabIndex        =   17
      Top             =   1320
      Width           =   4095
      Begin VB.OptionButton optLEAVE 
         BackColor       =   &H00C0C000&
         Caption         =   "Option1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   3360
         Width           =   3855
      End
      Begin VB.OptionButton optEXPT 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Option1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   480
         Width           =   3375
      End
      Begin VB.OptionButton optEXPR 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Option1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   840
         Width           =   3375
      End
      Begin VB.OptionButton optIMPT 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Option1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1800
         Width           =   3375
      End
      Begin VB.OptionButton optIMPR 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Option1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   2160
         Width           =   3375
      End
      Begin VB.Label lblEXP 
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
         ForeColor       =   &H00808000&
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   120
         Width           =   3375
      End
      Begin VB.Label lblIMP 
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
         ForeColor       =   &H00808000&
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   1440
         Width           =   3375
      End
   End
   Begin VB.CommandButton cmdSERVICES 
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
      Left            =   10080
      MousePointer    =   1  'Arrow
      Picture         =   "frmARTB01.frx":25EE
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   480
      Width           =   735
   End
   Begin VB.CommandButton cmdKESTIM 
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
      Height          =   255
      Left            =   8400
      Picture         =   "frmARTB01.frx":43E0
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   7440
      Width           =   255
   End
   Begin VB.CommandButton cmdKACTIVE 
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
      Height          =   255
      Left            =   7560
      Picture         =   "frmARTB01.frx":44EA
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7440
      Width           =   255
   End
   Begin VB.CommandButton cmdKLAND 
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
      Height          =   255
      Left            =   6720
      Picture         =   "frmARTB01.frx":45F4
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7440
      Width           =   255
   End
   Begin VB.CommandButton cmdKEFFORT 
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
      Height          =   255
      Left            =   5880
      Picture         =   "frmARTB01.frx":46FE
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7440
      Width           =   255
   End
   Begin VB.CommandButton cmdKTABLES 
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
      Height          =   255
      Left            =   480
      Picture         =   "frmARTB01.frx":4808
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7440
      Width           =   255
   End
   Begin VB.CommandButton cmdLAND 
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
      Left            =   6480
      Picture         =   "frmARTB01.frx":4912
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6600
      Width           =   735
   End
   Begin VB.CommandButton cmdEFFORT 
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
      Left            =   5640
      Picture         =   "frmARTB01.frx":5414
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6600
      Width           =   735
   End
   Begin VB.CommandButton cmdESTIM 
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
      Left            =   8160
      MousePointer    =   1  'Arrow
      Picture         =   "frmARTB01.frx":5696
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6600
      Width           =   735
   End
   Begin VB.CommandButton cmdACT 
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
      Left            =   7320
      Picture         =   "frmARTB01.frx":5918
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6600
      Width           =   735
   End
   Begin VB.CommandButton cmdTABLES 
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
      Picture         =   "frmARTB01.frx":5B9A
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6600
      Width           =   735
   End
   Begin VB.CommandButton cmdREP 
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
      Left            =   9000
      MousePointer    =   1  'Arrow
      Picture         =   "frmARTB01.frx":5E1C
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6600
      Width           =   735
   End
   Begin VB.CommandButton cmdQUIT 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   10440
      MousePointer    =   1  'Arrow
      Picture         =   "frmARTB01.frx":609E
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6960
      Width           =   375
   End
   Begin VB.ListBox lstSTRUS 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   5100
      Left            =   1680
      MultiSelect     =   1  'Simple
      TabIndex        =   2
      Top             =   1080
      Width           =   3135
   End
   Begin VB.CommandButton cmdDEL 
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
      Left            =   2880
      MousePointer    =   1  'Arrow
      Picture         =   "frmARTB01.frx":6320
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6600
      Width           =   735
   End
   Begin VB.CommandButton cmdRETO00 
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
      Height          =   975
      Left            =   9840
      MousePointer    =   1  'Arrow
      Picture         =   "frmARTB01.frx":65A2
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6360
      Width           =   975
   End
   Begin ComctlLib.ProgressBar pgbFILES 
      Height          =   255
      Left            =   6720
      TabIndex        =   30
      Top             =   5160
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
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
      TabIndex        =   28
      Top             =   7560
      Width           =   255
   End
   Begin VB.Label lblADMIN 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "???"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   2160
      TabIndex        =   26
      Top             =   0
      Width           =   6375
   End
   Begin VB.Label lblTYPE 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   4560
      TabIndex        =   24
      Top             =   4080
      Width           =   3615
   End
   Begin VB.Label lblSTRUS 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "???"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   1680
      TabIndex        =   3
      Top             =   840
      Width           =   3135
   End
End
Attribute VB_Name = "frmARTB01"
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
Private LOAD_TABLES_FLAG, LANG_IND, ADMYESNO
Private Sub cmdACT_Click()

Call CHECK_BACKUP_COMPLETE

Call CHECK_RESULTS_EXIST

If PROCEED_FLAG = "N" Then Exit Sub

Load frmACTIVE
Unload frmARTB01
frmACTIVE.Show

End Sub
Private Sub CHECK_READYA()

Dim fnm, resp

On Error GoTo EXIT_SUB

READY_ERROR = "N"

fnm = "A:\*.*"

If Dir(fnm) <> "" Then Exit Sub
If Dir(fnm) = "" Then Exit Sub

EXIT_SUB:

READY_ERROR = "Y"

End Sub
Private Sub cmdDEL_Click()

If lstSTRUS.Visible = True Then Exit Sub

Dim response, XXX As String

response = MsgBox(msgtab(9), vbOKCancel + vbCritical, " ")
   
If response = 2 Then Exit Sub
   
Dim fnm, ccy, ccm

ccy = current_year: ccm = current_month

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(ccy, "0000") + "M" + Format(ccm, "00") + _
      "*.*"

If Dir(fnm) <> "" Then Kill fnm

ccy = current_year: ccm = current_month

fnm = APPROOT + "\ARTBAS\EFFORT\Y" + Format(ccy, "0000") + "M" + Format(ccm, "00") + _
      "*.*"

If Dir(fnm) <> "" Then Kill fnm

ccy = current_year: ccm = current_month

fnm = APPROOT + "\ARTBAS\LANDINGS\Y" + Format(ccy, "0000") + "M" + Format(ccm, "00") + _
      "*.*"

If Dir(fnm) <> "" Then Kill fnm

ccy = current_year: ccm = current_month

fnm = APPROOT + "\ARTBAS\RESULTS\Y" + Format(ccy, "0000") + "M" + Format(ccm, "00") + _
      "*.*"

If Dir(fnm) <> "" Then Kill fnm

cmdDEL.Visible = False
cmdREP.Enabled = False

'If ADMYESNO = "YES" And cmdKTABLES.Visible = True Then cmdSERVICES.Visible = True

Call CHECK_STATUS

End Sub
Private Sub cmdDUPL_Click()

If lstSTRUS.SelCount = 0 Then GoTo END_SUB

Dim resp

resp = MsgBox(msgtab(36), vbOKCancel, " ")

If resp = vbCancel Then

lstSTRUS.Clear
lstSTRUS.Refresh
Call NEW_PERIOD
GoTo END_SUB

End If
   
Call CHECK_BACKUP_COMPLETE

Dim LY, LM, LI, XXX, CY, cm

CY = current_year: cm = current_month

LI = lstSTRUS.ListIndex

XXX = LISTMY(LI + 1)

LY = Val(Left(XXX, 4)): LM = Val(Mid(XXX, 8, 4))

Dim fr1, fr2, to1, to2

fr1 = APPROOT + "\ARTBAS\TABLES\Y" + Format(LY, "0000") + "M" + Format(LM, "00") + "_"

to1 = APPROOT + "\ARTBAS\TABLES\Y" + Format(CY, "0000") + "M" + Format(cm, "00") + "_"

'-------------------------------------------------------

fr2 = fr1 + "MINOR.TXT": to2 = to1 + "MINOR.TXT"

If Dir(fr2) <> "" Then FileCopy fr2, to2

fr2 = fr1 + "MAJOR.TXT": to2 = to1 + "MAJOR.TXT"

If Dir(fr2) <> "" Then FileCopy fr2, to2

fr2 = fr1 + "ASSOMN.TXT": to2 = to1 + "ASSOMN.TXT"

If Dir(fr2) <> "" Then FileCopy fr2, to2

fr2 = fr1 + "SITES.TXT": to2 = to1 + "SITES.TXT"

If Dir(fr2) <> "" Then FileCopy fr2, to2

fr2 = fr1 + "ASSOSI.TXT": to2 = to1 + "ASSOSI.TXT"

If Dir(fr2) <> "" Then FileCopy fr2, to2

fr2 = fr1 + "BG.TXT": to2 = to1 + "BG.TXT"

If Dir(fr2) <> "" Then FileCopy fr2, to2

fr2 = fr1 + "FRAME.TXT": to2 = to1 + "FRAME.TXT"

If Dir(fr2) <> "" Then FileCopy fr2, to2

fr2 = fr1 + "SPECIES.TXT": to2 = to1 + "SPECIES.TXT"

If Dir(fr2) <> "" Then FileCopy fr2, to2

fr2 = fr1 + "UNITS.TXT": to2 = to1 + "UNITS.TXT"

If Dir(fr2) <> "" Then FileCopy fr2, to2

fr2 = fr1 + "ACTIVE.TXT": to2 = to1 + "ACTIVE.TXT"

If Dir(fr2) <> "" Then
   FileCopy fr2, to2
   to2 = to1 + "WACTIVE.TXT"
   Open to2 For Output As #1
   Close #1
   End If

resp = MsgBox(msgtab(222), vbOKOnly, " ")

Call cmdRETO00_Click

END_SUB:

End Sub
Private Sub cmdEFFORT_Click()

Call CHECK_BACKUP_COMPLETE

Call CHECK_RESULTS_EXIST

If PROCEED_FLAG = "N" Then Exit Sub

cmdEFFORT.MousePointer = 13
Load frmEFFORT
Unload frmARTB01
frmEFFORT.Show

End Sub

Private Sub cmdESTIM_Click()

Call CHECK_BACKUP_COMPLETE

Load frmESTIM
Unload frmARTB01
frmESTIM.Show

End Sub

Private Sub cmdGO_Click()

Dim YMTAB(), NYM, NSEL, NN, I

Open APPROOT + "\ARTBAS\CONTROL\SELYM.TXT" For Output As #1

NYM = 0

NSEL = lstSTRUS.SelCount

If NSEL = 0 Then
   Close #1
   Exit Sub
   End If

NN = lstSTRUS.ListCount

lstSTRUS.ListIndex = 0

For I = 0 To lstSTRUS.ListCount - 1

lstSTRUS.ListIndex = I

If lstSTRUS.Selected(I) = True Then

   Print #1, LISTMY(I + 1)
         
   End If
   
Next I

Call optLEAVE_Click

EXIT_SUB:

Close #1

If NSEL <> 0 Then

   Dim PROGRAM_NAME
   PROGRAM_NAME = APPROOT + "\ARTBAS\CREATEDB.EXE"
   RUN_CODE = Shell(PROGRAM_NAME, 4)

End If

Call cmdREDO_Click

Exit Sub

Dim resp

resp = MsgBox(msgtab(286), vbOKOnly)

Call cmdRETO00_Click

End Sub

Private Sub cmdGUIDE_Click()

HTYPE = "00": HFNM = "XXX"

If CTLADMIN = "YES" And lstSTRUS.Visible = False And cmdREP.Visible = False Then
   HTYPE = "20"
   End If

If CTLADMIN <> "YES" And lstSTRUS.Visible = False And cmdREP.Visible = False Then
   HTYPE = "20W"
   End If

If CTLADMIN = "YES" And lstSTRUS.Visible = True Then
   HTYPE = "21"
   End If

If CTLADMIN <> "YES" And lstSTRUS.Visible = True Then
   HTYPE = "21W"
   End If

If CTLADMIN = "YES" And lstSTRUS.Visible = False And cmdKTABLES.Visible = False _
   And cmdREP.Visible = True Then
   HTYPE = "22"
   End If

If CTLADMIN <> "YES" And lstSTRUS.Visible = False And cmdKTABLES.Visible = False _
   And cmdREP.Visible = True Then
   HTYPE = "22W"
   End If

If CTLADMIN = "YES" And lstSTRUS.Visible = False And cmdKTABLES.Visible = True Then
   If cmdESTIM.Enabled = False Then HTYPE = "23"
   If cmdESTIM.Enabled = True Then HTYPE = "24"
   End If

If CTLADMIN <> "YES" And lstSTRUS.Visible = False And cmdKTABLES.Visible = True Then
   If cmdESTIM.Enabled = False Then HTYPE = "23W"
   If cmdESTIM.Enabled = True Then HTYPE = "24W"
   End If

HFNM = APPROOT + "\ARTBAS\HELP\" + current_language + "HELP" + HTYPE + ".rtf"

If Dir(HFNM) = "" Then Exit Sub

frmARTB01.Enabled = False
Load frmGUIDE
frmGUIDE.Show

End Sub

Private Sub cmdLAND_Click()

Call CHECK_BACKUP_COMPLETE

Call CHECK_RESULTS_EXIST

If PROCEED_FLAG = "N" Then Exit Sub

cmdLAND.MousePointer = 13
Load frmLAND
Unload frmARTB01
frmLAND.Show

End Sub
Private Sub cmdREDO_Click()

lstSTRUS.Visible = False
lblSTRUS.Visible = False
cmdGO.Visible = False
cmdREDO.Visible = False

optDB.Enabled = True
Frame1.Enabled = True
optDB.Value = False
optDB.Refresh

cmdSERVICES.Enabled = True

Call Form_Load

End Sub

Private Sub lblADMIN_Click()

Dim fnm, WADM

fnm = APPROOT + "\ARTBAS\CONTROL\ADMIN.TXT"

If Dir(fnm) = "" Then Exit Sub

Open fnm For Input As #1

Input #1, WADM
Close #1

Open fnm For Output As #1

If Left(WADM, 1) = "Y" Then
   Print #1, "NO"
   Close #1
   End If

If Left(WADM, 1) = "N" Then
   Print #1, "YES"
   Close #1
   End If

Open fnm For Input As #1

Input #1, CTLADMIN
Close #1

Call cmdRETO00_Click

End Sub
Private Sub lstSTRUS_Click()

If cmdGO.Visible = True Then GoTo END_SUB

Call cmdDUPL_Click

END_SUB:

End Sub

Private Sub optLEAVE_Click()

Frame1.Visible = False
optDB.Visible = False
cmdREDO.Visible = False
cmdGO.Visible = False

pgbFILES.Visible = False
optDB.Value = False

Call Form_Load

pgbFILES.Visible = False

cmdEFFORT.Visible = True
cmdLAND.Visible = True
cmdACT.Visible = True
cmdESTIM.Visible = True
cmdREP.Visible = True
cmdTABLES.Visible = True
cmdTABLES.Enabled = True
cmdDEL.Visible = True
Frame1.Visible = False
optLEAVE.Value = False

cmdKTABLES.Enabled = True
cmdKEFFORT.Enabled = True
cmdKLAND.Enabled = True
cmdKACTIVE.Enabled = True
cmdKESTIM.Enabled = True

Call CHECK_STATUS

If ADMYESNO <> "YES" Then cmdTABLES.Enabled = False

If lstSTRUS.Visible = True Then cmdDEL.Visible = False

cmdSERVICES.Enabled = True

End Sub
Private Sub optDB_Click()

optDB.Enabled = False
optEXPT.Enabled = False
optEXPR.Enabled = False
optIMPT.Enabled = False
optIMPR.Enabled = False
optDB.Refresh

cmdREDO.Visible = True
cmdGO.Visible = True

lstSTRUS.Visible = True
lblSTRUS.Visible = True

Call NEW_PERIOD

Frame1.Enabled = True
optDB.Enabled = False
optDB.Value = False

EXIT_SUB:

End Sub

Private Sub cmdQUIT_Click()

Call CHECK_BACKUP_COMPLETE

cmdRETO00.MousePointer = 13
cmdQUIT.MousePointer = 13
Call write_parms

Call KILL_ARTBASIC_FOLDER

Unload frmESTIM

End

End Sub
Private Sub cmdREP_Click()

Call CHECK_BACKUP_COMPLETE

cmdREP.MousePointer = 13
frmARTB01.MousePointer = 13
Load frmREP
Unload frmARTB01
frmREP.Show
End Sub
Private Sub cmdRETO00_Click()

cmdRETO00.MousePointer = 13
frmARTB01.MousePointer = 13

Call write_parms
Call KILL_ARTBASIC_FOLDER
' Load frmARTB00

frmARTB00.Show

Unload frmARTB01

End Sub
Private Sub cmdSERVICES_Click()

If lstSTRUS.Visible = True Then Exit Sub

lblSTRUS.Visible = False
lstSTRUS.Visible = False

optIMPT.Enabled = True
optIMPR.Enabled = True
optEXPT.Enabled = True
optEXPR.Enabled = True
optDB.Enabled = True

If cmdGO.Visible = True Then cmdGO.Visible = False
If cmdREDO.Visible = True Then cmdREDO.Visible = False

lblSTRUS.Caption = msgtab(7)
' -------------- CHECK OPTIONS FOR TABLES (ADMINISTRATOR)-------------------

If ADMYESNO = "YES" Then

'EXPORTING TABLES

   cmdEFFORT.Enabled = False
   'cmdKEFFORT.Enabled = False
   cmdLAND.Enabled = False
   'cmdKLAND.Enabled = False
   cmdACT.Enabled = False
   'cmdKACTIVE.Enabled = False
   cmdESTIM.Enabled = False
   'cmdKESTIM.Enabled = False
   cmdREP.Enabled = False
   cmdTABLES.Enabled = False
   'cmdKTABLES.Enabled = False
   cmdDEL.Enabled = False
   
   If NS <> 0 Then cmdDEL.Visible = False
   
   lblSTRUS.Visible = False
   lstSTRUS.Visible = False
   optEXPR.Enabled = False
   optEXPT.Enabled = True
   optDB.Visible = True
   
   If cmdKTABLES.Visible = False Then optEXPT.Enabled = False
   
   optEXPR.Enabled = False
   
'  CHECK IMPORTING RESULTS
   
   optIMPT.Enabled = False
   
   If cmdKTABLES.Visible = False Then
      optIMPR.Enabled = False
      optDB.Enabled = False
      End If
      
   End If
   
   Frame1.Visible = True

   If lstSTRUS.Visible = True Then
   
   optIMPT.Enabled = False
   optIMPR.Enabled = False
   optEXPT.Enabled = False
   optEXPR.Enabled = False
   'optDB.Enabled = False
   
End If

'----------------------------------------------------------------------
   
   ' CHECK OPTIONS FOR TABLES (LOCAL USER)

If ADMYESNO <> "YES" Then

'IMPORTING TABLES

   cmdEFFORT.Enabled = False
   cmdKEFFORT.Enabled = False
   cmdLAND.Enabled = False
   cmdKLAND.Enabled = False
   cmdACT.Enabled = False
   cmdKACTIVE.Enabled = False
   cmdESTIM.Enabled = False
   cmdKESTIM.Enabled = False
   cmdREP.Enabled = False
   cmdTABLES.Enabled = False
   cmdKTABLES.Enabled = False
   cmdDEL.Enabled = False
   
   lblSTRUS.Visible = False
   lstSTRUS.Visible = False
   
   optEXPR.Enabled = True
   If cmdKESTIM.Visible = False Then optEXPR.Enabled = False
   
   optEXPT.Enabled = False
   
   optDB.Enabled = True
   If cmdKESTIM.Visible = False Then optDB.Enabled = False
   
   optEXPT.Enabled = False
   
'  CHECK IMPORTING RESULTS
   
   optIMPR.Enabled = False
    
   Frame1.Visible = True
   
   optIMPT.Enabled = True
   
   ' If Dir("C:\ARTBASIC_TABLES\*.*") <> "" Then optIMPT.Enabled = True
   If Dir(APPROOT + "\ARTBAS\TABLES\*.*") <> "" Then optIMPT.Enabled = True
   
   optIMPR.Enabled = False
   optEXPT.Enabled = False
   optEXPR.Enabled = False
   
   If cmdKESTIM.Visible = True Then optEXPR.Enabled = True
   
   optDB.Enabled = True
   If cmdKESTIM.Visible = True Then optDB.Enabled = True

End If
         
Frame1.Visible = True
optDB.Visible = True

If cmdKESTIM.Visible = False Then optEXPR.Enabled = False

EXIT_SUB:

Exit Sub

If NS <> 0 Then
   lblSTRUS.Visible = True
   lstSTRUS.Visible = True
   cmdDEL.Visible = False
   End If

End Sub
Private Sub cmdTABLES_Click()

cmdTABLES.MousePointer = 13
frmARTB01.MousePointer = 13

Call CHECK_BACKUP_COMPLETE

Call CHECK_RESULTS_EXIST

frmARTB01.MousePointer = 1

If PROCEED_FLAG = "N" Then Exit Sub

Load frmTABLES
Unload frmARTB01
frmTABLES.Show

End Sub
Private Sub Form_Load()

pgbFILES.Visible = False
pgbFILES.Min = 0: pgbFILES.Max = 13

Frame1.Visible = False
optDB.Visible = False
cmdGO.Visible = False
cmdREDO.Visible = False

Set Picture = LoadPicture(APPROOT + "\ARTBAS\PICS_RUNTIME\SCREEN_02.JPG")

NSTRUS = 0

If CTLADMIN = "NO" Then cmdTABLES.Enabled = False

Call CALC_CALDAYS

Dim I As Integer

ReDim monthtab(1 To 12)

cmdQUIT.ToolTipText = msgtab(3)

For I = 1 To 12
monthtab(I) = msgtab(I + 17)
Next I

frmARTB01.Caption = msgtab(40) + " - " + monthtab(current_month) + " " + _
                    Format(current_year, "0000")

lblADMIN.Caption = msgtab(241)

If CTLADMIN <> "YES" Then
lblADMIN.Caption = msgtab(242)
cmdTABLES.Enabled = False
cmdKTABLES.Enabled = False
End If

cmdGUIDE.ToolTipText = msgtab(243)
cmdDEL.ToolTipText = msgtab(9)
cmdREP.ToolTipText = msgtab(5)
cmdTABLES.ToolTipText = msgtab(8)
cmdRETO00.ToolTipText = msgtab(13)
cmdREDO.ToolTipText = msgtab(283)
cmdGO.ToolTipText = msgtab(36)

lblSTRUS.Caption = msgtab(285)

cmdEFFORT.ToolTipText = msgtab(33)
cmdLAND.ToolTipText = msgtab(34)
cmdACT.ToolTipText = msgtab(35)
cmdESTIM.ToolTipText = msgtab(37)
cmdREP.ToolTipText = msgtab(38)
cmdSERVICES.ToolTipText = msgtab(70)

lblEXP.Caption = msgtab(213)
optEXPT.Caption = msgtab(214)
optEXPR.Caption = msgtab(215)

lblIMP.Caption = msgtab(216)
optIMPT.Caption = msgtab(214)
optIMPR.Caption = msgtab(215)

optDB.Caption = msgtab(276)

optLEAVE.Caption = msgtab(231)

Dim dbn, resp

dbn = APPROOT + "\ARTBAS\STRUS\ALLDATA.MDB"

If Dir(dbn) = "" Then optDB.Visible = False

Refresh

Call CHECK_STATUS

Call CHECK_STRUCTURES

Open APPROOT + "\ARTBAS\Control\ADMIN.TXT" For Input As #2
Input #2, ADMYESNO
Close #2

NSTRUS = NS: ADMYESNO = RTrim(ADMYESNO)

If ADMYESNO <> "YES" Then cmdTABLES.Enabled = False
If ADMYESNO <> "YES" Then
cmdKTABLES.Enabled = False
cmdTABLES.Enabled = False
End If

cmdSERVICES.Visible = True
cmdSERVICES.Enabled = True


If ADMYESNO = "YES" Then
cmdTABLES.Visible = True
cmdTABLES.Visible = True
cmdTABLES.Enabled = True
cmdTABLES.Enabled = True
End If

cmdDEL.Enabled = True
cmdDEL.Visible = True

Refresh

If NS = 0 Then
   lblSTRUS.Visible = False
   lstSTRUS.Visible = False
   optIMPT.Enabled = True
   cmdDEL.Enabled = True
   cmdDEL.Visible = True
   cmdSERVICES.Enabled = True
   Refresh
   Exit Sub
   End If

'If NSTRUS <> 0 Then cmdSERVICES.Enabled = False

Refresh

End Sub
Private Sub LOAD_STRUCTURES()

Dim I, J

NS = 0

For I = 1990 To 2050

For J = 1 To 12

If CY(I - 1989, J) <> "X" Then GoTo CONT_LOOP

NS = NS + 1

ReDim Preserve LISTMY(1 To NS)

LISTMY(NS) = Format(I, "0000") + " - " + Format(J, "00") + " " + msgtab(17 + J)

CONT_LOOP:

Next J
Next I

If NS = 0 Then
   lblSTRUS.Visible = False
   lstSTRUS.Visible = False
   End If

End Sub
Private Sub DISPLAY_STRUS()

lstSTRUS.Clear

Dim I

For I = 1 To NS
lstSTRUS.AddItem LISTMY(I)
Next I

lstSTRUS.ListIndex = NS - 1

cmdSERVICES.Enabled = False

End Sub
Private Sub CHECK_STRUCTURES()

Dim I, J

I = current_year - 1989: J = current_month

If CY(I, J) <> "X" Then
   Call NEW_PERIOD
   Exit Sub
   End If
   
lstSTRUS.Visible = False
lblSTRUS.Visible = False

cmdTABLES.ToolTipText = msgtab(12)

End Sub
Private Sub NEW_PERIOD()

If ADMYESNO = "YES" And cmdKTABLES.Visible = True Then cmdSERVICES.Visible = True

cmdDEL.Enabled = False
cmdLAND.Enabled = False
cmdEFFORT.Enabled = False
cmdACT.Enabled = False
cmdESTIM.Enabled = False
cmdREP.Enabled = False

If ADMYESNO = "YES" And cmdKTABLES.Visible = True Then cmdSERVICES.Visible = True

Call LOAD_STRUCTURES

If NS = 0 Then
   lstSTRUS.Visible = False
   lblSTRUS.Visible = False
   lstSTRUS.Visible = False
   Exit Sub
   End If

NSTRUS = NS

Call DISPLAY_STRUS

End Sub
Private Sub CHECK_STATUS()

Dim KTABLES, KEFFORT, KLAND, KACT, KESTIM, KREP

KTABLES = "Y": KEFFORT = " ": KLAND = " ": KACT = " ": KESTIM = " ": KREP = " "

cmdKTABLES.Visible = True
cmdKEFFORT.Visible = False
cmdKLAND.Visible = False
cmdKACTIVE.Visible = False
cmdKESTIM.Visible = False

Dim ccy, ccm, FNM1, fnm2

ccy = current_year: ccm = current_month

FNM1 = APPROOT + "\ARTBAS\TABLES\Y" + Format(ccy, "0000") + "M" + _
       Format(ccm, "00") + "_"

'-----------------------

fnm2 = FNM1 + "MAJOR.TXT"

If Dir(fnm2) = "" Then
   cmdKTABLES.Visible = False
   KTABLES = " "
   End If
   
fnm2 = FNM1 + "MINOR.TXT"

If Dir(fnm2) = "" Then
   cmdKTABLES.Visible = False
   KTABLES = " "
   End If
   
fnm2 = FNM1 + "ASSOMN.TXT"

If Dir(fnm2) = "" Then
   cmdKTABLES.Visible = False
   KTABLES = " "
   End If
   
fnm2 = FNM1 + "SITES.TXT"

If Dir(fnm2) = "" Then
   cmdKTABLES.Visible = False
   KTABLES = " "
   End If
   
fnm2 = FNM1 + "ASSOSI.TXT"

If Dir(fnm2) = "" Then
   cmdKTABLES.Visible = False
   KTABLES = " "
   End If
   
fnm2 = FNM1 + "BG.TXT"

If Dir(fnm2) = "" Then
   cmdKTABLES.Visible = False
   KTABLES = " "
   End If
   
fnm2 = FNM1 + "FRAME.TXT"

If Dir(fnm2) = "" Then
   cmdKTABLES.Visible = False
   KTABLES = " "
   End If
  
fnm2 = FNM1 + "WFRAME.TXT"

If Dir(fnm2) <> "" Then
   cmdKTABLES.Visible = False
   KTABLES = " "
   End If
 
fnm2 = FNM1 + "SPECIES.TXT"

If Dir(fnm2) = "" Then
   cmdKTABLES.Visible = False
   KTABLES = " "
   End If
   
fnm2 = FNM1 + "UNITS.TXT"

If Dir(fnm2) = "" Then
   cmdKTABLES.Visible = False
   KTABLES = " "
   End If
'--------------------------------------------

If KTABLES = " " Then
   cmdEFFORT.Enabled = False
   cmdLAND.Enabled = False
   cmdACT.Enabled = False
   cmdESTIM.Enabled = False
   Exit Sub
   End If
   
'---------------------------------------------

cmdKEFFORT.Visible = True
cmdEFFORT.Enabled = True
cmdESTIM.Enabled = True
cmdREP.Enabled = True

KEFFORT = "Y"

FNM1 = APPROOT + "\ARTBAS\EFFORT\Y" + Format(ccy, "0000") + "M" + _
       Format(ccm, "00") + "_"

fnm2 = FNM1 + "ESAMPLES.TXT"

If Dir(fnm2) = "" Then
   KEFFORT = " "
   cmdKEFFORT.Visible = False
   End If
   
If KEFFORT = " " Then
   cmdESTIM.Enabled = False
   End If
'-----------------------------------------------------

cmdKLAND.Visible = True
cmdLAND.Enabled = True
cmdREP.Enabled = True
cmdESTIM.Enabled = True

KLAND = "Y"

FNM1 = APPROOT + "\ARTBAS\LANDINGS\Y" + Format(ccy, "0000") + "M" + _
       Format(ccm, "00") + "_"

fnm2 = FNM1 + "LSAMPLES.TXT"

If Dir(fnm2) = "" Then
   KLAND = " "
   cmdKLAND.Visible = False
   End If
   
If KLAND = " " Then
   cmdESTIM.Enabled = False
   End If

'-----------------------------------------------------

cmdKACTIVE.Visible = True
cmdACT.Enabled = True
cmdREP.Enabled = True
cmdESTIM.Enabled = True

KACT = "Y"

FNM1 = APPROOT + "\ARTBAS\TABLES\Y" + Format(ccy, "0000") + "M" + _
       Format(ccm, "00") + "_"

fnm2 = FNM1 + "ACTIVE.TXT"

If Dir(fnm2) = "" Then
   KACT = " "
   cmdKACTIVE.Visible = False
   End If
   
fnm2 = FNM1 + "WACTIVE.TXT"

If Dir(fnm2) <> "" Then
   KACT = " "
   cmdKACTIVE.Visible = False
   End If
   
If KACT = " " Then
   cmdESTIM.Enabled = False
   End If


'-----------------------------------------------------
If KLAND = " " Or KEFFORT = " " Or KACT = " " Then
   cmdESTIM.Enabled = False
   Exit Sub
   End If
   
cmdKESTIM.Visible = True
cmdESTIM.Enabled = True
cmdREP.Enabled = True
cmdESTIM.Enabled = True

KESTIM = "Y"

FNM1 = APPROOT + "\ARTBAS\RESULTS\Y" + Format(ccy, "0000") + "M" + _
       Format(ccm, "00") + "*.*"
fnm2 = FNM1

If Dir(fnm2) = "" Then
   KESTIM = " "
   cmdKESTIM.Visible = False
   End If
   
End Sub
Private Sub optEXPR_Click()

Dim FNMBAT, XXX, FNM1, fnm2, resp

' Write out the Export.Bat file for our installation location
FNMBAT = APPROOT + "\ARTBAS\EXPORT\EXPORT.BAT"

Open FNMBAT For Output As #1

Print #1, "ECHO OFF"
Print #1, "COPY " & APPROOT & "\ARTBAS\TABLES\Y*.TXT " & APPROOT & "\ARTBAS\ARTBASIC_RESULTS"

Close #1

RUN_CODE = Shell(FNMBAT, 4)

resp = MsgBox(msgtab(222), vbOKOnly, " ")

Call cmdRETO00_Click

End Sub
Private Sub optEXPT_Click()

Dim FNMBAT, XXX, FNM1, fnm2, resp

FNMBAT = APPROOT + "\ARTBAS\EXPORT\EXPORT.BAT"

Open FNMBAT For Output As #1

Print #1, "ECHO OFF"
Print #1, "COPY " & APPROOT & "\ARTBAS\TABLES\Y*.TXT " & APPROOT & "\ARTBAS\ARTBASIC_TABLES"

Close #1

RUN_CODE = Shell(FNMBAT, 4)

resp = MsgBox(msgtab(222), vbOKOnly, " ")

Call cmdRETO00_Click

End Sub
Private Sub optIMPR_Click()

If Dir(APPROOT + "\ARTBAS\ARTBASIC_RESULTS\*.*") = "" Then GoTo IMP_ERR

Dim FNMBAT, XXX, FNM1, fnm2, resp

FNMBAT = APPROOT + "\ARTBAS\EXPORT\IMPORT.BAT"

Open FNMBAT For Output As #1

Print #1, "ECHO OFF"

FNM1 = APPROOT + "\ARTBAS\ARTBASIC_RESULTS\*.*"
fnm2 = APPROOT + "\ARTBAS\TRANSFER"

Print #1, " COPY " + FNM1 + " " + fnm2

Close #1

RUN_CODE = Shell(FNMBAT, 4)

resp = MsgBox(msgtab(222), vbOKOnly, "  ")

Call cmdRETO00_Click

Exit Sub

IMP_ERR:

resp = MsgBox(msgtab(271), vbOKOnly, "  ")

Call cmdRETO00_Click

End Sub
Private Sub optIMPT_Click()

If Dir(APPROOT + "\ARTBAS\ARTBASIC_TABLES\*.*") = "" Then GoTo IMP_ERR

Dim FNMBAT, XXX, FNM1, fnm2, resp

FNMBAT = APPROOT + "\ARTS\EXPORT\IMPORT.BAT"

Open FNMBAT For Output As #1

Print #1, "ECHO OFF"

FNM1 = APPROOT + "\ARTBAS\ARTBASIC_TABLES\*.*"
fnm2 = APPROOT + "\ARTBAS\TABLES"

Print #1, " COPY " + FNM1 + " " + fnm2

Close #1

RUN_CODE = Shell(FNMBAT, 4)

resp = MsgBox(msgtab(222), vbOKOnly, "  ")

Call cmdRETO00_Click

Exit Sub

IMP_ERR:

resp = MsgBox(msgtab(270), vbOKOnly, "  ")

Call cmdRETO00_Click

End Sub
Private Sub CHECK_BKPA()

BKP_ERROR = "N"

If Dir("A:\BACKUP.TXT") = "" Then BKP_ERROR = "Y"
If Dir("A:\MY.TXT") = "" Then BKP_ERROR = "Y"

End Sub
Private Sub CHECK_IMPA()

IMP_ERROR = "N"

If Dir("A:\EXPORT.TXT") = "" Then IMP_ERROR = "Y"
If Dir("A:\MY.TXT") = "" Then IMP_ERROR = "Y"

End Sub
Private Sub CHECK_MY()

MY_ERROR = "N"

Open "A:\MY.TXT" For Input As #1

Dim xm, xy

Input #1, xm, xy

Close #1

If xm <> current_month Or xy <> current_year Then MY_ERROR = "Y"

End Sub
Private Sub SAVE_TABLES()

Dim fnm, FNM1, fnm2

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "*.*"

If Dir(fnm) = "" Then Exit Sub

fnm = Dir(fnm)

FNM1 = APPROOT + "\ARTBAS\TABLES\" + fnm
fnm2 = "A:\TABLES\" + fnm

FileCopy FNM1, fnm2

fnm = "?"

Do Until fnm = ""

fnm = Dir

If fnm = "" Then Exit Sub

FNM1 = APPROOT + "\ARTBAS\TABLES\" + fnm
fnm2 = "A:\TABLES\" + fnm

FileCopy FNM1, fnm2

Loop

End Sub
Private Sub COPY_TABLES()

Dim fnm, FNM1, fnm2

fnm = "A:\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "*.*"

If Dir(fnm) = "" Then Exit Sub

fnm = Dir(fnm)

fnm2 = APPROOT + "\ARTBAS\TABLES\" + fnm
FNM1 = "A:\TABLES\" + fnm

FileCopy FNM1, fnm2

fnm = "?"

Do Until fnm = ""

fnm = Dir

If fnm = "" Then Exit Sub

fnm2 = APPROOT + "\ARTBAS\TABLES\" + fnm
FNM1 = "A:\TABLES\" + fnm

FileCopy FNM1, fnm2

If RESTORE_FLAG = "YES" Then
   GoTo SKIP_SAVE
   End If

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + "M" + Format(current_month, "00") + _
    "_WACTIVE.TXT"
   
Open fnm For Output As #1
Close #1

SKIP_SAVE:

Loop

End Sub

Private Sub SAVE_RESULTS()

Dim fnm, FNM1, fnm2

fnm = APPROOT + "\ARTBAS\RESULTS\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "*.*"

If Dir(fnm) = "" Then Exit Sub

fnm = Dir(fnm)

FNM1 = APPROOT + "\ARTBAS\RESULTS\" + fnm
fnm2 = "A:\RESULTS\" + fnm

FileCopy FNM1, fnm2

fnm = "?"

Do Until fnm = ""

fnm = Dir

If fnm = "" Then Exit Sub

FNM1 = APPROOT + "\ARTBAS\RESULTS\" + fnm
fnm2 = "A:\RESULTS\" + fnm

FileCopy FNM1, fnm2

Loop

End Sub
Private Sub COPY_RESULTS()

Dim fnm, FNM1, fnm2

fnm = "A:\RESULTS\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "*.*"

If Dir(fnm) = "" Then Exit Sub

fnm = Dir(fnm)

fnm2 = APPROOT + "\ARTBAS\RESULTS\" + fnm
FNM1 = "A:\RESULTS\" + fnm

FileCopy FNM1, fnm2

fnm = "?"

Do Until fnm = ""

fnm = Dir

If fnm = "" Then Exit Sub

fnm2 = APPROOT + "\ARTBAS\RESULTS\" + fnm
FNM1 = "A:\RESULTS\" + fnm

FileCopy FNM1, fnm2

Loop

End Sub
Private Sub COPY_RESULTS2()

Dim fnm, FNM1, fnm2

fnm = "A:\RESULTS\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "*.*"

If Dir(fnm) = "" Then Exit Sub

fnm = Dir(fnm)

fnm2 = APPROOT + "\ARTBAS\TRANSFER\" + fnm
FNM1 = "A:\RESULTS\" + fnm

FileCopy FNM1, fnm2

fnm = "?"

Do Until fnm = ""

fnm = Dir

If fnm = "" Then Exit Sub

fnm2 = APPROOT + "\ARTBAS\TRANSFER\" + fnm
FNM1 = "A:\RESULTS\" + fnm

FileCopy FNM1, fnm2

Loop

End Sub
Private Sub SAVE_LANDINGS()

Dim fnm, FNM1, fnm2

fnm = APPROOT + "\ARTBAS\LANDINGS\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_LSAMPLES.TXT"

If Dir(fnm) = "" Then Exit Sub

fnm = Dir(fnm)

FNM1 = APPROOT + "\ARTBAS\LANDINGS\" + fnm
fnm2 = "A:\LANDINGS\" + fnm

FileCopy FNM1, fnm2

fnm = APPROOT + "\ARTBAS\LANDINGS\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_LSPECIES.TXT"

If Dir(fnm) = "" Then Exit Sub

fnm = Dir(fnm)

FNM1 = APPROOT + "\ARTBAS\LANDINGS\" + fnm
fnm2 = "A:\LANDINGS\" + fnm

FileCopy FNM1, fnm2

End Sub
Private Sub COPY_LANDINGS()

Dim fnm, FNM1, fnm2

fnm = "A:\LANDINGS\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_LSAMPLES.TXT"

If Dir(fnm) = "" Then Exit Sub

fnm = Dir(fnm)

fnm2 = APPROOT + "\ARTBAS\LANDINGS\" + fnm
FNM1 = "A:\LANDINGS\" + fnm

FileCopy FNM1, fnm2

fnm = "A:\LANDINGS\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_LSPECIES.TXT"

If Dir(fnm) = "" Then Exit Sub

fnm = Dir(fnm)

fnm2 = APPROOT + "\ARTBAS\LANDINGS\" + fnm
FNM1 = "A:\LANDINGS\" + fnm

FileCopy FNM1, fnm2

End Sub
Private Sub SAVE_EFFORT()

Dim fnm, FNM1, fnm2

fnm = APPROOT + "\ARTBAS\EFFORT\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_ESAMPLES.TXT"

If Dir(fnm) = "" Then Exit Sub

fnm = Dir(fnm)

FNM1 = APPROOT + "\ARTBAS\EFFORT\" + fnm
fnm2 = "A:\EFFORT\" + fnm

FileCopy FNM1, fnm2

End Sub
Private Sub COPY_EFFORT()

Dim fnm, FNM1, fnm2

fnm = "A:\EFFORT\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_ESAMPLES.TXT"

If Dir(fnm) = "" Then Exit Sub

fnm = Dir(fnm)

fnm2 = APPROOT + "\ARTBAS\EFFORT\" + fnm
FNM1 = "A:\EFFORT\" + fnm

FileCopy FNM1, fnm2

End Sub
Private Sub CHECK_EMPTYA()

EMPTY_ERROR = "N"

Dim fnm, resp

On Error GoTo EXIT_SUB

fnm = "A:\*.*"

If Dir(fnm) <> "" Then GoTo EXIT_SUB
  
MkDir "A:\LANDINGS"
MkDir "A:\EFFORT"
MkDir "A:\RESULTS"
MkDir "A:\TABLES"

Exit Sub

EXIT_SUB:

EMPTY_ERROR = "Y"

End Sub
Private Sub CHECK_RESULTS_EXIST()

PROCEED_FLAG = "Y"

Dim fnm, resp

fnm = APPROOT + "\ARTBAS\RESULTS\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "*.*"

If Dir(fnm) = "" Then Exit Sub

resp = MsgBox(msgtab(230), vbCritical + vbOKCancel, " ")

If resp = 2 Then
   PROCEED_FLAG = "N"
   Exit Sub
   End If

Kill fnm

End Sub



