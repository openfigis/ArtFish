VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmTABLES 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7305
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
   ScaleHeight     =   7305
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
      Left            =   8640
      MousePointer    =   1  'Arrow
      Picture         =   "frmTABLES.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   6000
      Width           =   735
   End
   Begin VB.CommandButton cmdENDLIST 
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
      Left            =   960
      MousePointer    =   1  'Arrow
      Picture         =   "frmTABLES.frx":2262
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   6000
      Width           =   495
   End
   Begin VB.CommandButton cmdPRLIST 
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
      Height          =   495
      Left            =   120
      MousePointer    =   1  'Arrow
      Picture         =   "frmTABLES.frx":24E4
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   6000
      Width           =   495
   End
   Begin VB.CommandButton cmdVIEW 
      BackColor       =   &H00FFFFFF&
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3480
      MousePointer    =   1  'Arrow
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   6240
      Width           =   735
   End
   Begin VB.CommandButton cmdADD 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   2640
      MousePointer    =   1  'Arrow
      Picture         =   "frmTABLES.frx":2766
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   6240
      Width           =   735
   End
   Begin VB.CommandButton cmdMINOR 
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
      Left            =   960
      MousePointer    =   1  'Arrow
      Picture         =   "frmTABLES.frx":29E8
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   6000
      Width           =   735
   End
   Begin VB.CommandButton cmdKUNITS 
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
      Height          =   255
      Left            =   7080
      Picture         =   "frmTABLES.frx":4E42
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   6840
      Width           =   255
   End
   Begin VB.CommandButton cmdKSPECIES 
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
      Height          =   255
      Left            =   6240
      Picture         =   "frmTABLES.frx":4F4C
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   6840
      Width           =   255
   End
   Begin VB.CommandButton cmdKFRAME 
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
      Height          =   255
      Left            =   5400
      Picture         =   "frmTABLES.frx":5056
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   6840
      Width           =   255
   End
   Begin VB.CommandButton cmdKBG 
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
      Height          =   255
      Left            =   4560
      Picture         =   "frmTABLES.frx":5160
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   6840
      Width           =   255
   End
   Begin VB.CommandButton cmdKASSO 
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
      Height          =   255
      Left            =   3720
      Picture         =   "frmTABLES.frx":526A
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   6840
      Width           =   255
   End
   Begin VB.CommandButton cmdKSITES 
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
      Height          =   255
      Left            =   2880
      Picture         =   "frmTABLES.frx":5374
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   6840
      Width           =   255
   End
   Begin VB.CommandButton cmdKASSOM 
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
      Height          =   255
      Left            =   2040
      Picture         =   "frmTABLES.frx":547E
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   6840
      Width           =   255
   End
   Begin VB.CommandButton cmdKMINOR 
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
      Height          =   255
      Left            =   1200
      Picture         =   "frmTABLES.frx":5588
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   6840
      Width           =   255
   End
   Begin VB.CommandButton cmdKMAJOR 
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
      Height          =   255
      Left            =   360
      Picture         =   "frmTABLES.frx":5692
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   6840
      Width           =   255
   End
   Begin VB.CommandButton cmdASSOM 
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
      Left            =   1800
      Picture         =   "frmTABLES.frx":579C
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   6000
      Width           =   735
   End
   Begin VB.CommandButton cmdBG 
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
      Left            =   4320
      Picture         =   "frmTABLES.frx":5A1E
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   6000
      Width           =   735
   End
   Begin VB.CommandButton cmdUNITS 
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
      Left            =   6840
      Picture         =   "frmTABLES.frx":5CA0
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   6000
      Width           =   735
   End
   Begin VB.CommandButton cmdFRAME 
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
      Left            =   5160
      Picture         =   "frmTABLES.frx":5F22
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6000
      Width           =   735
   End
   Begin VB.CommandButton cmdSPECIES 
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
      Left            =   6000
      Picture         =   "frmTABLES.frx":6088
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   6000
      Width           =   735
   End
   Begin VB.CommandButton cmdSITES 
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
      Left            =   2640
      Picture         =   "frmTABLES.frx":6B8A
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6000
      Width           =   735
   End
   Begin VB.CommandButton cmdOKUN 
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
      Height          =   375
      Left            =   7200
      Picture         =   "frmTABLES.frx":6E0C
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5400
      Width           =   375
   End
   Begin VB.TextBox txtMONEY 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   14
      Top             =   5520
      Width           =   1455
   End
   Begin VB.TextBox txtWEIGHT 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   13
      Top             =   5160
      Width           =   1455
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
      Picture         =   "frmTABLES.frx":6F16
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6240
      Width           =   735
   End
   Begin VB.CommandButton cmdASSO 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3480
      Picture         =   "frmTABLES.frx":7198
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6000
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
      Picture         =   "frmTABLES.frx":741A
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6240
      Width           =   735
   End
   Begin VB.CommandButton cmdMAJOR 
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
      Left            =   120
      Picture         =   "frmTABLES.frx":769C
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6000
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
      Picture         =   "frmTABLES.frx":95BE
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6240
      Width           =   735
   End
   Begin VB.CommandButton cmdNUM 
      BackColor       =   &H00FFFFFF&
      Caption         =   "1,2,.."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1800
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   3
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
      RecordSource    =   "GENTAB"
      Top             =   1560
      Width           =   2175
   End
   Begin MSDBGrid.DBGrid DBGGEN 
      Bindings        =   "frmTABLES.frx":B3B0
      Height          =   5415
      Left            =   120
      OleObjectBlob   =   "frmTABLES.frx":B3C5
      TabIndex        =   2
      Top             =   360
      Width           =   10455
   End
   Begin VB.CommandButton cmdQUIT 
      Appearance      =   0  'Flat
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
      Height          =   375
      Left            =   10200
      MousePointer    =   1  'Arrow
      Picture         =   "frmTABLES.frx":C0C6
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6600
      Width           =   375
   End
   Begin VB.CommandButton cmdRETURN 
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
      Height          =   1095
      Left            =   9480
      MousePointer    =   1  'Arrow
      Picture         =   "frmTABLES.frx":C348
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5880
      Width           =   1095
   End
   Begin RichTextLib.RichTextBox rtsDISP 
      Height          =   5655
      Left            =   120
      TabIndex        =   33
      Top             =   120
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   9975
      _Version        =   393217
      BackColor       =   12648447
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      MousePointer    =   1
      Appearance      =   0
      TextRTF         =   $"frmTABLES.frx":C5CA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FFFF&
      Caption         =   " 03"
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
      TabIndex        =   38
      Top             =   7080
      Width           =   255
   End
   Begin VB.Label lblCAL1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   11
      Top             =   0
      Width           =   3975
   End
   Begin VB.Label lblCAL2 
      BackStyle       =   0  'Transparent
      Caption         =   "1234567890123456789012345678901"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   10
      Top             =   240
      Width           =   3975
   End
   Begin VB.Label lblDEL 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   4320
      TabIndex        =   9
      Top             =   6480
      Width           =   4935
   End
   Begin VB.Label lblUPD 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   4320
      TabIndex        =   8
      Top             =   6000
      Width           =   4935
   End
End
Attribute VB_Name = "frmTABLES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private NGTAB, MJF, MNF, SIF, BGF, SPF, UNF, FRF, ASF1, ASF2
Private NEDIT, TC(), TN(), TR(), TS(), LSPT()
Private NUMN, NUMJ, NUSI, NUBG, NUSP, NUMYN, CURFNM, CURROW, CURDEL
Private NUM_FLAG, SPOK, ENDVIEW
Private WSI(), WMN(), WBG(), WSP()
Private Sub cmdADD_Click()

cmdVIEW.Visible = False
rtsDISP.Visible = False

Dim dbn As String

Dim I

dbn = APPROOT + "\ARTBAS\TABLES\GENTAB.MDB"

If Dir(dbn) = "" Then Exit Sub

dtaGEN.DatabaseName = dbn
dtaGEN.Refresh

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(dbn)
Set prm_record = prm_database.OpenRecordset("GENTAB")

With prm_record

Dim maxc
    maxc = 0
    .MoveFirst
    Do Until .EOF
    If ![CODE] > maxc Then maxc = ![CODE]
    .MoveNext
    Loop

For I = 1 To 10

.AddNew

![CODE] = maxc + I
![Sort] = "999999"
![remarks] = String(31, ".")
![Name] = "???"

.Update

Next I

End With

prm_record.Close
prm_database.Close

dtaGEN.DatabaseName = APPROOT + "\ARTBAS\STRUS\GENTAB.MDB"
dtaGEN.Refresh

dtaGEN.DatabaseName = APPROOT + "\ARTBAS\TABLES\GENTAB.MDB"
dtaGEN.Refresh

dbgGEN.AllowAddNew = False

End Sub
Private Sub cmdASSO_Click()

Dim fnm, RESP

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_ASSOSI.TXT"

If Dir(fnm) = "" Then GoTo NO_DISPLAY

RESP = MsgBox(msgtab(104) + Chr(13) + msgtab(106), vbOKCancel, " ")

If RESP = 2 Then Exit Sub

NO_DISPLAY:

txtWEIGHT.Visible = False
txtMONEY.Visible = False
cmdOKUN.Visible = False
frmTABLES.MousePointer = 13
Load frmASSO
Unload frmTABLES
frmASSO.Show

End Sub

Private Sub cmdASSOM_Click()

Dim fnm, RESP

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_ASSOMN.TXT"

If Dir(fnm) = "" Then GoTo NO_DISPLAY

RESP = MsgBox(msgtab(104) + Chr(13) + msgtab(105), vbOKCancel, " ")

If RESP = 2 Then Exit Sub

NO_DISPLAY:

txtWEIGHT.Visible = False
txtMONEY.Visible = False
cmdOKUN.Visible = False
frmTABLES.MousePointer = 13
Load frmASSOM
Unload frmTABLES
frmASSOM.Show

End Sub
Private Sub cmdBACK_Click()

cmdVIEW.Visible = False
rtsDISP.Visible = False

Dim RESP

RESP = MsgBox(msgtab(60), vbOKCancel, " ")

If RESP = 2 Then Exit Sub

Call CHECK_NULLS

lblCAL1.Visible = False
lblCAL2.Visible = False

frmTABLES.Caption = monthtab(current_month) + " " + _
                    Format(current_year, "0000") + " - " + msgtab(39)

dtaGEN.DatabaseName = APPROOT + "\ARTBAS\STRUS\GENAUX.MDB"
dtaGEN.Refresh

cmdMAJOR.Visible = True
cmdMINOR.Visible = True
cmdSITES.Visible = True
cmdASSO.Visible = True
cmdASSOM.Visible = True
cmdBG.Visible = True
cmdFRAME.Visible = True
cmdSPECIES.Visible = True
cmdUNITS.Visible = True
cmdRETURN.Visible = True
cmdQUIT.Visible = True
cmdGUIDE.Visible = True

cmdADD.Visible = False
cmdNUM.Visible = False
cmdNUM.Enabled = True
cmdEND.Visible = False
cmdPRINT.Visible = False
lblUPD.Visible = False
lblDEL.Visible = False
cmdBACK.Visible = False

dbgGEN.Visible = False

Call CONTROL_TABLES

End Sub
Private Sub cmdBG_Click()

NUM_FLAG = "Y"

Dim fnm

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_FRAME.TXT"

If Dir(fnm) <> "" Then NUM_FLAG = "N"

fnm = APPROOT + "\ARTBAS\EFFORT\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_ESAMPLES.TXT"

If Dir(fnm) <> "" Then NUM_FLAG = "N"

fnm = APPROOT + "\ARTBAS\LANDINGS\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_LSAMPLES.TXT"

If Dir(fnm) <> "" Then NUM_FLAG = "N"

txtWEIGHT.Visible = False
txtMONEY.Visible = False
cmdOKUN.Visible = False
cmdKMAJOR.Visible = False
cmdKMINOR.Visible = False
cmdKASSOM.Visible = False
cmdKSITES.Visible = False
cmdKASSO.Visible = False
cmdKBG.Visible = False
cmdKFRAME.Visible = False
cmdKSPECIES.Visible = False
cmdKSPECIES.Visible = False
cmdKUNITS.Visible = False

NUMJ = " ": NUMN = " ": NUSI = " ": NUBG = " ": NUSP = " "

NUBG = "Y"

MJF = " ": MNF = " ": SIF = " ": BGF = " ": SPF = " ": UNF = " ": FRF = " ": ASF1 = " ": ASF2 = " "

BGF = "Y"

frmTABLES.Caption = monthtab(current_month) + " " + _
                    Format(current_year, "0000") + " - " + msgtab(39) + _
                    " - " + msgtab(45)

Call MOVE_TABLE
Call PREP_GENTAB

End Sub
Private Sub cmdEND_Click()

cmdVIEW.Visible = False
rtsDISP.Visible = False

lblCAL1.Visible = False
lblCAL2.Visible = False

frmTABLES.Caption = monthtab(current_month) + " " + _
                    Format(current_year, "0000") + " - " + msgtab(39)

Call DUMP_TAB

dtaGEN.DatabaseName = APPROOT + "\ARTBAS\STRUS\GENAUX.MDB"
dtaGEN.Refresh

cmdMAJOR.Visible = True
cmdMINOR.Visible = True
cmdSITES.Visible = True
cmdASSO.Visible = True
cmdASSOM.Visible = True
cmdBG.Visible = True
cmdFRAME.Visible = True
cmdSPECIES.Visible = True
cmdUNITS.Visible = True
cmdRETURN.Visible = True
cmdQUIT.Visible = True
cmdGUIDE.Visible = True

cmdADD.Visible = False
cmdNUM.Visible = False
cmdNUM.Enabled = True
cmdEND.Visible = False
cmdPRINT.Visible = False
lblUPD.Visible = False
lblDEL.Visible = False
cmdBACK.Visible = False

dbgGEN.Visible = False

Call CONTROL_TABLES

End Sub

Private Sub cmdENDLIST_Click()

If Dir(APPROOT + "\ARTBAS\TABLES\WTEXT.TXT") <> "" Then Kill APPROOT + "\ARTBAS\TABLES\WTEXT.TXT"

Call cmdSPECIES_Click

End Sub

Private Sub cmdFRAME_Click()
txtWEIGHT.Visible = False
txtMONEY.Visible = False
cmdOKUN.Visible = False
frmTABLES.MousePointer = 13
Load frmFRAME
Unload frmTABLES
frmFRAME.Show

End Sub

Private Sub cmdGUIDE_Click()

HTYPE = "30"

HFNM = APPROOT + "\ARTBAS\HELP\" + current_language + "HELP" + HTYPE + ".rtf"

If Dir(HFNM) = "" Then Exit Sub

frmTABLES.Enabled = False
Load frmGUIDE
frmGUIDE.Show

End Sub

Private Sub cmdMINOR_Click()

Dim fnm, RESP

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_ASSOMN.TXT"

If Dir(fnm) = "" Then GoTo NO_DISPLAY_1

RESP = MsgBox(msgtab(104) + Chr(13) + msgtab(105), vbOKCancel, " ")

If RESP = 2 Then Exit Sub

NO_DISPLAY_1:

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_ASSOSI.TXT"

If Dir(fnm) = "" Then GoTo NO_DISPLAY_2

RESP = MsgBox(msgtab(104) + Chr(13) + msgtab(106), vbOKCancel, " ")

If RESP = 2 Then Exit Sub

NO_DISPLAY_2:

NUM_FLAG = "Y"

fnm = APPROOT + "\ARTBAS\EFFORT\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_ESAMPLES.TXT"

If Dir(fnm) <> "" Then NUM_FLAG = "N"

fnm = APPROOT + "\ARTBAS\LANDINGS\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_LSAMPLES.TXT"

If Dir(fnm) <> "" Then NUM_FLAG = "N"

txtWEIGHT.Visible = False
txtMONEY.Visible = False
cmdOKUN.Visible = False
cmdKMAJOR.Visible = False
cmdKMINOR.Visible = False
cmdKASSOM.Visible = False
cmdKSITES.Visible = False
cmdKASSO.Visible = False
cmdKBG.Visible = False
cmdKFRAME.Visible = False
cmdKSPECIES.Visible = False
cmdKSPECIES.Visible = False
cmdKUNITS.Visible = False

NUMJ = " ": NUMN = " ": NUSI = " ": NUBG = " ": NUSP = " "

NUMN = "Y"

MJF = " ": MNF = " ": SIF = " ": BGF = " ": SPF = " ": UNF = " ": FRF = " ": ASF1 = " ": ASF2 = " "

MNF = "Y"

lblCAL1.Visible = False
lblCAL2.Visible = False

frmTABLES.Caption = monthtab(current_month) + " " + _
                    Format(current_year, "0000") + " - " + msgtab(39) + _
                    " - " + msgtab(42)

Call MOVE_TABLE
Call PREP_GENTAB

End Sub
Private Sub cmdNUM_Click()

cmdVIEW.Visible = False
rtsDISP.Visible = False

Call CHECK_NUM

If NUMYN = "NO" Then Exit Sub

dtaGEN.DatabaseName = APPROOT + "\ARTBAS\STRUS\GENAUX.MDB"
dtaGEN.Refresh

Open APPROOT + "\ARTBAS\TABLES\GENTAB.TXT" For Output As #1

Dim I

Dim dbn As String

dbn = APPROOT + "\ARTBAS\TABLES\GENTAB.MDB"

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(dbn)
Set prm_record = prm_database.OpenRecordset("GENTAB")

With prm_record

.Index = "SORT"
.MoveFirst

I = 0

Do Until .EOF

If RTrim(![Name]) <> "???" And RTrim(![Name]) <> "" Then

I = I + 1

Print #1, Format(I, "0000") + " " + _
          Left(![Name] + Space(30), 30) + " " + _
          Left(![remarks] + Space(31), 31) + " " + _
          Left(![Sort] + Space(6), 6)
End If

.MoveNext

Loop

End With

Close #1

prm_record.Close
prm_database.Close

Call PREP_GENTAB

End Sub

Private Sub cmdOKUN_Click()

Dim fnm

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_UNITS.TXT"

UNW = txtWEIGHT.Text: UNW = RTrim(UNW)

If Len(UNW) = 0 Then
   txtWEIGHT.Visible = False
   txtMONEY.Visible = False
   cmdOKUN.Visible = False
   cmdKUNITS.Visible = False
   Kill fnm
   Exit Sub
   End If

UNM = txtMONEY.Text: UNM = RTrim(UNM)

If Len(UNM) = 0 Then
   txtWEIGHT.Visible = False
   txtMONEY.Visible = False
   cmdOKUN.Visible = False
   cmdKUNITS.Visible = False
   Kill fnm
   Exit Sub
   End If

txtWEIGHT.Visible = False
txtMONEY.Visible = False
cmdOKUN.Visible = False

Open fnm For Output As #1

Print #1, UNW
Print #1, UNM

Close #1

End Sub

Private Sub cmdPRINT_Click()

cmdVIEW.Visible = False
rtsDISP.Visible = False

Printer.FontBold = True
Printer.FontName = "Courier"
Printer.FontName = "Courier New"
Printer.FontSize = 11

Dim I, pageno, lineno

pageno = 0

GoSub CHANGE_PAGE

Dim XXX, dbn As String

dbn = APPROOT + "\ARTBAS\TABLES\GENTAB.MDB"

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(dbn)
Set prm_record = prm_database.OpenRecordset("GENTAB")

With prm_record

.Index = "SORT"
.MoveFirst

Do Until .EOF

XXX = RTrim(![Name])

If Len(XXX) <> 0 Then

Printer.Print Tab(4); Format(![CODE], "0000"); _
              Tab(12); Left(![Name] + Space(30), 30); _
              Tab(44); Left(![remarks] + Space(31), 31); _
              Tab(76); Left(![Sort] + Space(6), 6)

lineno = lineno + 1

If lineno > 55 Then GoSub CHANGE_PAGE

End If

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

Printer.Print Tab(4); frmTABLES.Caption

Printer.Print

Printer.FontItalic = True
   
Printer.Print Tab(4); dbgGEN.Columns(0).Caption; _
              Tab(12); dbgGEN.Columns(1).Caption; _
              Tab(44); dbgGEN.Columns(2).Caption; _
              Tab(76); dbgGEN.Columns(3).Caption
Printer.Print Tab(4); String(78, "-")

Printer.FontItalic = False
Printer.FontUnderline = False

Return
'====================================

End Sub

Private Sub cmdPRLIST_Click()

Call PRINT_LIST

End Sub
Private Sub cmdRETURN_Click()

Call CHECK_STATUS

Dim dbn

dbn = APPROOT + "\ARTBAS\TABLES\GENTAB.MDB"

dtaGEN.DatabaseName = APPROOT + "\ARTBAS\STRUS\GENAUX.MDB"
dtaGEN.Refresh

If Dir(dbn) <> "" Then Kill dbn

dbn = APPROOT + "\ARTBAS\TABLES\WORK.TXT"

If Dir(dbn) <> "" Then Kill dbn

dbn = APPROOT + "\ARTBAS\TABLES\GENTAB.TXT"

If Dir(dbn) <> "" Then Kill dbn

cmdRETURN.MousePointer = 13
frmTABLES.MousePointer = 13
Load frmARTB01
Unload frmTABLES
frmARTB01.Show

End Sub
Private Sub cmdQUIT_Click()

Call CHECK_STATUS

Beep

Dim dbn

dbn = APPROOT + "\ARTBAS\TABLES\GENTAB.MDB"

dtaGEN.DatabaseName = APPROOT + "\ARTBAS\STRUS\GENAUX.MDB"
dtaGEN.Refresh

If Dir(dbn) <> "" Then Kill dbn

dbn = APPROOT + "\ARTBAS\TABLES\WORK.TXT"

If Dir(dbn) <> "" Then Kill dbn

dbn = APPROOT + "\ARTBAS\TABLES\GENTAB.TXT"

If Dir(dbn) <> "" Then Kill dbn

cmdRETURN.MousePointer = 13
cmdQUIT.MousePointer = 13
frmTABLES.MousePointer = 13
Call write_parms
Unload frmTABLES

End

End Sub
Private Sub PREP_GENTAB()

NEDIT = 0

If MNF <> "Y" Then dbgGEN.Columns(2).Caption = msgtab(55)

dtaGEN.DatabaseName = APPROOT + "\ARTBAS\STRUS\GENAUX.MDB"
dtaGEN.Refresh

cmdMAJOR.Visible = False
cmdMINOR.Visible = False
cmdSITES.Visible = False
cmdASSO.Visible = False
cmdASSOM.Visible = False
cmdBG.Visible = False
cmdFRAME.Visible = False
cmdSPECIES.Visible = False
cmdUNITS.Visible = False
cmdRETURN.Visible = False
cmdQUIT.Visible = False
cmdGUIDE.Visible = False

cmdEND.Visible = True
cmdNUM.Visible = False

If NUM_FLAG = "Y" Then cmdNUM.Visible = False

cmdADD.Visible = True
cmdPRINT.Visible = True
cmdBACK.Visible = True

lblDEL.Visible = True

Dim N, I, XXX, m, yyy, WWW, CODE(), NME(), RM(), SRT(), OLDC

Open APPROOT + "\ARTBAS\TABLES\GENTAB.TXT" For Input As #1

N = 0

Do Until EOF(1)

Line Input #1, XXX

N = N + 1

ReDim Preserve CODE(1 To N), NME(1 To N), RM(1 To N), SRT(1 To N)

CODE(N) = Mid(XXX, 1, 4)
NME(N) = Mid(XXX, 6, 30): NME(N) = RTrim(NME(N))
RM(N) = Mid(XXX, 37, 31): WWW = RM(N)
RM(N) = RTrim(RM(N))
SRT(N) = Mid(XXX, 69, 6)
SRT(N) = RTrim(SRT(N)): SRT(N) = Left(SRT(N) + "000000", 6)

Loop

Close #1

Dim dbn As String

FileCopy APPROOT + "\ARTBAS\STRUS\GENTAB.MDB", APPROOT + "\ARTBAS\TABLES\GENTAB.MDB"

dbn = APPROOT + "\ARTBAS\TABLES\GENTAB.MDB"

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(dbn)
Set prm_record = prm_database.OpenRecordset("GENTAB")

With prm_record

.Index = "CODE"

For I = 1 To N

.AddNew

![CODE] = CODE(I)
![Name] = NME(I)
![remarks] = RM(I)
![Sort] = SRT(I)

.Update

Next I

End With

prm_record.Close
prm_database.Close

dtaGEN.DatabaseName = APPROOT + "\ARTBAS\TABLES\GENTAB.MDB"
dtaGEN.Refresh

dbgGEN.Visible = True

End Sub
Private Sub cmdMAJOR_Click()

Dim fnm, RESP

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_ASSOMN.TXT"

If Dir(fnm) = "" Then GoTo NO_DISPLAY

RESP = MsgBox(msgtab(104) + Chr(13) + msgtab(105), vbOKCancel, " ")

If RESP = 2 Then Exit Sub

NO_DISPLAY:

NUM_FLAG = "Y"

txtWEIGHT.Visible = False
txtMONEY.Visible = False
cmdOKUN.Visible = False

cmdKMAJOR.Visible = False
cmdKMINOR.Visible = False
cmdKASSOM.Visible = False
cmdKSITES.Visible = False
cmdKASSO.Visible = False
cmdKBG.Visible = False
cmdKFRAME.Visible = False
cmdKSPECIES.Visible = False
cmdKSPECIES.Visible = False
cmdKUNITS.Visible = False

NUMJ = " ": NUMN = " ": NUSI = " ": NUBG = " ": NUSP = " "

NUMJ = "Y"

MJF = " ": MNF = " ": SIF = " ": BGF = " ": SPF = " ": UNF = " ": FRF = " ": ASF1 = " ": ASF2 = " "

MJF = "Y"

frmTABLES.Caption = monthtab(current_month) + " " + _
                    Format(current_year, "0000") + " - " + msgtab(39) + _
                    " - " + msgtab(41)

Call MOVE_TABLE
Call PREP_GENTAB

End Sub
Private Sub DUMP_TAB()

Open APPROOT + "\ARTBAS\TABLES\GENTAB.TXT" For Output As #1

Dim J
Dim XXX, dbn, WWW, yyy As String

dbn = APPROOT + "\ARTBAS\TABLES\GENTAB.MDB"

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(dbn)
Set prm_record = prm_database.OpenRecordset("GENTAB")

With prm_record

.Index = "primarykey"

For J = 1 To NEDIT

.Seek "=", TC(J)

If .NoMatch = True Then

    Dim maxc
    maxc = 0
    .MoveFirst
    Do Until .EOF
    If ![CODE] > maxc Then maxc = ![CODE]
    .MoveNext
    Loop
   
   .AddNew
   
   ![CODE] = maxc + 1
   ![Name] = TN(J)
   
   If Len(TR(J)) = 0 Then TR(J) = String(31, ".")
   
   ![remarks] = TR(J)
   
   If Len(TS(J)) = 0 Then TS(J) = "999999"
   
   ![Sort] = TS(J)
    
   .Update
   
   GoTo next_j
   
   End If

.Edit

![Name] = TN(J)
![remarks] = TR(J)
![Sort] = TS(J)

.Update

next_j:

Next J

NEDIT = 0

.Index = "SORT"

.MoveFirst

Do Until .EOF

XXX = RTrim(![Name])

If Len(XXX) <> 0 And XXX <> "???" Then

yyy = Left(![remarks] + Space(31), 31)

Print #1, Format(![CODE], "0000") + " " + _
          Left(![Name] + Space(30), 30) + " " + _
          yyy + " " + _
          Left(![Sort] + Space(6), 6)
End If

.MoveNext

Loop

End With

Close #1

prm_record.Close
prm_database.Close

Call CREATE_FILE

End Sub
Private Sub cmdSITES_Click()

Dim fnm, RESP

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_ASSOSI.TXT"

If Dir(fnm) = "" Then GoTo NO_DISPLAY

RESP = MsgBox(msgtab(104) + Chr(13) + msgtab(106), vbOKCancel, " ")

If RESP = 2 Then Exit Sub

NO_DISPLAY:

NUM_FLAG = "Y"

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_FRAME.TXT"

If Dir(fnm) <> "" Then NUM_FLAG = "N"

fnm = APPROOT + "\ARTBAS\EFFORT\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_ESAMPLES.TXT"

If Dir(fnm) <> "" Then NUM_FLAG = "N"

fnm = APPROOT + "\ARTBAS\LANDINGS\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_LSAMPLES.TXT"

If Dir(fnm) <> "" Then NUM_FLAG = "N"

txtWEIGHT.Visible = False
txtMONEY.Visible = False
cmdOKUN.Visible = False
cmdKMAJOR.Visible = False
cmdKMINOR.Visible = False
cmdKASSOM.Visible = False
cmdKSITES.Visible = False
cmdKASSO.Visible = False
cmdKBG.Visible = False
cmdKFRAME.Visible = False
cmdKSPECIES.Visible = False
cmdKSPECIES.Visible = False
cmdKUNITS.Visible = False

NUMJ = " ": NUMN = " ": NUSI = " ": NUBG = " ": NUSP = " "

NUSI = "Y"

MJF = " ": MNF = " ": SIF = " ": BGF = " ": SPF = " ": UNF = " ": FRF = " ": ASF1 = " ": ASF2 = " "

SIF = "Y"

frmTABLES.Caption = monthtab(current_month) + " " + _
                    Format(current_year, "0000") + " - " + msgtab(39) + _
                    " - " + msgtab(43)

Call MOVE_TABLE
Call PREP_GENTAB

End Sub
Private Sub cmdSPECIES_Click()

cmdPRLIST.Visible = False
cmdENDLIST.Visible = False
rtsDISP.Visible = False

Call PREP_WTABS

Dim DELF, XXX, K

ReDim LSPT(1 To 10000)

NUM_FLAG = "Y"

Dim fnm

fnm = APPROOT + "\ARTBAS\LANDINGS\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_LSPECIES.TXT"

If Dir(fnm) <> "" Then NUM_FLAG = "N"

If Dir(fnm) <> "" Then

   Open fnm For Input As #3
   
   Do Until EOF(3)
   
   Line Input #3, XXX
   K = Val(Mid(XXX, 10, 4))
   LSPT(K) = "Y"
   
   Loop
   
   Close #3
   
   End If

txtWEIGHT.Visible = False
txtMONEY.Visible = False
cmdOKUN.Visible = False
cmdKMAJOR.Visible = False
cmdKMINOR.Visible = False
cmdKASSOM.Visible = False
cmdKSITES.Visible = False
cmdKASSO.Visible = False
cmdKBG.Visible = False
cmdKFRAME.Visible = False
cmdKSPECIES.Visible = False
cmdKSPECIES.Visible = False
cmdKUNITS.Visible = False

NUMJ = " ": NUMN = " ": NUSI = " ": NUBG = " ": NUSP = " "

NUSP = "Y"

MJF = " ": MNF = " ": SIF = " ": BGF = " ": SPF = " ": UNF = " ": FRF = " ": ASF1 = " ": ASF2 = " "

SPF = "Y"

frmTABLES.Caption = monthtab(current_month) + " " + _
                    Format(current_year, "0000") + " - " + msgtab(39) + _
                    " - " + msgtab(47)

Call MOVE_TABLE
Call PREP_GENTAB

End Sub

Private Sub cmdUNITS_Click()

cmdKMAJOR.Visible = False
cmdKMINOR.Visible = False
cmdKASSOM.Visible = False
cmdKSITES.Visible = False
cmdKASSO.Visible = False
cmdKBG.Visible = False
cmdKFRAME.Visible = False
cmdKSPECIES.Visible = False
cmdKSPECIES.Visible = False
cmdKUNITS.Visible = False

txtWEIGHT.Visible = True
txtMONEY.Visible = True
cmdOKUN.Visible = True

Dim fnm

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_UNITS.TXT"

If Dir(fnm) = "" Then
   Open fnm For Output As #1
   Print #1, "Kg"
   Print #1, "US$"
   Close #1
   End If

Open fnm For Input As #1

Line Input #1, UNW
Line Input #1, UNM

UNW = RTrim(UNW): UNM = RTrim(UNM)

Close #1

txtWEIGHT.Text = UNW: txtMONEY.Text = UNM

Call CONTROL_TABLES

End Sub
Private Sub cmdVIEW_Click()

Call LIST_DOCS

End Sub
Private Sub DBGGEN_AfterColEdit(ByVal ColIndex As Integer)

Dim RESP

CURROW = dbgGEN.Row

If Len(dbgGEN.Columns(1).Value) = 0 And SPF = "Y" Then
   CURDEL = dbgGEN.Columns(0).Value
   Call CHECK_LANDINGS
   If SPOK = "N" Then Call cmdSPECIES_Click
      
   If SPOK = "W" Then
     cmdVIEW.Visible = True
     Call cmdSPECIES_Click
     End If
     
   End If

If Len(dbgGEN.Columns(1).Value) >= 30 Then
   dbgGEN.Columns(1).Value = Left(dbgGEN.Columns(1).Value, 30)
   End If
   
If Len(dbgGEN.Columns(2).Value) >= 31 Then
   dbgGEN.Columns(2).Value = Left(dbgGEN.Columns(2).Value, 31)
   End If
   
If Len(dbgGEN.Columns(3).Value) >= 6 Then
   dbgGEN.Columns(3).Value = Left(dbgGEN.Columns(3).Value, 6)
   End If

Dim I

NEDIT = NEDIT + 1: I = NEDIT

ReDim Preserve TC(1 To NEDIT), TN(1 To NEDIT), TR(1 To NEDIT), TS(1 To NEDIT)

TC(I) = dbgGEN.Columns(0).Value
TN(I) = dbgGEN.Columns(1).Value
TR(I) = dbgGEN.Columns(2).Value
TS(I) = dbgGEN.Columns(3).Value

End Sub
Private Sub CHECK_LANDINGS()

SPOK = "Y"

Dim K, RESP

K = dbgGEN.Columns(0).Value

If LSPT(K) <> "Y" Then Exit Sub

SPOK = "N"

RESP = MsgBox(msgtab(234) + Chr(13) + msgtab(235) + Chr(13) + msgtab(236), _
     vbCritical + vbOKCancel, " ")

If RESP = 2 Then Exit Sub

SPOK = "W"

End Sub
Private Sub LIST_DOCS()

dbgGEN.Visible = False
rtsDISP.Visible = True

cmdPRLIST.Visible = True
cmdENDLIST.Visible = True

cmdEND.Visible = False
cmdPRINT.Visible = False
cmdNUM.Visible = False
cmdADD.Visible = False
cmdVIEW.Visible = False
cmdBACK.Visible = False
lblDEL.Visible = False

Dim fnm, XXX, K, PPP

fnm = APPROOT + "\ARTBAS\LANDINGS\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_LSPECIES.TXT"

If Dir(fnm) = "" Then GoTo END_SUB

Open fnm For Input As #1
Open APPROOT + "\ARTBAS\TABLES\WTEXT.TXT" For Output As #2

Print #2, msgtab(235) + " ( " + RTrim(WSP(CURDEL)) + " )"
Print #2, " "

Do Until EOF(1)

Line Input #1, XXX

If CURDEL <> Val(Mid(XXX, 10, 4)) Then GoTo next_rec

PPP = Mid(XXX, 2, 6) + " "

K = Val(Mid(XXX, 95, 4))

PPP = PPP + RTrim(WMN(K)) + ", "

K = Val(Mid(XXX, 101, 4))

PPP = PPP + RTrim(WSI(K)) + ", "

K = Val(Mid(XXX, 107, 4))

PPP = PPP + RTrim(WBG(K))

Print #2, PPP

next_rec:

Loop

Close #1
Close #2

rtsDISP.FileName = APPROOT + "\ARTBAS\TABLES\WTEXT.TXT"

rtsDISP.Refresh

END_SUB:

End Sub
Private Sub PRINT_LIST()

Printer.FontBold = True
Printer.FontName = "Courier"
Printer.FontName = "Courier New"
Printer.FontSize = 8

Dim XXX, NOD, NOP

NOD = 0: NOP = 0

Open APPROOT + "\ARTBAS\TABLES\WTEXT.TXT" For Input As #4

Printer.Print " "
Printer.Print " "
Printer.Print " "

Do Until EOF(4)

Line Input #4, XXX

Printer.Print Tab(3); XXX
   
Loop

Close #4

Printer.EndDoc

End Sub
Private Sub PREP_WTABS()

ReDim WSI(1 To 10000), WBG(1 To 10000), WMN(1 To 10000), WSP(1 To 10000)

Dim fnm, K, XXX

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_SPECIES.TXT"

If Dir(fnm) = "" Then GoTo LOAD_SI

Open fnm For Input As #4

Do Until EOF(4)

Line Input #4, XXX

K = Val(Left(XXX, 4)): WSP(K) = Mid(XXX, 6, 30)

Loop

Close #4

LOAD_SI:

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_SITES.TXT"

If Dir(fnm) = "" Then GoTo LOAD_BG

Open fnm For Input As #4

Do Until EOF(4)

Line Input #4, XXX

K = Val(Left(XXX, 4)): WSI(K) = Mid(XXX, 6, 30)

Loop

Close #4

LOAD_BG:

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_BG.TXT"

If Dir(fnm) = "" Then GoTo LOAD_MN

Open fnm For Input As #4

Do Until EOF(4)

Line Input #4, XXX

K = Val(Left(XXX, 4)): WBG(K) = Mid(XXX, 6, 30)

Loop

Close #4

LOAD_MN:

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_MINOR.TXT"

If Dir(fnm) = "" Then Exit Sub

Open fnm For Input As #4

Do Until EOF(4)

Line Input #4, XXX

K = Val(Left(XXX, 4)): WMN(K) = Mid(XXX, 6, 30)

Loop

Close #4

End Sub
Private Sub DBGGEN_AfterUpdate()
NEDIT = 0
End Sub
Private Sub DBGGEN_KeyUp(KeyCode As Integer, Shift As Integer)

If Len(dbgGEN.Columns(1).Value) >= 30 Then
   dbgGEN.Columns(1).Value = Left(dbgGEN.Columns(1).Value, 30)
   End If
   
If Len(dbgGEN.Columns(2).Value) >= 31 Then
   dbgGEN.Columns(2).Value = Left(dbgGEN.Columns(2).Value, 31)
   End If
   
If Len(dbgGEN.Columns(3).Value) >= 6 Then
   dbgGEN.Columns(3).Value = Left(dbgGEN.Columns(3).Value, 6)
   End If

With dbgGEN

If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then NEDIT = 0

End With

End Sub
Private Sub DBGGEN_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

With dbgGEN

If .Col = 0 Then .Col = 1

End With

End Sub
Private Sub Form_Load()

dtaGEN.DatabaseName = APPROOT + "\ARTBAS\STRUS\GENAUX.mdb"

Set Picture = LoadPicture(APPROOT + "\ARTBAS\PICS_RUNTIME\SCREEN_03.JPG")

Dim cal1, cal2

cal2 = "1234567890123456789012345678901"
cal1 = "         1         2         3 "

lblCAL1.Caption = Left(cal1, CURCAL)
lblCAL2.Caption = Left(cal2, CURCAL)

lblCAL1.Visible = False
lblCAL2.Visible = False

txtWEIGHT.Visible = False
txtMONEY.Visible = False
cmdOKUN.Visible = False

MJF = " ": MNF = " ": SIF = " ": BGF = " ": SPF = " ": UNF = " ": FRF = " ": ASF1 = " ": ASF2 = " "

frmTABLES.Caption = monthtab(current_month) + " " + _
                    Format(current_year, "0000") + " - " + msgtab(39)
                    
cmdVIEW.ToolTipText = msgtab(239)
cmdENDLIST.ToolTipText = msgtab(240)
cmdPRLIST.ToolTipText = msgtab(52)

cmdPRLIST.Visible = False
cmdENDLIST.Visible = False
cmdVIEW.Visible = False
rtsDISP.Visible = False
cmdADD.Visible = False
cmdNUM.Visible = False
cmdEND.Visible = False
cmdPRINT.Visible = False
lblUPD.Visible = False
lblDEL.Visible = False
cmdBACK.Visible = False

dtaGEN.Visible = False
dbgGEN.Visible = False

dbgGEN.Columns(0).Caption = msgtab(53)
dbgGEN.Columns(1).Caption = msgtab(54)
dbgGEN.Columns(2).Caption = msgtab(55)
dbgGEN.Columns(3).Caption = msgtab(56)

cmdADD.ToolTipText = msgtab(119)
cmdMAJOR.ToolTipText = msgtab(41)
cmdMINOR.ToolTipText = msgtab(42)
cmdSITES.ToolTipText = msgtab(43)
cmdASSO.ToolTipText = msgtab(44)
cmdASSOM.ToolTipText = msgtab(63)
cmdFRAME.ToolTipText = msgtab(46)
cmdBG.ToolTipText = msgtab(45)
cmdSPECIES.ToolTipText = msgtab(47)
cmdUNITS.ToolTipText = msgtab(48)
cmdRETURN.ToolTipText = msgtab(49)
cmdQUIT.ToolTipText = msgtab(3)
cmdNUM.ToolTipText = msgtab(50)
cmdEND.ToolTipText = msgtab(51)
cmdPRINT.ToolTipText = msgtab(52)
cmdBACK.ToolTipText = msgtab(60)
cmdOKUN.ToolTipText = msgtab(64)
cmdGUIDE.ToolTipText = msgtab(243)

lblUPD.Caption = msgtab(58)
lblDEL.Caption = msgtab(59)

dbgGEN.AllowAddNew = False

Call CONTROL_TABLES

End Sub
Private Sub CREATE_FILE()

Dim MYF, fnm

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_"

If MJF = "Y" Then
   Call DELETE_TABLES
   MJF = " "
   fnm = fnm + "MAJOR.TXT"
   End If

If MNF = "Y" Then
   Call DELETE_TABLES
   MNF = " "
   fnm = fnm + "MINOR.TXT"
   End If

If SIF = "Y" Then
   Call DELETE_TABLES
   SIF = " "
   fnm = fnm + "SITES.TXT"
   End If

If BGF = "Y" Then
   Call DELETE_TABLES
   BGF = " "
   fnm = fnm + "BG.TXT"
   End If

If SPF = "Y" Then
   SPF = " "
   fnm = fnm + "SPECIES.TXT"
   End If

FileCopy APPROOT + "\ARTBAS\TABLES\GENTAB.TXT", fnm

Open fnm For Input As #1

Dim N, XXX

N = 0

Do Until EOF(1)

Line Input #1, XXX

XXX = Mid(XXX, 6, 30): XXX = RTrim(XXX)

If Len(XXX) <> 0 And Left(XXX, 3) <> "???" Then N = N + 1

Loop

Close #1

If N = 0 Then Kill fnm

End Sub
Private Sub MOVE_TABLE()

Dim MYF, fnm

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_"

If MJF = "Y" Then
   fnm = fnm + "MAJOR.TXT"
   End If

If MNF = "Y" Then
   fnm = fnm + "MINOR.TXT"
   End If

If SIF = "Y" Then
   fnm = fnm + "SITES.TXT"
   End If

If BGF = "Y" Then
   fnm = fnm + "BG.TXT"
   End If

If SPF = "Y" Then
   fnm = fnm + "SPECIES.TXT"
   End If

If Dir(fnm) = "" Then
   FileCopy APPROOT + "\ARTBAS\STRUS\BLTAB.TXT", fnm
   FileCopy fnm, APPROOT + "\ARTBAS\TABLES\GENTAB.TXT"
   Exit Sub
   End If

FileCopy fnm, APPROOT + "\ARTBAS\TABLES\GENTAB.TXT"

End Sub
Private Sub CONTROL_TABLES()

Dim KMAJOR, KMINOR, KASSOM, KSITES, KASSO, KBG, KFRAME, KSPECIES, KUNITS

KMAJOR = " ": KMINOR = " ": KASSOM = " ": KSITES = " ": KASSO = " "
KBG = " ": KFRAME = " ": KSPECIES = " ": KUNITS = " "

cmdMAJOR.Enabled = True
cmdMINOR.Enabled = True
cmdASSOM.Enabled = True
cmdSITES.Enabled = True
cmdASSO.Enabled = True
cmdBG.Enabled = True
cmdFRAME.Enabled = True
cmdSPECIES.Enabled = True
cmdSPECIES.Enabled = True
cmdUNITS.Enabled = True

cmdKMAJOR.Visible = False
cmdKMINOR.Visible = False
cmdKASSOM.Visible = False
cmdKSITES.Visible = False
cmdKASSO.Visible = False
cmdKBG.Visible = False
cmdKFRAME.Visible = False
cmdKSPECIES.Visible = False
cmdKSPECIES.Visible = False
cmdKUNITS.Visible = False

Dim fnm

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_MAJOR.TXT"
      
If Dir(fnm) <> "" Then
   KMAJOR = "Y"
   cmdKMAJOR.Visible = True
   End If
   
fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_MINOR.TXT"
      
If Dir(fnm) <> "" Then
   cmdKMINOR.Visible = True
   KMINOR = "Y"
   End If
   
fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_ASSOMN.TXT"
      
If Dir(fnm) <> "" Then
   cmdKASSOM.Visible = True
   KASSOM = "Y"
   End If
   
fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_SITES.TXT"
      
If Dir(fnm) <> "" Then
   cmdKSITES.Visible = True
   KSITES = "Y"
   End If
   
fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_ASSOSI.TXT"
      
If Dir(fnm) <> "" Then
   cmdKASSO.Visible = True
   KASSO = "Y"
   End If
   
fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_BG.TXT"
      
If Dir(fnm) <> "" Then
   cmdKBG.Visible = True
   KBG = "Y"
   End If
   
fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_FRAME.TXT"
      
If Dir(fnm) <> "" Then
   cmdKFRAME.Visible = True
   cmdFRAME.Visible = True
   KFRAME = "Y"
   End If
   
fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_WFRAME.TXT"
      
If Dir(fnm) <> "" Then
   cmdKFRAME.Visible = False
   cmdFRAME.Visible = True
   KFRAME = "N"
   End If
   
fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_SPECIES.TXT"
      
If Dir(fnm) <> "" Then
   cmdKSPECIES.Visible = True
   KSPECIES = "Y"
   End If
   
fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_UNITS.TXT"
      
If Dir(fnm) <> "" Then
   cmdKUNITS.Visible = True
   KUNITS = "Y"
   End If
'=================================================================

If KMAJOR <> "Y" Then
   cmdASSOM.Enabled = False
   cmdKASSOM.Visible = False
   End If
  
If KMINOR <> "Y" Then
   cmdASSOM.Enabled = False
   cmdASSO.Enabled = False
   cmdKASSOM.Visible = False
   cmdKASSO.Visible = False
   cmdKASSO.Visible = False
   End If
  
If KSITES <> "Y" Then
   cmdASSO.Enabled = False
   cmdKASSO.Visible = False
   cmdFRAME.Enabled = False
   cmdKFRAME.Visible = False
   End If
  
If KBG <> "Y" Then
   cmdFRAME.Enabled = False
   cmdKFRAME.Visible = False
   End If

End Sub
Private Sub CHECK_STATUS()

Dim fnm, ccy, ccm

ccy = current_year: ccm = current_month

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(ccy, "0000") + "M" + Format(ccm, "00") + "_MAJOR.TXT"

CY(ccy - 1989, ccm) = " "

If Dir(fnm) <> "" Then CY(ccy - 1989, ccm) = "X"

Call write_parms

End Sub
Private Sub DELETE_TABLES()

Dim fnm, ccy, ccm

ccy = current_year: ccm = current_month

If MJF = "Y" Then
   
   fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(ccy, "0000") + "M" + Format(ccm, "00") _
   + "_ASSOMN.TXT"
   
   If Dir(fnm) <> "" Then Kill fnm
   
   End If

If MNF = "Y" Then

      fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(ccy, "0000") + "M" + Format(ccm, "00") _
      + "_ACTIVE.TXT"
   
      If Dir(fnm) <> "" Then
   
      fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(ccy, "0000") + "M" + Format(ccm, "00") _
            + "_WACTIVE.TXT"
            
      Open fnm For Output As #2
      Close #2
      
      End If
   
   fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(ccy, "0000") + "M" + Format(ccm, "00") _
   + "_ASSOMN.TXT"
   
   If Dir(fnm) <> "" Then Kill fnm
   
   fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(ccy, "0000") + "M" + Format(ccm, "00") _
   + "_ASSOSI.TXT"
   
   If Dir(fnm) <> "" Then Kill fnm
   
   End If

If SIF = "Y" Then
    
   fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(ccy, "0000") + "M" + Format(ccm, "00") _
   + "_FRAME.TXT"
   
   If Dir(fnm) <> "" Then
   
      fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(ccy, "0000") + "M" + Format(ccm, "00") _
            + "_WFRAME.TXT"
            
      Open fnm For Output As #2
      Close #2
      
      End If
   
   fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(ccy, "0000") + "M" + Format(ccm, "00") _
   + "_ASSOSI.TXT"
   
   If Dir(fnm) <> "" Then Kill fnm
   
   End If
      
If BGF = "Y" Then
    
   fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(ccy, "0000") + "M" + Format(ccm, "00") _
   + "_FRAME.TXT"
   
   If Dir(fnm) <> "" Then
   
      fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(ccy, "0000") + "M" + Format(ccm, "00") _
            + "_WFRAME.TXT"
            
      Open fnm For Output As #2
      Close #2
      
      End If
   
   fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(ccy, "0000") + "M" + Format(ccm, "00") _
   + "_ACTIVE.TXT"
   
   If Dir(fnm) <> "" Then
   
      fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(ccy, "0000") + "M" + Format(ccm, "00") _
            + "_WACTIVE.TXT"
            
      Open fnm For Output As #2
      Close #2
      
      End If
   
   
    End If
   
End Sub
Private Sub CHECK_NUM()

Dim fnm, RESP, wmsgccy, ccm, ccy, wmsg, wmsg1, wmsg2, wmsg3, wmsg4
Dim wmsg5, wmsg6, wmsg7, wmsg8

Dim fnm2, fnm3, fnm4, fnm5, fnm6, fnm7, fnm8

ccy = current_year: ccm = current_month

wmsg1 = msgtab(104)

wmsg2 = "": wmsg3 = "": wmsg4 = "": wmsg5 = "": wmsg6 = "": wmsg7 = "": wmsg8 = ""

fnm2 = APPROOT + "\ARTBAS\TABLES\Y" + Format(ccy, "0000") + "M" + Format(ccm, "00") _
   + "_ASSOMN.TXT"

If Dir(fnm2) <> "" Then wmsg2 = msgtab(105)

fnm3 = APPROOT + "\ARTBAS\TABLES\Y" + Format(ccy, "0000") + "M" + Format(ccm, "00") _
   + "_ASSOSI.TXT"

If Dir(fnm3) <> "" Then wmsg3 = msgtab(106)

fnm4 = APPROOT + "\ARTBAS\TABLES\Y" + Format(ccy, "0000") + "M" + Format(ccm, "00") _
   + "_FRAME.TXT"

If Dir(fnm4) <> "" Then wmsg4 = msgtab(107)

fnm5 = APPROOT + "\ARTBAS\TABLES\Y" + Format(ccy, "0000") + "M" + Format(ccm, "00") _
   + "_ACTIVE.TXT"

If Dir(fnm5) <> "" Then wmsg5 = msgtab(108)

fnm6 = APPROOT + "\ARTBAS\EFFORT\Y" + Format(ccy, "0000") + "M" + Format(ccm, "00") _
   + "_*.*"

If Dir(fnm6) <> "" Then wmsg6 = msgtab(109)

fnm7 = APPROOT + "\ARTBAS\LANDINGS\Y" + Format(ccy, "0000") + "M" + Format(ccm, "00") _
   + "_*.*"

If Dir(fnm7) <> "" Then wmsg7 = msgtab(110)

fnm8 = APPROOT + "\ARTBAS\RESULTS\Y" + Format(ccy, "0000") + "M" + Format(ccm, "00") _
   + "_*.*"

If Dir(fnm8) <> "" Then wmsg8 = msgtab(111)

'====================================
'MAJOR STRATA
'====================================

If NUMJ = "Y" Then

   If wmsg2 = "" And wmsg8 = "" Then
      NUMYN = "YES"
      Exit Sub
      End If
   
   wmsg = wmsg1 + Chr(13) + wmsg2 + Chr(13) + wmsg8
   RESP = MsgBox(wmsg, vbCritical + vbOKCancel, " ")
   If RESP = 2 Then
      NUMYN = "NO"
      Exit Sub
      End If
   
   NUMYN = "YES"
    
   If Dir(fnm2) <> "" Then Kill fnm2
   
   If Dir(fnm8) <> "" Then Kill fnm8
   
   Exit Sub

End If
  
'==================
'MINOR STRATA
'==================

If NUMN = "Y" Then

   If wmsg2 = "" And wmsg3 = "" And wmsg5 = "" And wmsg6 = "" And wmsg7 = "" And wmsg8 = "" Then
      NUMYN = "YES"
      Exit Sub
      End If
   
   wmsg = wmsg1 + Chr(13) + wmsg2 + Chr(13) + _
          wmsg3 + Chr(13) + wmsg5 + Chr(13) + wmsg6 + Chr(13) + wmsg7 + Chr(13) + wmsg8
   RESP = MsgBox(wmsg, vbCritical + vbOKCancel, " ")
   If RESP = 2 Then
      NUMYN = "NO"
      Exit Sub
      End If
   
   NUMYN = "YES"
   
   If Dir(fnm2) <> "" Then Kill fnm2
   If Dir(fnm3) <> "" Then Kill fnm3
   If Dir(fnm5) <> "" Then Kill fnm5
   If Dir(fnm6) <> "" Then Kill fnm6
   If Dir(fnm7) <> "" Then Kill fnm7
   If Dir(fnm8) <> "" Then Kill fnm8
   
   Exit Sub

End If
  
'=================
'SITES
'=================

If NUSI = "Y" Then

   If wmsg3 = "" And wmsg4 = "" And wmsg6 = "" And wmsg7 = "" And wmsg8 = "" Then
      NUMYN = "YES"
      Exit Sub
      End If
   
   wmsg = wmsg1 + Chr(13) + wmsg3 + Chr(13) + _
          wmsg4 + Chr(13) + wmsg6 + Chr(13) + wmsg7 + Chr(13) + wmsg8
   RESP = MsgBox(wmsg, vbCritical + vbOKCancel, " ")
   If RESP = 2 Then
      NUMYN = "NO"
      Exit Sub
      End If
   
   NUMYN = "YES"
   
   If Dir(fnm3) <> "" Then Kill fnm3
   If Dir(fnm4) <> "" Then Kill fnm4
   If Dir(fnm6) <> "" Then Kill fnm6
   If Dir(fnm7) <> "" Then Kill fnm7
   If Dir(fnm8) <> "" Then Kill fnm8
   
   Exit Sub

End If
  
'===================
'BOAT/GEAR TYPES
'===================

If NUBG = "Y" Then

   If wmsg4 = "" And wmsg5 = "" And wmsg6 = "" And wmsg7 = "" And wmsg8 = "" Then
      NUMYN = "YES"
      Exit Sub
      End If
   
   wmsg = wmsg1 + Chr(13) + wmsg4 + Chr(13) + _
          wmsg5 + Chr(13) + wmsg6 + Chr(13) + wmsg7 + Chr(13) + wmsg8
   RESP = MsgBox(wmsg, vbCritical + vbOKCancel, " ")
   If RESP = 2 Then
      NUMYN = "NO"
      Exit Sub
      End If
   
   NUMYN = "YES"
   
   If Dir(fnm4) <> "" Then Kill fnm4
   If Dir(fnm5) <> "" Then Kill fnm5
   If Dir(fnm6) <> "" Then Kill fnm6
   If Dir(fnm7) <> "" Then Kill fnm7
   If Dir(fnm8) <> "" Then Kill fnm8
   
   Exit Sub

End If
  
'=========
'SPECIES
'=========

If NUSP = "Y" Then

   If wmsg7 = "" And wmsg8 = "" Then
      NUMYN = "YES"
      Exit Sub
      End If
   
   wmsg = wmsg1 + Chr(13) + wmsg7 + Chr(13) + wmsg8
   RESP = MsgBox(wmsg, vbCritical + vbOKCancel, " ")
   If RESP = 2 Then
      NUMYN = "NO"
      Exit Sub
      End If
   
   NUMYN = "YES"
   
   If Dir(fnm7) <> "" Then Kill fnm7
   If Dir(fnm8) <> "" Then Kill fnm8
   
   Exit Sub

End If
  
End Sub
Private Sub CHECK_NULLS()

Dim fnm, N, XXX

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_MAJOR.TXT"
      
GoSub ERASE_NULLS

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_MINOR.TXT"
      
GoSub ERASE_NULLS

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_SITES.TXT"
      
GoSub ERASE_NULLS

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_BG.TXT"
      
GoSub ERASE_NULLS

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_SPECIES.TXT"
      
GoSub ERASE_NULLS

Exit Sub

ERASE_NULLS:

If Dir(fnm) = "" Then Return

Open fnm For Input As #1

N = 0

Do Until EOF(1)

Line Input #1, XXX

XXX = Mid(XXX, 6, 30): XXX = RTrim(XXX)

If Len(XXX) <> 0 And Left(XXX, 3) <> "???" Then N = N + 1

Loop

Close #1

If N = 0 Then Kill fnm

Return

End Sub

