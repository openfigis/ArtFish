VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form frmDB00 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ARTBASIC"
   ClientHeight    =   7530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10545
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
   ScaleHeight     =   7530
   ScaleWidth      =   10545
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdLOCAL 
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
      Height          =   495
      Left            =   9720
      MousePointer    =   1  'Arrow
      Picture         =   "frmDB00.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   6000
      Width           =   615
   End
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
      Picture         =   "frmDB00.frx":1806
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   6720
      Width           =   735
   End
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
      Left            =   9600
      MousePointer    =   1  'Arrow
      Picture         =   "frmDB00.frx":3A68
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   6720
      Width           =   735
   End
   Begin VB.CommandButton cmdENTER 
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
      Left            =   7560
      MousePointer    =   1  'Arrow
      Picture         =   "frmDB00.frx":3CEA
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   6720
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
      Left            =   9960
      TabIndex        =   33
      Top             =   6480
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
      Left            =   9240
      TabIndex        =   32
      Top             =   6480
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
      Left            =   8520
      TabIndex        =   31
      Top             =   6480
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
      TabIndex        =   30
      Top             =   6480
      Value           =   1  'Checked
      Width           =   135
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
      Left            =   9000
      MousePointer    =   1  'Arrow
      Picture         =   "frmDB00.frx":3F6C
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   6000
      Width           =   615
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
      Picture         =   "frmDB00.frx":41EE
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   6000
      Width           =   615
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
      Left            =   7560
      MousePointer    =   1  'Arrow
      Picture         =   "frmDB00.frx":4470
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   6000
      Width           =   615
   End
   Begin ComCtl2.UpDown updYEAR 
      Height          =   495
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   873
      _Version        =   327681
      Value           =   2000
      OrigLeft        =   840
      OrigTop         =   3000
      OrigRight       =   1080
      OrigBottom      =   3375
      Max             =   2050
      Min             =   1990
      Enabled         =   -1  'True
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "label3"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   3240
      TabIndex        =   39
      Top             =   720
      Width           =   4695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "label2"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   1560
      TabIndex        =   38
      Top             =   120
      Width           =   8415
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FFFF&
      Caption         =   " 01"
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
      TabIndex        =   36
      Top             =   7320
      Width           =   255
   End
   Begin VB.Label lblSEL 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "???"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   240
      TabIndex        =   26
      Top             =   6000
      Width           =   420
   End
   Begin VB.Label lblDATA 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   11
      Left            =   2280
      TabIndex        =   25
      Top             =   5280
      Width           =   255
   End
   Begin VB.Label lblDATA 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   10
      Left            =   2280
      TabIndex        =   24
      Top             =   4920
      Width           =   255
   End
   Begin VB.Label lblDATA 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   9
      Left            =   2280
      TabIndex        =   23
      Top             =   4560
      Width           =   255
   End
   Begin VB.Label lblDATA 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   8
      Left            =   2280
      TabIndex        =   22
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label lblDATA 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   7
      Left            =   2280
      TabIndex        =   21
      Top             =   3840
      Width           =   255
   End
   Begin VB.Label lblDATA 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   6
      Left            =   2280
      TabIndex        =   20
      Top             =   3480
      Width           =   255
   End
   Begin VB.Label lblDATA 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   5
      Left            =   2280
      TabIndex        =   19
      Top             =   3120
      Width           =   255
   End
   Begin VB.Label lblDATA 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   4
      Left            =   2280
      TabIndex        =   18
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label lblDATA 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   3
      Left            =   2280
      TabIndex        =   17
      Top             =   2400
      Width           =   255
   End
   Begin VB.Label lblDATA 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   2
      Left            =   2280
      TabIndex        =   16
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label lblDATA 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   1
      Left            =   2280
      TabIndex        =   15
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label lblMONTHS 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   11
      Left            =   240
      TabIndex        =   14
      Top             =   5280
      Width           =   1935
   End
   Begin VB.Label lblMONTHS 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   10
      Left            =   240
      TabIndex        =   13
      Top             =   4920
      Width           =   1935
   End
   Begin VB.Label lblMONTHS 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   9
      Left            =   240
      TabIndex        =   12
      Top             =   4560
      Width           =   1935
   End
   Begin VB.Label lblMONTHS 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   8
      Left            =   240
      TabIndex        =   11
      Top             =   4200
      Width           =   1935
   End
   Begin VB.Label lblMONTHS 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   7
      Left            =   240
      TabIndex        =   10
      Top             =   3840
      Width           =   1935
   End
   Begin VB.Label lblMONTHS 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   6
      Left            =   240
      TabIndex        =   9
      Top             =   3480
      Width           =   1935
   End
   Begin VB.Label lblMONTHS 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   5
      Left            =   240
      TabIndex        =   8
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Label lblMONTHS 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   4
      Left            =   240
      TabIndex        =   7
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Label lblMONTHS 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   3
      Left            =   240
      TabIndex        =   6
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Label lblMONTHS 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   2
      Left            =   240
      TabIndex        =   5
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label lblMONTHS 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   1
      Left            =   240
      TabIndex        =   4
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label lblDATA 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   0
      Left            =   2280
      TabIndex        =   3
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label lblMONTHS 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label lblYEAR 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmDB00"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

Dim XXX, LLL

XXX = CurDir()

LLL = InStr(XXX, "\ARTBAS")

APPROOT = GetShortName(Left(XXX, LLL - 1))

Call CHECK_ADMIN

cmdENTER.Enabled = False
current_month = 999

Call READ_CONTENTS

frmDB00.MousePointer = 1

Call read_parms

language = current_language

If language = "E" Then        ' English
    chkENGLISH.Visible = True
    chkENGLISH.Value = 1
    chkFRENCH.Visible = False
    chkSPANISH.Visible = False
    chkLOCAL.Visible = False
End If

If language = "F" Then         ' French
    chkENGLISH.Visible = False
    chkFRENCH.Value = 1
    chkFRENCH.Visible = True
    chkSPANISH.Visible = False
    chkLOCAL.Visible = False
End If

If language = "S" Then        ' Spanish
    chkENGLISH.Visible = False
    chkSPANISH.Value = 1
    chkFRENCH.Visible = False
    chkSPANISH.Visible = True
    chkLOCAL.Visible = False
End If

If language = "L" Then          ' Local language
    chkENGLISH.Visible = False
    chkLOCAL.Value = 1
    chkFRENCH.Visible = False
    chkSPANISH.Visible = False
    chkLOCAL.Visible = True
End If

Call MSGLOAD

frmDB00.Caption = msgtab(4)
Label2.Caption = msgtab(249)
Label3.Caption = msgtab(250)

lblYEAR.Caption = current_year

Call DISPLAY_MONTHS

CTLMONTH = current_month

Load frmDB01
Unload frmDB00

frmDB01.Show
frmDB01.Refresh

End Sub
Private Sub read_parms()

CTLADMIN = "NO"

If Dir(APPROOT + "\ARTBAS\CONTROL\ADMIN.TXT") <> "" Then
   Open APPROOT + "\ARTBAS\CONTROL\ADMIN.TXT" For Input As #1
   Input #1, CTLADMIN
   CTLADMIN = RTrim(CTLADMIN)
   Close #1
   End If

Dim textline As String * 80, ll As Integer
Open APPROOT + "\ARTBAS\CONTROL\SYSPARM.TXT" For Input As #1

Input #1, current_language

Input #1, current_year

updYEAR.Value = current_year

Close #1

End Sub
Private Sub updYEAR_Change()
lblYEAR.Caption = updYEAR.Value
current_year = updYEAR.Value
Call DISPLAY_MONTHS
End Sub
Private Sub DISPLAY_MONTHS()

language = current_language
Call MSGLOAD

lblSEL.Caption = msgtab(5)

Dim I, J

For I = 0 To 11

lblMONTHS(I).Caption = msgtab(18 + I)

Next I

I = updYEAR.Value - 1989

For J = 0 To 11
lblDATA(J).Caption = CY(I, J + 1)
Next J

End Sub
Private Sub READ_CONTENTS()

Open APPROOT + "\ARTBAS\CONTROL\CONTENTS.TXT" For Input As #1

Dim xy, I, J, XCHAR, ZZZ, fnm
Dim XC As String * 12

ReDim CY(1 To 100, 1 To 12)

For I = 1 To 100
For J = 1 To 12
CY(I, J) = " "
Next J
Next I

Do Until EOF(1)

Line Input #1, ZZZ

xy = Val(Left(ZZZ, 4)): XC = Mid(ZZZ, 6, 31)

For J = 1 To 12

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(xy, "0000") + "M" + Format(J, "00") + "_*.*"

CY(xy - 1989, J) = " "

If Dir(fnm) <> "" Then CY(xy - 1989, J) = "X"

Next J

Loop

Close #1

End Sub
Private Sub CHECK_ADMIN()

Dim fnm, YN

fnm = APPROOT + "\ARTBAS\CONTROL\ADMIN.TXT"

Open fnm For Input As #1

Input #1, YN
Close #1

Open fnm For Output As #1

Print #1, UCase(YN)
Close #1

End Sub


