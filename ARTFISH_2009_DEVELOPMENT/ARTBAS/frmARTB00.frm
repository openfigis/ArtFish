VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form frmARTB00 
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
      Picture         =   "frmARTB00.frx":0000
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
      Picture         =   "frmARTB00.frx":1806
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
      Picture         =   "frmARTB00.frx":3A68
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
      Picture         =   "frmARTB00.frx":3CEA
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
      Picture         =   "frmARTB00.frx":3F6C
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
      Picture         =   "frmARTB00.frx":41EE
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
      Picture         =   "frmARTB00.frx":4470
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
      Visible         =   0   'False
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
      Visible         =   0   'False
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
Attribute VB_Name = "frmARTB00"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private CTAB()

' VVVVVVVVVV Added by TNJ 24/1/08 VVVVVVVVVV
Private Const LOCALE_SMONDECIMALSEP  As Long = &H16  'decimal separator
Private Const LOCALE_SMONTHOUSANDSEP As Long = &H17  'thousand separator

Private Declare Function GetThreadLocale Lib "kernel32" () As Long

Private Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long

Private Declare Function GetUserDefaultLCID Lib "kernel32" () As Long

Private Declare Function GetLocaleInfo Lib "kernel32" _
   Alias "GetLocaleInfoA" _
  (ByVal Locale As Long, _
   ByVal LCType As Long, _
   ByVal lpLCData As String, _
   ByVal cchData As Long) As Long

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

' ^^^^^^^^ Added by TNJ 24/1/08 ^^^^^^^^

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

Open APPROOT + "\ARTBAS\CONTROL\YEARMONTH.TXT" For Output As #1

Write #1, current_year, current_month

Close #1

CTLMONTH = current_month

frmARTB00.MousePointer = 13
cmdENTER.MousePointer = 13

Load frmARTB01

frmARTB00.Hide
' Unload frmARTB00

frmARTB01.Show

End Sub
Private Sub cmdEXIT_Click()

Call CHECK_BACKUP_COMPLETE

cmdEXIT.MousePointer = 13
frmARTB00.MousePointer = 13

Call write_parms

Call KILL_ARTBASIC_FOLDER

End

End Sub

Private Sub cmdGUIDE_Click()

HTYPE = "10"

HFNM = APPROOT + "\ARTBAS\HELP\" + current_language + "HELP" + HTYPE + ".rtf"

If Dir(HFNM) = "" Then Exit Sub

frmARTB00.Enabled = False
Load frmGUIDE
frmGUIDE.Show

End Sub
Private Sub Form_Load()

Dim XXX, LLL

XXX = CurDir()

LLL = InStr(XXX, "\ARTBAS")

' TNJ Note: If LLL is 0 (zero), the directory being executed from
' is not "\ARTBAS"
If LLL = 0 Then
   MsgBox "The directory where the ArtBasic executable is being run is not '\ARTBAS' !!!"
   MsgBox "The directory is: " + XXX
   End
End If

APPROOT = GetShortName(Left(XXX, LLL - 1))

Label2.Visible = False
Label3.Visible = False

Set Picture = LoadPicture(APPROOT + "\ARTBAS\PICS_RUNTIME\SCREEN_01.JPG")

Call CHECK_ADMIN

cmdENTER.Enabled = False
current_month = 999

Call READ_CONTENTS

frmARTB00.MousePointer = 1

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

frmARTB00.Caption = msgtab(4)
' Label2.Caption = msgtab(249)
' Label3.Caption = msgtab(250)

lblYEAR.Caption = current_year

Call DISPLAY_MONTHS

cmdENGLISH.ToolTipText = msgtab_e
cmdFRENCH.ToolTipText = msgtab_f
cmdSPANISH.ToolTipText = msgtab_s
cmdLOCAL.ToolTipText = msgtab_l
cmdENTER.ToolTipText = msgtab(2)
cmdEXIT.ToolTipText = msgtab(3)
cmdGUIDE.ToolTipText = msgtab(243)

' Added - 14/1/08 by TNJ - writes data to the registry to ensure that
' Regional Settings are not changed since the first execution of the program
SaveSetting "ArtBasic", "ARTB00", "StartUp", Me.hwnd

' Previously, the program wrote to the C: Drive the file: C:\ARTBASIC_CURRENT_FOLDER.TXT
' If the file was there, that meant that there was already an ArtBasic instance running
' Since writing to the C: drive is not allowed at FAO, the program was changed to
' create an exclusive 'Semaphore' in memory.  This ensures that only 1 instance is active.
' Dim NORUN, NORUNMSG

' NORUNMSG = msgtab(252) + "  " + msgtab(253) + "  " + msgtab(254) + _
           "  " + msgtab(255) + "  " + msgtab(256)

'If Dir("C:\ARTBASIC_CURRENT_FOLDER.TXT") <> "" Then
'   NORUN = MsgBox(NORUNMSG, vbOKCancel, "  ")
'   End
'   End If

' Previously, the program wrote to the C: drive to ensure that only 1 instance
' of the program was running - this is now accomplished by creating a mutually
' exclusive sempahore.
' Call APPROOT_WRITE

' Determine where Excel is installed on the system
Call LOCATE_EXCEL

Dim LCID As Long
Dim curDecimal, curSep1000, rtnValue As String

   ' If Option1.Value = True Then
   '    LCID = GetSystemDefaultLCID()
   ' Else
   ' TNJ - currently, get the user default that should be the same as the system
      LCID = GetUserDefaultLCID()
   ' End If

'VERY IMPORTANT!!!  VERY IMPORTANT!!!  VERY IMPORTANT!!!  VERY IMPORTANT!!!
'Insure that Regional Setting had not been changed since the first program execution
'If using different regional settings for data entry, the data may be very corrupted

'LOCALE_SMONDECIMALSEP
'Check the character(s) used as the  decimal separator.
'The maximum characters allowed is four.
rtnValue = GetSetting("ArtBasic", "ARTB00", "Decimal", "New")
curDecimal = GetCurrentLocaleInfo(LCID, LOCALE_SMONDECIMALSEP)
If rtnValue <> "New" Then
   If rtnValue <> curDecimal Then
      MsgBox "Decimal separator has been changed!!!"
      Unload Me
      End
   End If
Else
   'SaveSetting stores the data in the Registry
   '(in HKEY_CURRENT_USER|Software|VB and VBA Program Settings|YourAppName).
   SaveSetting "ArtBasic", "ARTB00", "Decimal", curDecimal
End If

'LOCALE_SMONTHOUSANDSEP
'Check the character(s) used as the  separator between groups of
'digits to the left of the decimal.
'The maximum number of characters allowed is four.
rtnValue = GetSetting("ArtBasic", "ARTB00", "Sep1000", "New")
curSep1000 = GetCurrentLocaleInfo(LCID, LOCALE_SMONTHOUSANDSEP)
If rtnValue <> "New" Then
   If rtnValue <> curSep1000 Then
      MsgBox "Decimal group separator has been changed!!!"
      Unload Me
      End
   End If
Else
   SaveSetting "ArtBasic", "ARTB00", "Sep1000", curSep1000
End If

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
Private Sub cmdLOCAL_Click()
language = "L": current_language = language         ' Local language
chkENGLISH.Visible = False
chkFRENCH.Visible = False
chkSPANISH.Visible = False
chkLOCAL.Visible = True
chkLOCAL.Value = 1
Call MSGLOAD
cmdGUIDE.ToolTipText = msgtab(243)
cmdENTER.ToolTipText = msgtab(2)
cmdEXIT.ToolTipText = msgtab(3)
frmARTB00.Caption = msgtab(4)
Label2.Caption = msgtab(249)
Label3.Caption = msgtab(250)
Call DISPLAY_MONTHS
End Sub
Private Sub cmdENGLISH_Click()
language = "E": current_language = language         ' English
chkENGLISH.Visible = True
chkENGLISH.Value = 1
chkFRENCH.Visible = False
chkSPANISH.Visible = False
chkLOCAL.Visible = False
Call MSGLOAD
cmdGUIDE.ToolTipText = msgtab(243)
cmdENTER.ToolTipText = msgtab(2)
cmdEXIT.ToolTipText = msgtab(3)
frmARTB00.Caption = msgtab(4)
Label2.Caption = msgtab(249)
Label3.Caption = msgtab(250)
Call DISPLAY_MONTHS
End Sub
Private Sub cmdFRENCH_Click()
language = "F": current_language = language         'French
chkENGLISH.Visible = False
chkFRENCH.Visible = True
chkFRENCH.Value = 1
chkSPANISH.Visible = False
chkLOCAL.Visible = False
Call MSGLOAD
cmdGUIDE.ToolTipText = msgtab(243)
cmdENTER.ToolTipText = msgtab(2)
cmdEXIT.ToolTipText = msgtab(3)
frmARTB00.Caption = msgtab(4)
Label2.Caption = msgtab(249)
Label3.Caption = msgtab(250)
Call DISPLAY_MONTHS
End Sub
Private Sub cmdSPANISH_Click()
language = "S": current_language = language         ' Spanish
chkENGLISH.Visible = False
chkFRENCH.Visible = False
chkSPANISH.Visible = True
chkSPANISH.Value = 1
chkLOCAL.Visible = False
Call MSGLOAD
cmdGUIDE.ToolTipText = msgtab(243)
cmdENTER.ToolTipText = msgtab(2)
cmdEXIT.ToolTipText = msgtab(3)
frmARTB00.Caption = msgtab(4)
Label2.Caption = msgtab(249)
Label3.Caption = msgtab(250)
Call DISPLAY_MONTHS
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseMutex Mutex
    CloseHandle Mutex
    DeleteSetting "ArtBasic", "ARTB00", "StartUp"
End Sub

Private Sub lblDATA_Click(Index As Integer)

Dim I, J

I = Index

If I + 1 = current_month Then
   lblMONTHS(I).BackColor = vbWhite
   lblDATA(I).BackColor = vbWhite
   current_month = 999
   cmdENTER.Enabled = False
   Exit Sub
   End If

If I + 1 <> current_month And current_month = 999 Then
   current_month = I + 1
   lblMONTHS(I).BackColor = vbYellow
   lblDATA(I).BackColor = vbYellow
   cmdENTER.Enabled = True
   Exit Sub
   End If

If I + 1 <> current_month And current_month <> 999 Then
   J = current_month - 1
   lblMONTHS(J).BackColor = vbWhite
   lblDATA(J).BackColor = vbWhite
   current_month = I + 1
   lblMONTHS(I).BackColor = vbYellow
   lblDATA(I).BackColor = vbYellow
   cmdENTER.Enabled = True
   Exit Sub
   End If
End Sub
Private Sub lblMONTHS_Click(Index As Integer)

Dim I, J

I = Index

If I + 1 = current_month Then
   lblMONTHS(I).BackColor = vbWhite
   lblDATA(I).BackColor = vbWhite
   current_month = 999
   cmdENTER.Enabled = False
   Exit Sub
   End If

If I + 1 <> current_month And current_month = 999 Then
   current_month = I + 1
   lblMONTHS(I).BackColor = vbYellow
   lblDATA(I).BackColor = vbYellow
   cmdENTER.Enabled = True
   Exit Sub
   End If

If I + 1 <> current_month And current_month <> 999 Then
   J = current_month - 1
   lblMONTHS(J).BackColor = vbWhite
   lblDATA(J).BackColor = vbWhite
   current_month = I + 1
   lblMONTHS(I).BackColor = vbYellow
   lblDATA(I).BackColor = vbYellow
   cmdENTER.Enabled = True
   Exit Sub
   End If

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
Private Sub LOCATE_EXCEL()

' VVVVVVVV Commented out by TNJ 24/1/08 VVVVVVVV
' Previously, the program tried to locate the Excel application like this
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
    MsgBox ("!!!!!!!!!!!!!! Important - The Excel application was not located by ArtBasic !!!!!!!!!!!!!")
    MsgBox ("Batch file 'GENERAL.BAT' in the EXCEL_REPORTS subdirectory must be edited manually to access the Excel application")
    MsgBox ("Note: The path for the Excel application can NOT have spaces.  You must use 'short' format")
    End If

EXCEL_OK:

' Write out the General.Bat file for our installation location
Open APPROOT + "\ARTBAS\EXCEL_REPORTS\GENERAL.BAT" For Output As #1

' We write the Excel call with passing the parameter _AR: that gives the
' AppRoot (Application Root) to tell where we are executing the application
Print #1, "ECHO OFF"
Print #1, GetShortName(LTrim(PATH_EXCEL)) & "\EXCEL.EXE /_AR:" _
    & APPROOT & " " _
    & APPROOT & "\ARTBAS\EXCEL_REPORTS\ARTFISH_WORK.XLS"

Close #1

' Write out the Import.Bat file for our installation location
Open APPROOT + "\ARTBAS\EXPORT\IMPORT.BAT" For Output As #1

Print #1, "ECHO OFF"
Print #1, "COPY " & APPROOT & "\ARTBAS\ARTBASIC_RESULTS\*.* " & APPROOT & "\ARTBAS\TRANSFER"

Close #1

' VVVVVVVV Added by TNJ 24/1/08 but commented out on 15/6/09 VVVVVVVV
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

   ' First, get the clsid from the progid from the registry key:
   ' HKEY_LOCAL_MACHINE\Software\Classes\<PROGID>\CLSID
   ' RetVal = RegOpenKeyEx(HKEY_LOCAL_MACHINE, "Software\Classes\" & _
   '    sProgId & "\CLSID", 0&, KEY_ALL_ACCESS, hKey)
   ' If RetVal = 0 Then
   '    Dim N As Long
   '    RetVal = RegQueryValueEx(hKey, "", 0&, REG_SZ, "", N)
   '    sCLSID = Space(N)
   '    RetVal = RegQueryValueEx(hKey, "", 0&, REG_SZ, sCLSID, N)
   '    sCLSID = Left(sCLSID, N - 1)  'drop null-terminator
   '    RegCloseKey hKey
   ' Else
      ' We have no Local Administrator rights on this PC - the Excel application
      ' location must be determined manually
   '    MsgBox ("This PC does not have Local Administrator rights")
   '    MsgBox ("Batch file 'GENERAL.BAT' in the EXCEL_REPORTS subdirectory must be edited manually to access the Excel application")
   '    End
   ' End If
      
   'Now that we have the CLSID, locate the server path at
   'HKEY_LOCAL_MACHINE\Software\Classes\CLSID\
   '     {xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxx}\LocalServer32

   ' RetVal = RegOpenKeyEx(HKEY_LOCAL_MACHINE, _
   '      "Software\Classes\CLSID\" & sCLSID & "\LocalServer32", 0&, _
   '    KEY_ALL_ACCESS, hKey)
   ' If RetVal = 0 Then
   '    RetVal = RegQueryValueEx(hKey, "", 0&, REG_SZ, "", N)
   '    sPath = Space(N)

   '    RetVal = RegQueryValueEx(hKey, "", 0&, REG_SZ, sPath, N)
   '    sPath = Left(sPath, N - 1)
   '    MsgBox sPath
   '    RegCloseKey hKey
   ' End If

   ' Now - only take first part of the path
   ' N = InStr(sPath, "EXCEL.EXE")
   ' sPath = Left(sPath, N - 2)

   ' Open APPROOT + "\ARTBAS\EXCEL_REPORTS\GENERAL.BAT" For Output As #1

   ' Print #1, "ECHO OFF"
   ' Print #1, "CD\"
   ' Print #1, "CD " & sPath
   ' Print #1, "EXCEL.EXE C:\ARTFISH_WORK.XLS"

   ' Close #1

' ^^^^^^^^ Added by TNJ 24/1/08 but commented out on 15/6/09 ^^^^^^^^

End Sub

Private Function GetCurrentLocaleInfo(ByVal dwLocaleID As Long, ByVal dwLCType As Long) As String

   Dim sReturn As String
   Dim r As Long

  'call the function passing the Locale type
  'variable to retrieve the required size of
  'the string buffer needed
   r = GetLocaleInfo(dwLocaleID, dwLCType, sReturn, Len(sReturn))
    
  'if successful..
   If r Then
    
     'pad the buffer with spaces
      sReturn = Space$(r)
       
     'and call again passing the buffer
      r = GetLocaleInfo(dwLocaleID, dwLCType, sReturn, Len(sReturn))
     
     'if successful (r > 0)
      If r Then
      
        'r holds the size of the string
        'including the terminating null
         GetCurrentLocaleInfo = Left$(sReturn, r - 1)
      
      End If
   
   End If
    
End Function

' ^^^^^^^^ Added by TNJ 24/1/08 ^^^^^^^^
