VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.0#0"; "COMCT232.OCX"
Begin VB.Form frmARTS00 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ARTBASIC"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8910
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
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
   Picture         =   "frmARTS00.frx":0000
   ScaleHeight     =   5655
   ScaleWidth      =   8910
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
      Height          =   615
      Left            =   7800
      MousePointer    =   1  'Arrow
      Picture         =   "frmARTS00.frx":3966
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   3120
      Width           =   615
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
      Height          =   615
      Left            =   4920
      MousePointer    =   1  'Arrow
      Picture         =   "frmARTS00.frx":3BE8
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   3120
      Width           =   615
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
      Left            =   8040
      TabIndex        =   36
      Top             =   2760
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
      Left            =   7080
      TabIndex        =   35
      Top             =   2760
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
      Left            =   6120
      TabIndex        =   34
      Top             =   2760
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
      Left            =   5160
      TabIndex        =   33
      Top             =   2760
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
      Left            =   7800
      MousePointer    =   1  'Arrow
      Picture         =   "frmARTS00.frx":3E6A
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   2280
      Width           =   615
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
      Left            =   6840
      MousePointer    =   1  'Arrow
      Picture         =   "frmARTS00.frx":40EC
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   2280
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
      Left            =   5880
      MousePointer    =   1  'Arrow
      Picture         =   "frmARTS00.frx":436E
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   2280
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
      Left            =   4920
      MousePointer    =   1  'Arrow
      Picture         =   "frmARTS00.frx":45F0
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   2280
      Width           =   615
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3795
      Left            =   4680
      Picture         =   "frmARTS00.frx":4872
      ScaleHeight     =   3795
      ScaleWidth      =   3975
      TabIndex        =   28
      Top             =   240
      Width           =   3975
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
         Left            =   1680
         MousePointer    =   1  'Arrow
         Picture         =   "frmARTS00.frx":6E3E
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   2760
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Version 2000 - Copyright FAO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   1680
         Width           =   3735
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "By C. Stamatopoulos and T. Jarrett"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   1320
         Width           =   3735
      End
   End
   Begin VB.PictureBox picFAO 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   4680
      Picture         =   "frmARTS00.frx":90A0
      ScaleHeight     =   705
      ScaleWidth      =   3945
      TabIndex        =   0
      Top             =   4440
      Width           =   3975
   End
   Begin ComCtl2.UpDown updYEAR 
      Height          =   495
      Left            =   1200
      TabIndex        =   1
      Top             =   240
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   873
      _Version        =   327680
      Value           =   2000
      OrigLeft        =   840
      OrigTop         =   3000
      OrigRight       =   1080
      OrigBottom      =   3375
      Max             =   2050
      Min             =   1990
      Enabled         =   -1  'True
   End
   Begin VB.Label lblSEL 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "???"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   27
      Top             =   5280
      Width           =   5415
   End
   Begin VB.Label lblDATA 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   11
      Left            =   2280
      TabIndex        =   26
      Top             =   4800
      Width           =   255
   End
   Begin VB.Label lblDATA 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   10
      Left            =   2280
      TabIndex        =   25
      Top             =   4440
      Width           =   255
   End
   Begin VB.Label lblDATA 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   9
      Left            =   2280
      TabIndex        =   24
      Top             =   4080
      Width           =   255
   End
   Begin VB.Label lblDATA 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   8
      Left            =   2280
      TabIndex        =   23
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label lblDATA 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   7
      Left            =   2280
      TabIndex        =   22
      Top             =   3360
      Width           =   255
   End
   Begin VB.Label lblDATA 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   6
      Left            =   2280
      TabIndex        =   21
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label lblDATA 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   5
      Left            =   2280
      TabIndex        =   20
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label lblDATA 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   4
      Left            =   2280
      TabIndex        =   19
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label lblDATA 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   3
      Left            =   2280
      TabIndex        =   18
      Top             =   1920
      Width           =   255
   End
   Begin VB.Label lblDATA 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   2
      Left            =   2280
      TabIndex        =   17
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label lblDATA 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   2280
      TabIndex        =   16
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label lblMONTHS 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   11
      Left            =   240
      TabIndex        =   15
      Top             =   4800
      Width           =   1935
   End
   Begin VB.Label lblMONTHS 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   10
      Left            =   240
      TabIndex        =   14
      Top             =   4440
      Width           =   1935
   End
   Begin VB.Label lblMONTHS 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   9
      Left            =   240
      TabIndex        =   13
      Top             =   4080
      Width           =   1935
   End
   Begin VB.Label lblMONTHS 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   8
      Left            =   240
      TabIndex        =   12
      Top             =   3720
      Width           =   1935
   End
   Begin VB.Label lblMONTHS 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   7
      Left            =   240
      TabIndex        =   11
      Top             =   3360
      Width           =   1935
   End
   Begin VB.Label lblMONTHS 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   6
      Left            =   240
      TabIndex        =   10
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Label lblMONTHS 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   5
      Left            =   240
      TabIndex        =   9
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Label lblMONTHS 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   4
      Left            =   240
      TabIndex        =   8
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Label lblMONTHS 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   3
      Left            =   240
      TabIndex        =   7
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label lblMONTHS 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   2
      Left            =   240
      TabIndex        =   6
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label lblMONTHS 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label lblDATA 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   2280
      TabIndex        =   4
      Top             =   840
      Width           =   255
   End
   Begin VB.Label lblMONTHS 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label lblYEAR 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
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
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "frmARTS00"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private CTAB()
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

CTLMONTH = current_month

frmARTB00.MousePointer = 13
cmdENTER.MousePointer = 13
Load frmARTB01
Unload frmARTB00
frmARTB01.Show

End Sub
Private Sub cmdEXIT_Click()

Call CHECK_BACKUP_COMPLETE

cmdEXIT.MousePointer = 13
frmARTB00.MousePointer = 13

Call write_parms

End
End Sub

Private Sub cmdGUIDE_Click()

HTYPE = "10"

HFNM = "C:\ARTBAS\HELP\HELP" + HTYPE + Left(current_language, 1) + ".rtf"

If Dir(HFNM) = "" Then Exit Sub

frmARTB00.Enabled = False
Load frmGUIDE
frmGUIDE.Show

End Sub

Private Sub Form_Load()

cmdENTER.Enabled = False
current_month = 999

Call READ_CONTENTS

frmARTB00.MousePointer = 1

Call read_parms

language = current_language

If language = "ENGLISH" Then

chkENGLISH.Visible = True
chkENGLISH.Value = 1
chkFRENCH.Visible = False
chkSPANISH.Visible = False
chkLOCAL.Visible = False

End If

If language = "FRENCH" Then

chkENGLISH.Visible = False
chkFRENCH.Value = 1
chkFRENCH.Visible = True
chkSPANISH.Visible = False
chkLOCAL.Visible = False

End If

If language = "SPANISH" Then

chkENGLISH.Visible = False
chkSPANISH.Value = 1
chkFRENCH.Visible = False
chkSPANISH.Visible = True
chkLOCAL.Visible = False

End If

If language = "LOCAL" Then

chkENGLISH.Visible = False
chkLOCAL.Value = 1
chkFRENCH.Visible = False
chkSPANISH.Visible = False
chkLOCAL.Visible = True

End If

Call MSGLOAD

frmARTB00.Caption = "ÄRTBASIC: " + msgtab(4)

lblYEAR.Caption = current_year

Call DISPLAY_MONTHS

cmdENGLISH.ToolTipText = msgtab_e
cmdFRENCH.ToolTipText = msgtab_f
cmdSPANISH.ToolTipText = msgtab_s
cmdLOCAL.ToolTipText = msgtab_l
cmdENTER.ToolTipText = msgtab(2)
cmdEXIT.ToolTipText = msgtab(3)

End Sub
Private Sub read_parms()

CTLADMIN = "NO"

If Dir("C:\ARTBAS\CONTROL\ADMIN.TXT") <> "" Then
   Open "C:\ARTBAS\CONTROL\ADMIN.TXT" For Input As #1
   Input #1, CTLADMIN
   CTLADMIN = RTrim(CTLADMIN)
   Close #1
   End If

Dim textline As String * 80, ll As Integer
Open "C:\ARTBAS\CONTROL\SYSPARM.TXT" For Input As #1

Input #1, current_language

Input #1, current_year

updYEAR.Value = current_year

Close #1

End Sub
Private Sub cmdLOCAL_Click()
language = "LOCAL": current_language = language
chkENGLISH.Visible = False
chkFRENCH.Visible = False
chkSPANISH.Visible = False
chkLOCAL.Visible = True
chkLOCAL.Value = 1
Call MSGLOAD
cmdGUIDE.ToolTipText = msgtab(243)
cmdENTER.ToolTipText = msgtab(2)
cmdEXIT.ToolTipText = msgtab(3)
frmARTB00.Caption = "ÄRTBASIC: " + msgtab(4)
Call DISPLAY_MONTHS
End Sub
Private Sub cmdENGLISH_Click()
language = "ENGLISH": current_language = language
chkENGLISH.Visible = True
chkENGLISH.Value = 1
chkFRENCH.Visible = False
chkSPANISH.Visible = False
chkLOCAL.Visible = False
Call MSGLOAD
cmdGUIDE.ToolTipText = msgtab(243)
cmdENTER.ToolTipText = msgtab(2)
cmdEXIT.ToolTipText = msgtab(3)
frmARTB00.Caption = "ARTBASIC: " + msgtab(4)
Call DISPLAY_MONTHS
End Sub
Private Sub cmdFRENCH_Click()
language = "FRENCH": current_language = language
chkENGLISH.Visible = False
chkFRENCH.Visible = True
chkFRENCH.Value = 1
chkSPANISH.Visible = False
chkLOCAL.Visible = False
Call MSGLOAD
cmdGUIDE.ToolTipText = msgtab(243)
cmdENTER.ToolTipText = msgtab(2)
cmdEXIT.ToolTipText = msgtab(3)
frmARTB00.Caption = "ARTBASIC: " + msgtab(4)
Call DISPLAY_MONTHS
End Sub
Private Sub cmdSPANISH_Click()
language = "SPANISH": current_language = language
chkENGLISH.Visible = False
chkFRENCH.Visible = False
chkSPANISH.Visible = True
chkSPANISH.Value = 1
chkLOCAL.Visible = False
Call MSGLOAD
cmdGUIDE.ToolTipText = msgtab(243)
cmdENTER.ToolTipText = msgtab(2)
cmdEXIT.ToolTipText = msgtab(3)
frmARTB00.Caption = "ARTBASIC: " + msgtab(4)
Call DISPLAY_MONTHS
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

Open "C:\ARTBAS\CONTROL\CONTENTS.TXT" For Input As #1

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

fnm = "C:\ARTBAS\TABLES\Y" + Format(xy, "0000") + "M" + Format(J, "00") + "_*.*"

CY(xy - 1989, J) = " "

If Dir(fnm) <> "" Then CY(xy - 1989, J) = "X"

Next J

Loop

Close #1

End Sub
