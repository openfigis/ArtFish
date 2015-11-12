VERSION 5.00
Begin VB.Form frmREPORTS 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   8820
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13275
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   1  'Arrow
   Moveable        =   0   'False
   ScaleHeight     =   8820
   ScaleWidth      =   13275
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optSTD 
      BackColor       =   &H00808000&
      Caption         =   "Option1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   133
      Top             =   8040
      Width           =   5055
   End
   Begin VB.CommandButton cmdEXPORT 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   12360
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmREPORTS.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   131
      Top             =   5160
      Width           =   735
   End
   Begin VB.CommandButton cmdGUIDE 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   12360
      MousePointer    =   1  'Arrow
      Picture         =   "frmREPORTS.frx":1326
      Style           =   1  'Graphical
      TabIndex        =   95
      Top             =   6840
      Width           =   735
   End
   Begin VB.CommandButton cmdPLOT 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   12360
      Picture         =   "frmREPORTS.frx":3588
      Style           =   1  'Graphical
      TabIndex        =   94
      Top             =   6000
      Width           =   735
   End
   Begin VB.ListBox lstDATA 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2370
      Left            =   240
      TabIndex        =   91
      Top             =   240
      Width           =   11895
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   5415
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   3000
      Width           =   11895
      Begin VB.OptionButton optLOC 
         BackColor       =   &H00808000&
         Caption         =   "Option1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   132
         Top             =   4800
         Width           =   5055
      End
      Begin VB.Label lblVALUES 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   500
         Left            =   5280
         TabIndex        =   130
         Top             =   4800
         Width           =   2895
      End
      Begin VB.Label lblSIZE 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   500
         Left            =   8640
         TabIndex        =   129
         Top             =   4800
         Width           =   2655
      End
      Begin VB.Label lblRANK1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         Index           =   1
         Left            =   3840
         TabIndex        =   126
         Top             =   -120
         Width           =   3375
      End
      Begin VB.Label lblRANKC 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   3840
         TabIndex        =   125
         Top             =   120
         Width           =   4815
      End
      Begin VB.Label lblRANK2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   3840
         TabIndex        =   124
         Top             =   360
         Width           =   4815
      End
      Begin VB.Label lblF 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   9480
         TabIndex        =   123
         Top             =   4440
         Width           =   1995
      End
      Begin VB.Label lblF 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   9480
         TabIndex        =   122
         Top             =   4080
         Width           =   1995
      End
      Begin VB.Label lblF 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   9480
         TabIndex        =   121
         Top             =   3840
         Width           =   1995
      End
      Begin VB.Label lblF 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   9480
         TabIndex        =   120
         Top             =   3600
         Width           =   1995
      End
      Begin VB.Label lblF 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   9480
         TabIndex        =   119
         Top             =   3360
         Width           =   1995
      End
      Begin VB.Label lblF 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   9480
         TabIndex        =   118
         Top             =   3120
         Width           =   1995
      End
      Begin VB.Label lblF 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   9480
         TabIndex        =   117
         Top             =   2880
         Width           =   1995
      End
      Begin VB.Label lblF 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   9480
         TabIndex        =   116
         Top             =   2640
         Width           =   1995
      End
      Begin VB.Label lblF 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   9480
         TabIndex        =   115
         Top             =   2400
         Width           =   1995
      End
      Begin VB.Label lblF 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   9480
         TabIndex        =   114
         Top             =   2160
         Width           =   1995
      End
      Begin VB.Label lblF 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   9480
         TabIndex        =   113
         Top             =   1920
         Width           =   1995
      End
      Begin VB.Label lblF 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   9480
         TabIndex        =   112
         Top             =   1680
         Width           =   1995
      End
      Begin VB.Label lblF 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   9480
         TabIndex        =   111
         Top             =   1440
         Width           =   1995
      End
      Begin VB.Label lblTF 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   9480
         TabIndex        =   110
         Top             =   960
         Width           =   1995
      End
      Begin VB.Label lblTW 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   8400
         TabIndex        =   109
         Top             =   960
         Width           =   975
      End
      Begin VB.Label lblW 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   8400
         TabIndex        =   108
         Top             =   4440
         Width           =   1005
      End
      Begin VB.Label lblW 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   8400
         TabIndex        =   107
         Top             =   4080
         Width           =   1005
      End
      Begin VB.Label lblW 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   8400
         TabIndex        =   106
         Top             =   3840
         Width           =   1005
      End
      Begin VB.Label lblW 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   8400
         TabIndex        =   105
         Top             =   3600
         Width           =   1005
      End
      Begin VB.Label lblW 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   8400
         TabIndex        =   104
         Top             =   3360
         Width           =   1005
      End
      Begin VB.Label lblW 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   8400
         TabIndex        =   103
         Top             =   3120
         Width           =   1005
      End
      Begin VB.Label lblW 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   8400
         TabIndex        =   102
         Top             =   2880
         Width           =   1005
      End
      Begin VB.Label lblW 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   8400
         TabIndex        =   101
         Top             =   2640
         Width           =   1005
      End
      Begin VB.Label lblW 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   8400
         TabIndex        =   100
         Top             =   2400
         Width           =   1005
      End
      Begin VB.Label lblW 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   8400
         TabIndex        =   99
         Top             =   2160
         Width           =   1005
      End
      Begin VB.Label lblW 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   8400
         TabIndex        =   98
         Top             =   1920
         Width           =   1005
      End
      Begin VB.Label lblW 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   8400
         TabIndex        =   97
         Top             =   1680
         Width           =   1005
      End
      Begin VB.Label lblW 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   8400
         TabIndex        =   96
         Top             =   1440
         Width           =   1005
      End
      Begin VB.Label lblREC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   13.5
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   10440
         TabIndex        =   93
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label lblRANK1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         Index           =   0
         Left            =   8040
         TabIndex        =   90
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label lblSP 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   255
         Left            =   0
         TabIndex        =   89
         Top             =   720
         Width           =   3500
      End
      Begin VB.Label lblBG 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   0
         TabIndex        =   88
         Top             =   480
         Width           =   3500
      End
      Begin VB.Label lblMN 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         Height          =   255
         Left            =   0
         TabIndex        =   87
         Top             =   240
         Width           =   3500
      End
      Begin VB.Label lblMJ 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   0
         TabIndex        =   86
         Top             =   0
         Width           =   3500
      End
      Begin VB.Label lblTV 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   6360
         TabIndex        =   85
         Top             =   960
         Width           =   1995
      End
      Begin VB.Label lblTP 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   5040
         TabIndex        =   84
         Top             =   960
         Width           =   1245
      End
      Begin VB.Label lblTU 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3960
         TabIndex        =   83
         Top             =   960
         Width           =   975
      End
      Begin VB.Label lblTE 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2520
         TabIndex        =   82
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label lblTC 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   720
         TabIndex        =   81
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label lblV 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   6360
         TabIndex        =   80
         Top             =   4440
         Width           =   1995
      End
      Begin VB.Label lblV 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   6360
         TabIndex        =   79
         Top             =   4080
         Width           =   1995
      End
      Begin VB.Label lblV 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   6360
         TabIndex        =   78
         Top             =   3840
         Width           =   1995
      End
      Begin VB.Label lblV 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   6360
         TabIndex        =   77
         Top             =   3600
         Width           =   1995
      End
      Begin VB.Label lblV 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   6360
         TabIndex        =   76
         Top             =   3360
         Width           =   1995
      End
      Begin VB.Label lblV 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   6360
         TabIndex        =   75
         Top             =   3120
         Width           =   1995
      End
      Begin VB.Label lblV 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   6360
         TabIndex        =   74
         Top             =   2880
         Width           =   1995
      End
      Begin VB.Label lblV 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   6360
         TabIndex        =   73
         Top             =   2640
         Width           =   1995
      End
      Begin VB.Label lblV 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   6360
         TabIndex        =   72
         Top             =   2400
         Width           =   1995
      End
      Begin VB.Label lblV 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   6360
         TabIndex        =   71
         Top             =   2160
         Width           =   1995
      End
      Begin VB.Label lblV 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   6360
         TabIndex        =   70
         Top             =   1920
         Width           =   1995
      End
      Begin VB.Label lblV 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   6360
         TabIndex        =   69
         Top             =   1680
         Width           =   1995
      End
      Begin VB.Label lblV 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   6360
         TabIndex        =   68
         Top             =   1440
         Width           =   1995
      End
      Begin VB.Label lblP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   5040
         TabIndex        =   67
         Top             =   4440
         Width           =   1250
      End
      Begin VB.Label lblP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   5040
         TabIndex        =   66
         Top             =   4080
         Width           =   1250
      End
      Begin VB.Label lblP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   5040
         TabIndex        =   65
         Top             =   3840
         Width           =   1250
      End
      Begin VB.Label lblP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   5040
         TabIndex        =   64
         Top             =   3600
         Width           =   1250
      End
      Begin VB.Label lblP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   5040
         TabIndex        =   63
         Top             =   3360
         Width           =   1250
      End
      Begin VB.Label lblP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   5040
         TabIndex        =   62
         Top             =   3120
         Width           =   1250
      End
      Begin VB.Label lblP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   5040
         TabIndex        =   61
         Top             =   2880
         Width           =   1250
      End
      Begin VB.Label lblP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   5040
         TabIndex        =   60
         Top             =   2640
         Width           =   1250
      End
      Begin VB.Label lblP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   5040
         TabIndex        =   59
         Top             =   2400
         Width           =   1250
      End
      Begin VB.Label lblP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   5040
         TabIndex        =   58
         Top             =   2160
         Width           =   1250
      End
      Begin VB.Label lblP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   5040
         TabIndex        =   57
         Top             =   1920
         Width           =   1250
      End
      Begin VB.Label lblP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   5040
         TabIndex        =   56
         Top             =   1680
         Width           =   1250
      End
      Begin VB.Label lblP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   5040
         TabIndex        =   55
         Top             =   1440
         Width           =   1250
      End
      Begin VB.Label lblU 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   3960
         TabIndex        =   54
         Top             =   4440
         Width           =   1000
      End
      Begin VB.Label lblU 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   3960
         TabIndex        =   53
         Top             =   4080
         Width           =   1000
      End
      Begin VB.Label lblU 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   3960
         TabIndex        =   52
         Top             =   3840
         Width           =   1000
      End
      Begin VB.Label lblU 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   3960
         TabIndex        =   51
         Top             =   3600
         Width           =   1000
      End
      Begin VB.Label lblU 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   3960
         TabIndex        =   50
         Top             =   3360
         Width           =   1000
      End
      Begin VB.Label lblU 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   3960
         TabIndex        =   49
         Top             =   3120
         Width           =   1000
      End
      Begin VB.Label lblU 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   3960
         TabIndex        =   48
         Top             =   2880
         Width           =   1000
      End
      Begin VB.Label lblU 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   3960
         TabIndex        =   47
         Top             =   2640
         Width           =   1000
      End
      Begin VB.Label lblU 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   3960
         TabIndex        =   46
         Top             =   2400
         Width           =   1000
      End
      Begin VB.Label lblU 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   3960
         TabIndex        =   45
         Top             =   2160
         Width           =   1000
      End
      Begin VB.Label lblU 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   3960
         TabIndex        =   44
         Top             =   1920
         Width           =   1000
      End
      Begin VB.Label lblU 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   3960
         TabIndex        =   43
         Top             =   1680
         Width           =   1000
      End
      Begin VB.Label lblU 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   3960
         TabIndex        =   42
         Top             =   1440
         Width           =   1000
      End
      Begin VB.Label lblE 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   2520
         TabIndex        =   41
         Top             =   4440
         Width           =   1335
      End
      Begin VB.Label lblE 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   2520
         TabIndex        =   40
         Top             =   4080
         Width           =   1335
      End
      Begin VB.Label lblE 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   2520
         TabIndex        =   39
         Top             =   3840
         Width           =   1335
      End
      Begin VB.Label lblE 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   2520
         TabIndex        =   38
         Top             =   3600
         Width           =   1335
      End
      Begin VB.Label lblE 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   2520
         TabIndex        =   37
         Top             =   3360
         Width           =   1335
      End
      Begin VB.Label lblE 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   2520
         TabIndex        =   36
         Top             =   3120
         Width           =   1335
      End
      Begin VB.Label lblE 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   2520
         TabIndex        =   35
         Top             =   2880
         Width           =   1335
      End
      Begin VB.Label lblE 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   2520
         TabIndex        =   34
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Label lblE 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   2520
         TabIndex        =   33
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label lblE 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   2520
         TabIndex        =   32
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label lblE 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   2520
         TabIndex        =   31
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label lblE 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   2520
         TabIndex        =   30
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label lblE 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   2520
         TabIndex        =   29
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label lblC 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   720
         TabIndex        =   28
         Top             =   4440
         Width           =   1695
      End
      Begin VB.Label lblC 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   720
         TabIndex        =   27
         Top             =   4080
         Width           =   1695
      End
      Begin VB.Label lblC 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   720
         TabIndex        =   26
         Top             =   3840
         Width           =   1695
      End
      Begin VB.Label lblC 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   720
         TabIndex        =   25
         Top             =   3600
         Width           =   1695
      End
      Begin VB.Label lblC 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   720
         TabIndex        =   24
         Top             =   3360
         Width           =   1695
      End
      Begin VB.Label lblC 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   720
         TabIndex        =   23
         Top             =   3120
         Width           =   1695
      End
      Begin VB.Label lblC 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   720
         TabIndex        =   22
         Top             =   2880
         Width           =   1695
      End
      Begin VB.Label lblC 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   720
         TabIndex        =   21
         Top             =   2640
         Width           =   1695
      End
      Begin VB.Label lblC 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   720
         TabIndex        =   20
         Top             =   2400
         Width           =   1695
      End
      Begin VB.Label lblC 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   720
         TabIndex        =   19
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Label lblC 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   720
         TabIndex        =   18
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label lblC 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   17
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label lblC 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   720
         TabIndex        =   16
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label lblMONTH 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   15
         Top             =   4440
         Width           =   495
      End
      Begin VB.Label lblMONTH 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "12"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   11
         Left            =   360
         TabIndex        =   14
         Top             =   4080
         Width           =   255
      End
      Begin VB.Label lblMONTH 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "11"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   10
         Left            =   360
         TabIndex        =   13
         Top             =   3840
         Width           =   255
      End
      Begin VB.Label lblMONTH 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   9
         Left            =   360
         TabIndex        =   12
         Top             =   3600
         Width           =   255
      End
      Begin VB.Label lblMONTH 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "09"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   8
         Left            =   360
         TabIndex        =   11
         Top             =   3360
         Width           =   255
      End
      Begin VB.Label lblMONTH 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "08"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   7
         Left            =   360
         TabIndex        =   10
         Top             =   3120
         Width           =   255
      End
      Begin VB.Label lblMONTH 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "07"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   6
         Left            =   360
         TabIndex        =   9
         Top             =   2880
         Width           =   255
      End
      Begin VB.Label lblMONTH 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "06"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   8
         Top             =   2640
         Width           =   255
      End
      Begin VB.Label lblMONTH 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "05"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   7
         Top             =   2400
         Width           =   255
      End
      Begin VB.Label lblMONTH 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "04"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   6
         Top             =   2160
         Width           =   255
      End
      Begin VB.Label lblMONTH 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "03"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   5
         Top             =   1920
         Width           =   255
      End
      Begin VB.Label lblMONTH 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "02"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   4
         Top             =   1680
         Width           =   255
      End
      Begin VB.Label lblMONTH 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "01"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   3
         Top             =   1440
         Width           =   255
      End
   End
   Begin VB.CommandButton cmdEXIT 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   12720
      MousePointer    =   1  'Arrow
      Picture         =   "frmREPORTS.frx":3B82
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8040
      Width           =   375
   End
   Begin VB.CommandButton cmdRETURN 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   12360
      MousePointer    =   1  'Arrow
      Picture         =   "frmREPORTS.frx":3E04
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7680
      Width           =   735
   End
   Begin VB.Label lblMSG 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   128
      Top             =   8520
      Width           =   11655
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FFFF&
      Caption         =   "04"
      Height          =   255
      Left            =   0
      TabIndex        =   127
      Top             =   8520
      Width           =   255
   End
   Begin VB.Label lblCONT 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Contents"
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
      Left            =   240
      TabIndex        =   92
      Top             =   0
      Width           =   9255
   End
End
Attribute VB_Name = "frmREPORTS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private DECIDE_VAL, DECIDE_SIZE, ADDFISH, ADDVALUE
Private TKEY(), TCONT(), TMJ(), TMN(), TBG(), TSP()
Private CURMJ, CURMN, CURBG, CURSP, CURREC
Private MJN(), MNN(), BGN(), SPN()
Private TC(), TE(), TU(), TP(), TV(), RANK, CUM, PER, CRIT, CRITCODE, CRITNAME
Private TT(), DETELM, DETTIT, RPER, RCUM, resp, TW(1 To 13), TF(1 To 13)

Private Sub cmdEXIT_Click()

Call write_parms

End

End Sub
Private Sub cmdEXPORT_Click()

On Error GoTo REPORT_ERROR

Dim fnm1, fnm2, DBN, XXX

fnm2 = APPROOT + "\ARTS\EXPORT\EXPORT.TXT"

Open fnm2 For Output As #2

Print #2, Format(CURY, "0000")
Print #2, " "
Print #2, msgtab(93)

Print #2, " "

fnm1 = APPROOT + "\ARTS\WORK\WSEL.TXT"

Open fnm1 For Input As #1

Do Until EOF(1)

Line Input #1, XXX
If Left(XXX, 5) = "=====" Then XXX = " "
Print #2, XXX

Loop

Close #1

Print #2, " "

If VALIDC = "Y" Then
DETELM = "C"
DETTIT = RTrim(msgtab(63)) + " (" + RTrim(UNW) + ")"
Call PREPARE_EXPORT
End If

If VALIDE = "Y" Then
DETELM = "E"
DETTIT = RTrim(msgtab(64)) + " (" + RTrim(msgtab(71)) + ")"
Call PREPARE_EXPORT
End If

If VALIDU = "Y" Then
DETELM = "U"
DETTIT = RTrim(msgtab(65)) + " (" + RTrim(UNW) + "/" + RTrim(msgtab(71)) + ")"
Call PREPARE_EXPORT
End If

If VALIDP = "Y" And TVALUE <> 0 Then
DETELM = "P"
DETTIT = RTrim(msgtab(66)) + " (" + RTrim(UNV) + "/" + RTrim(UNW) + ")"
Call PREPARE_EXPORT
End If

If VALIDV = "Y" And TVALUE <> 0 Then
DETELM = "V"
DETTIT = RTrim(msgtab(67)) + "(" + RTrim(UNV) + ")"
Call PREPARE_EXPORT
End If

If TFISH = 0 Then GoTo NO_CONT
If COMPGEN = "YES" And COMPFISH = "NOTOK" Then GoTo NO_CONT

If VALIDW = "Y" Then
DETELM = "W"
DETTIT = RTrim(msgtab(111))
Call PREPARE_EXPORT
End If

If VALIDF = "Y" Then
DETELM = "F"
DETTIT = RTrim(msgtab(112))
Call PREPARE_EXPORT
End If

NO_CONT:

Close #2

Call EXECUTE_EXCEL

Exit Sub

REPORT_ERROR:

Call EXCEL_REQUEST_ERROR

End Sub
Private Sub EXECUTE_EXCEL()

On Error GoTo CLOSE_EXCEL

FileCopy APPROOT + "\ARTS\BLANK_WORKSHEETS\ARTSER.XLS", APPROOT + "\ARTS\EXCEL_REPORTS\ARTFISH_WORK.XLS"

RUN_FILE = APPROOT + "\ARTS\EXCEL_REPORTS\GENERAL.BAT"

RUN_CODE = Shell(RUN_FILE, 4)

Exit Sub

CLOSE_EXCEL:

Call EXCEL_REQUEST_ERROR

End Sub
Private Sub cmdGUIDE_Click()

HTYPE = "40"

HFNM = APPROOT + "\ARTS\HELP\" + current_language + "HELP" + HTYPE + ".rtf"

If Dir(HFNM) = "" Then Exit Sub

frmREPORTS.Enabled = False
Load frmGUIDE
frmGUIDE.Show

End Sub

Private Sub cmdPLOT_Click()

Load frmPLOT
frmREPORTS.Enabled = False
frmPLOT.Show

End Sub
Private Sub cmdRETURN_Click()

Load frmSEL
Unload frmREPORTS
frmSEL.Show

End Sub
Private Sub Form_Load()

lblSIZE.Caption = msgtab(114) + " " + msgtab(106)
lblVALUES.Caption = msgtab(114) + " " + msgtab(106)
lblMSG.Caption = msgtab(105) + " " + msgtab(106)
optSTD.Caption = msgtab(126)
optLOC.Caption = msgtab(127)

optLOC.Value = True
optSTD.Value = False

lblSIZE.Visible = False
lblMSG.Visible = False
lblVALUES.Visible = False

Set Picture = LoadPicture(APPROOT + "\ARTS\PICS_RUNTIME\SCREEN_04.JPG")

Dim J

Call GENERAL_CLEANING

cmdPLOT.Visible = False

Dim fnm

fnm = APPROOT + "\ARTS\TABLES\Y" + Format(CURY, "0000") + "_UNITS.TXT"

Open fnm For Input As #1

Line Input #1, UNW
Line Input #1, UNV

Close #1

UNW = RTrim(UNW): UNV = RTrim(UNV)

frmREPORTS.MousePointer = 1
frmREPORTS.Caption = msgtab(15) + ": " + Format(CURY, "0000") + " - " + _
                 msgtab(54)

cmdRETURN.ToolTipText = msgtab(56)
cmdEXIT.ToolTipText = msgtab(57)
cmdPLOT.ToolTipText = msgtab(90)

cmdEXPORT.ToolTipText = msgtab(93)
cmdGUIDE.ToolTipText = msgtab(6)

lblCONT.Caption = msgtab(62)

lblTC.Caption = msgtab(63) + " (" + UNW + ")"
lblTE.Caption = msgtab(64) + " (" + RTrim(msgtab(71)) + ")"
lblTU.Caption = msgtab(65)
lblTP.Caption = msgtab(66)
lblTV.Caption = msgtab(67) + " (" + UNV + ")"
lblTW.Caption = msgtab(111)
lblTF.Caption = msgtab(112)

lblMONTH(12).Caption = CURY

Dim I

For I = 0 To 12
lblC(I).Caption = " "
lblE(I).Caption = " "
lblU(I).Caption = " "
lblP(I).Caption = " "
lblV(I).Caption = " "
lblW(I).Caption = "  "
lblF(I).Caption = "  "
Next I

lblMJ.Caption = " "
lblMN.Caption = " "
lblBG.Caption = " "
lblSP.Caption = " "
lblRANKC.Caption = " "
lblRANK2.Caption = " "

Call LOAD_MAJOR
Call LOAD_MINOR
Call LOAD_BG
Call LOAD_SPECIES

Call PREP_CONTENTS

End Sub
Private Sub PREP_CONTENTS()

Dim LLL

lstDATA.Clear

Dim I, J, K, DBN, NREC, IREC, XKEY, XXX

DBN = APPROOT + "\ARTS\WORK\WGS" + Format(CURY, "0000") + ".MDB"

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(DBN)
Set prm_record = prm_database.OpenRecordset("ASITAB")

With prm_record

.MoveFirst

CRITCODE = " "

If RTrim(![CRIT]) <> "" Then
   CRITCODE = Left(![CRIT], 3)
   RANKYN = "Y"
   LLL = Len(RTrim(![CRIT]))
   RANK_CRIT = Right(RTrim(![CRIT]), LLL - 4)
   CRITNAME = RANK_CRIT
   End If

If RANKYN <> "Y" Then .Index = "primarykey"
If RANKYN = "Y" Then .Index = "rank"

If RANKYN <> "Y" Then
   lblRANKC.Visible = False
   lblRANK2.Visible = False
   End If

NREC = .RecordCount

ReDim TKEY(1 To NREC), TMJ(1 To NREC), TMN(1 To NREC), TBG(1 To NREC), TSP(1 To NREC)

VALIDC = "N": VALIDE = "N": VALIDU = "N": VALIDP = "N": VALIDV = "N"
VALIDW = "N": VALIDF = "N"

.MoveFirst

If RANKYN = "Y" Then
   lblRANKC.Caption = msgtab(85)
   lblRANK2.Caption = RANK_CRIT
   End If
   
IREC = 0

Do Until .EOF

If ![C13] <> 0 Then VALIDC = "Y"
If ![E13] <> 0 Then VALIDE = "Y"
If ![U13] <> 0 Then VALIDU = "Y"
If ![P13] <> 0 Then VALIDP = "Y"
If ![V13] <> 0 Then VALIDV = "Y"
If ![W13] <> 0 Then VALIDW = "Y"
If ![F13] <> 0 Then VALIDF = "Y"

IREC = IREC + 1

XKEY = ![akey]: TKEY(IREC) = XKEY

If Left(XKEY, 3) = " GT" Then
   TMJ(IREC) = msgtab(68)
   TMN(IREC) = "": TBG(IREC) = "": TSP(IREC) = ""
   GoTo CONT_1
   End If

'Decode
'============

I = Val(Mid(XKEY, 2, 4))

If I <> 0 Then TMJ(IREC) = MJN(I)
'If I = 0 Then TMJ(IREC) = msgtab(58)
If I = 0 Then TMJ(IREC) = ""

I = Val(Mid(XKEY, 8, 4))

If I <> 0 Then TMN(IREC) = MNN(I)
'If I = 0 Then TMN(IREC) = msgtab(59)
If I = 0 Then TMN(IREC) = ""

If Mid(XKEY, 13, 1) = "B" Then

    I = Val(Mid(XKEY, 14, 4))

    If I <> 0 Then TBG(IREC) = BGN(I)
    'If I = 0 Then TBG(IREC) = msgtab(60)
    If I = 0 Then TBG(IREC) = ""
    
    End If

If Mid(XKEY, 19, 1) = "B" Then

    I = Val(Mid(XKEY, 20, 4))

    If I <> 0 Then TBG(IREC) = BGN(I)
    'If I = 0 Then TBG(IREC) = msgtab(60)
    If I = 0 Then TBG(IREC) = ""
    
    End If

If Mid(XKEY, 13, 1) = "S" Then

    I = Val(Mid(XKEY, 14, 4))

    If I <> 0 Then TSP(IREC) = SPN(I)
    'If I = 0 Then TSP(IREC) = msgtab(61)
    If I = 0 Then TSP(IREC) = ""
    
    End If

If Mid(XKEY, 19, 1) = "S" Then

    I = Val(Mid(XKEY, 20, 4))

    If I <> 0 Then TSP(IREC) = SPN(I)
    'If I = 0 Then TSP(IREC) = msgtab(61)
    If I = 0 Then TSP(IREC) = ""
    
    End If

CONT_1:

XXX = Format(IREC, "00000")

If TMJ(IREC) <> "" Then XXX = XXX + " " + RTrim(TMJ(IREC))

If TMN(IREC) <> "" Then XXX = XXX + " / " + RTrim(TMN(IREC))

If TBG(IREC) <> "" Then XXX = XXX + " / " + RTrim(TBG(IREC))

If TSP(IREC) <> "" Then XXX = XXX + " / " + RTrim(TSP(IREC))

lstDATA.AddItem XXX

.MoveNext

Loop

End With

prm_record.Close
prm_database.Close

End Sub
Private Sub LOAD_BG()

Dim I, fnm, XXX

fnm = APPROOT + "\ARTS\TABLES\Y" + Format(CURY, "0000") + "_BG.TXT"

If Dir(fnm) = "" Then Exit Sub

ReDim BGN(1 To 10000)

Open fnm For Input As #1

Do Until EOF(1)

Line Input #1, XXX

I = Val(Left(XXX, 4)): BGN(I) = Mid(XXX, 6, 30)

If optSTD.Value = True Then BGN(I) = Mid(XXX, 37, 30)

Loop

Close #1

End Sub
Private Sub LOAD_MAJOR()

Dim I, fnm, XXX

fnm = APPROOT + "\ARTS\TABLES\Y" + Format(CURY, "0000") + "_MAJOR.TXT"

If Dir(fnm) = "" Then Exit Sub

ReDim MJN(1 To 10000)

Open fnm For Input As #1

Do Until EOF(1)

Line Input #1, XXX

I = Val(Left(XXX, 4)): MJN(I) = Mid(XXX, 6, 30)
If optSTD.Value = True Then MJN(I) = Mid(XXX, 37, 30)

Loop

Close #1

End Sub
Private Sub LOAD_MINOR()

Dim I, fnm, XXX

ReDim MNN(1 To 10000)

fnm = APPROOT + "\ARTS\TABLES\Y" + Format(CURY, "0000") + "_MINOR.TXT"

If Dir(fnm) = "" Then Exit Sub

Open fnm For Input As #1

Do Until EOF(1)

Line Input #1, XXX

I = Val(Left(XXX, 4)): MNN(I) = Mid(XXX, 6, 30)
If optSTD.Value = True Then MNN(I) = Mid(XXX, 111, 30)

Loop

Close #1

End Sub
Private Sub LOAD_SPECIES()

Dim I, fnm, XXX

ReDim SPN(1 To 10000)

fnm = APPROOT + "\ARTS\TABLES\Y" + Format(CURY, "0000") + "_SPECIES.TXT"

If Dir(fnm) = "" Then Exit Sub

Open fnm For Input As #1

Do Until EOF(1)

Line Input #1, XXX

I = Val(Left(XXX, 4)): SPN(I) = Mid(XXX, 6, 30)
If optSTD.Value = True Then SPN(I) = Mid(XXX, 37, 30)

Loop

Close #1

End Sub
Private Sub lblMSG_Click()

Load frmNOPRICES
frmNOPRICES.Show

End Sub

Private Sub lblSIZE_Click()
Load frmNOPRICES
frmNOPRICES.Show
End Sub

Private Sub lblVALUES_Click()

Load frmNOPRICES
frmNOPRICES.Show

End Sub

Private Sub lstDATA_Click()

cmdPLOT.Visible = True

Dim I, DBN, XKEY, J, FFC, FFE, FFU, FFP, FFV, FFW, FFN, RPER, RCUM

For I = 0 To 12
   lblW(I).ForeColor = vbBlack: lblF(I).ForeColor = vbBlack
   lblSIZE.Visible = False
   lblP(I).ForeColor = vbBlack: lblV(I).ForeColor = vbBlack
   lblVALUES.Visible = False
   Next I

CURREC = lstDATA.ListIndex + 1: I = CURREC

lblREC.Caption = Format(CURREC, "00000")

CURMJ = TMJ(I): CURMN = TMN(I): CURBG = TBG(I): CURSP = TSP(I)

lblMJ.Caption = CURMJ
lblMN.Caption = CURMN
lblBG.Caption = CURBG
lblSP.Caption = CURSP

REPMJN = CURMJ: REPMNN = CURMN: REPBGN = CURBG: REPSPN = CURSP

ReDim TC(1 To 13), TE(1 To 13), TU(1 To 13), TP(1 To 13), TV(1 To 13)

DBN = APPROOT + "\ARTS\WORK\WGS" + Format(CURY, "0000") + ".MDB"

If Dir(DBN) = "" Then DBN = APPROOT + "\ARTS\WORK\WS" + Format(CURY, "0000") + ".MDB"

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(DBN)
Set prm_record = prm_database.OpenRecordset("ASITAB")

With prm_record

.Index = "primarykey"

XKEY = TKEY(I)

.Seek "=", XKEY

RPER = ![PER]: RCUM = ![CUM]

ReDim REPTC(1 To 13), REPTE(1 To 13), REPTU(1 To 13)
ReDim REPTP(1 To 13), REPTV(1 To 13), REPTW(1 To 13)

For J = 1 To 13

FFC = "C" + Format(J, "00"): TC(J) = .Fields(FFC)
FFE = "E" + Format(J, "00"): TE(J) = .Fields(FFE)
FFU = "U" + Format(J, "00"): TU(J) = .Fields(FFU)
FFP = "P" + Format(J, "00"): TP(J) = .Fields(FFP)
FFV = "V" + Format(J, "00"): TV(J) = .Fields(FFV)
FFW = "W" + Format(J, "00"): TW(J) = .Fields(FFW)
FFN = "F" + Format(J, "00"): TF(J) = .Fields(FFN)

ADDFISH = ![ADDFISH]
ADDVALUE = ![ADDVALUES]

REPTC(J) = TC(J)
REPTE(J) = TE(J)
REPTU(J) = TU(J)
REPTP(J) = TP(J)
REPTV(J) = TV(J)
REPTW(J) = TW(J)

Next J

End With

prm_record.Close
prm_database.Close

For J = 1 To 13

If J <= 12 Then

   If VALID_MONTHS(J) = "-" Then
   
   lblC(J - 1).Caption = " "
   lblC(J - 1).BackColor = QBColor(8)
   lblE(J - 1).Caption = " "
   lblE(J - 1).BackColor = QBColor(8)
   lblU(J - 1).Caption = " "
   lblU(J - 1).BackColor = QBColor(8)
   lblP(J - 1).Caption = " "
   lblP(J - 1).BackColor = QBColor(8)
   lblV(J - 1).Caption = " "
   lblV(J - 1).BackColor = QBColor(8)
   lblW(J - 1).Caption = " "
   lblW(J - 1).BackColor = QBColor(8)
   lblF(J - 1).Caption = " "
   lblF(J - 1).BackColor = QBColor(8)
   GoTo NEXT_J
   
   End If
   
   End If

'----------- Check ADDFISH ----------
If ADDFISH = "NO" And lblTF.Visible = True Then

   lblSIZE.Visible = True
   For I = 0 To 12
   lblW(I).ForeColor = vbRed: lblF(I).ForeColor = vbRed
   Next I
   
   End If
'----------- Check ADDVALUE ----------
If ADDVALUE = "NO" And lblTP.Visible = True Then
   lblVALUES.Visible = True
   For I = 0 To 12
   lblP(I).ForeColor = vbRed: lblV(I).ForeColor = vbRed
   Next I
   End If
'-----------------------------------------


lblC(J - 1).Caption = Format(TC(J), "### ### ### ### ##0")
If TC(J) = 0 Then lblC(J - 1).Caption = "..."

lblE(J - 1).Caption = Format(TE(J), "### ### ### ### ##0")
If TE(J) = 0 Then lblE(J - 1).Caption = "..."

lblU(J - 1).Caption = Format(TU(J), "### ##0.000")
If TU(J) = 0 Then lblU(J - 1).Caption = "..."

lblP(J - 1).Caption = Format(TP(J), "### ##0.000")
If TP(J) = 0 Then lblP(J - 1).Caption = "..."

lblV(J - 1).Caption = Format(TV(J), "### ### ### ### ##0")
If TV(J) = 0 Then lblV(J - 1).Caption = "..."

lblF(J - 1).Caption = Format(TF(J), "### ### ### ### ### ##0")
If TF(J) = 0 Then lblF(J - 1).Caption = "..."

lblW(J - 1).Caption = Format(TW(J), "### ##0.000")
If TW(J) = 0 Then lblW(J - 1).Caption = "..."

NEXT_J:

Next J

If ADDFISH = "NO" And lblTF.Visible = True Then
lblW(12).Caption = "...": lblF(12).Caption = "..."
End If

'If lblMSG.Visible = True And lblW(12).Visible = False Then
'lblV(12).Caption = "...": lblP(12).Caption = "..."
'End If

For I = 0 To 11
If lblC(I).Caption <> "..." And lblP(I).Caption = "..." Then
lblV(12).Caption = "...": lblP(12).Caption = "..."
End If
Next I

If lblW(12).Caption <> "..." And lblP(12).Caption <> "..." Then
lblMSG.Visible = False
End If

If lblW(12).Caption = "..." Or lblP(12).Caption = "..." Then
lblMSG.Visible = True
End If

Dim KK, MM, CEUPV

If CRITCODE = " " Then GoTo NO_RANK

KK = Val(Mid(CRITCODE, 2, 2))

CEUPV = Left(CRITCODE, 1)

If CEUPV = "C" Then
   lblC(KK - 1).BackColor = vbBlue
   lblC(KK - 1).ForeColor = vbWhite
   End If

If CEUPV = "E" Then
   lblE(KK - 1).BackColor = vbBlue
   lblE(KK - 1).ForeColor = vbWhite
   End If

If CEUPV = "U" Then
   lblU(KK - 1).BackColor = vbBlue
   lblU(KK - 1).ForeColor = vbWhite
   End If

If CEUPV = "P" Then
   lblP(KK - 1).BackColor = vbBlue
   lblP(KK - 1).ForeColor = vbWhite
   End If

If CEUPV = "V" Then
   lblV(KK - 1).BackColor = vbBlue
   lblV(KK - 1).ForeColor = vbWhite
   End If

If RCUM = 0 Then GoTo NO_RANK

lblRANK2 = CRITNAME + ":  " + Format(RPER, "##0.00") + " %, " + " " + _
           msgtab(89) + " " + Format(RCUM, "##0.00") + " %"
         

NO_RANK:

End Sub
Private Sub PREPARE_EXPORT()

Dim PMJ1, PMN1, PBG1, PSP1, PMJ2, PMN2, PBG2, PSP2

PMJ1 = "*": PMN1 = "*": PBG1 = "*": PSP1 = "*"

Dim LLL

Dim I, J, K, DBN, NREC, IREC, XKEY, XXX, FFF

DBN = APPROOT + "\ARTS\WORK\WGS" + Format(CURY, "0000") + ".MDB"

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(DBN)
Set prm_record = prm_database.OpenRecordset("ASITAB")

With prm_record

.MoveFirst

CRITCODE = " "

If RTrim(![CRIT]) <> "" Then
   CRITCODE = Left(![CRIT], 3)
   RANKYN = "Y"
   LLL = Len(RTrim(![CRIT]))
   RANK_CRIT = Right(RTrim(![CRIT]), LLL - 4)
   CRITNAME = RANK_CRIT
   End If

If RANKYN <> "Y" Then .Index = "primarykey"
If RANKYN = "Y" Then .Index = "rank"

NREC = .RecordCount

Dim PRLN()

ReDim PRLN(1 To 14)

Dim PRF1, PRF2, PRM1, PRM2, PRM3, PRM4, PRM5, PRM6, PRFM

PRF1 = "###,###,###,###,###,##0": PRF2 = "###,###,##0.000"

PRFM = PRF1

If DETELM = "P" Or DETELM = "U" Then PRFM = PRF2

PRLN(1) = " "
PRLN(2) = "[" + Format(CURY, "0000") + "]"

For I = 1 To 12
PRLN(I + 2) = "[" + Format(I, "00") + "]"
Next I

Print #2, " "

Print #2, DETTIT

Write #2, " "; " "; " "; PRLN(1); PRLN(2); PRLN(3); PRLN(4); PRLN(5); PRLN(6); _
          PRLN(7); PRLN(8); PRLN(9); PRLN(10); PRLN(11); PRLN(12); PRLN(13); PRLN(14)

PRLN(1) = " "

For I = 2 To 14
PRLN(I) = " "
Next I

Write #2, " "; " "; " "; " "; PRLN(2); PRLN(3); PRLN(4); PRLN(5); PRLN(6); _
          PRLN(7); PRLN(8); PRLN(9); PRLN(10); PRLN(11); PRLN(12); PRLN(13); PRLN(14)

.MoveFirst

IREC = 0

Do Until .EOF

IREC = IREC + 1

XKEY = ![akey]
RPER = ![PER]: RCUM = ![CUM]

ReDim TT(1 To 13)

For I = 1 To 13

FFF = DETELM + Format(I, "00")
TT(I) = .Fields(FFF)

PRFM = PRF1

If DETELM = "U" Or DETELM = "P" Or DETELM = "W" Then PRFM = PRF2

If TT(I) = 0 Then
   TT(I) = "..."
   GoTo NO_FORMAT
   End If

If DETELM = "U" Or DETELM = "P" Or DETELM = "W" Then
   TT(I) = Int(TT(I) * 1000) / 1000
   End If

NO_FORMAT:

Next I

If Left(XKEY, 3) = " GT" Then
   CURMJ = RTrim(msgtab(68))
   CURMN = " ": CURBG = " ": CURSP = " "
   GoTo CONT_1
   End If

ReDim TT(1 To 13)

For I = 1 To 13

FFF = DETELM + Format(I, "00")
TT(I) = .Fields(FFF)

PRFM = PRF1

If DETELM = "U" Or DETELM = "P" Or DETELM = "W" Then PRFM = PRF2

If TT(I) = 0 Then
   TT(I) = "..."
   GoTo NO_FORMAT2
   End If

If DETELM = "U" Or DETELM = "P" Or DETELM = "W" Then
   TT(I) = Int(TT(I) * 1000) / 1000
   End If


NO_FORMAT2:

Next I

'Decode
'============

I = Val(Mid(XKEY, 2, 4))

If I <> 0 Then CURMJ = RTrim(MJN(I))
If I = 0 Then CURMJ = " "

I = Val(Mid(XKEY, 8, 4))

If I <> 0 Then CURMN = RTrim(MNN(I))
If I = 0 Then CURMN = " "

If Mid(XKEY, 13, 1) = "B" Then

    I = Val(Mid(XKEY, 14, 4))

    If I <> 0 Then CURBG = RTrim(BGN(I))
    If I = 0 Then CURBG = " "
    
    End If

If Mid(XKEY, 19, 1) = "B" Then

    I = Val(Mid(XKEY, 20, 4))

    If I <> 0 Then CURBG = RTrim(BGN(I))
    If I = 0 Then CURBG = " "
    
    End If

If Mid(XKEY, 13, 1) = "S" Then

    I = Val(Mid(XKEY, 14, 4))

    If I <> 0 Then CURSP = RTrim(SPN(I))
    If I = 0 Then CURSP = " "
    
    End If

If Mid(XKEY, 19, 1) = "S" Then

    I = Val(Mid(XKEY, 20, 4))

    If I <> 0 Then CURSP = RTrim(SPN(I))
    If I = 0 Then CURSP = " "
    
    End If

CONT_1:

'FORMATTING
'==========

CONT_FORM:

If Mid(XKEY, 13, 1) <> "S" Then
Write #2, CURMJ; CURMN; CURBG; CURSP; TT(13); TT(1); TT(2); TT(3); TT(4); TT(5); TT(6); TT(7); _
               TT(8); TT(9); TT(10); TT(11); TT(12)
End If

If Mid(XKEY, 13, 1) = "S" Then
Write #2, CURMJ; CURMN; CURSP; CURBG; TT(13); TT(1); TT(2); TT(3); TT(4); TT(5); TT(6); TT(7); _
               TT(8); TT(9); TT(10); TT(11); TT(12)
End If

.MoveNext

Loop

End With

Print #2, " "

prm_record.Close
prm_database.Close

Close #1

End Sub
Private Sub GENERAL_CLEANING()

AW_OPTION = 0

Dim J

'--------Neutralize sizes etc. if system empty

If TFISH = 0 Then

lblTW.Visible = False: lblTF.Visible = False: lblSIZE.Visible = False

For J = 0 To 12
lblW(J).Visible = False: lblF(J).Visible = False
Next J

End If

'--------Neutralize price-values, etc. if system empty

If TVALUE = 0 Then

lblTP.Visible = False: lblTV.Visible = False: lblVALUES.Visible = False

For J = 0 To 12
lblP(J).Visible = False: lblV(J).Visible = False
Next J

End If

'------------------------------------------------

If TFISH = 0 And TVALUE = 0 Then lblMSG.Visible = False

'=============================================================

If COMPGEN = "YES" And COMPFISH = "NOTOK" Then
   lblTW.Visible = False: lblTF.Visible = False: lblSIZE.Visible = False
      
   For J = 0 To 12
   lblW(J).Visible = False: lblF(J).Visible = False
   Next J

   End If

lblMSG.Visible = False

If COMPGEN = "YES" And PCATCH <> 0 Then lblMSG.Visible = True
If COMPGEN = "YES" And FCATCH <> 0 Then lblMSG.Visible = True
           
If COMPGEN = "NO" Then lblMSG.Visible = False
           
If lblW(1).Visible = True Then AW_OPTION = 1
           
End Sub
Private Sub optLOC_Click()

optSTD.Value = False

Call LOAD_MAJOR
Call LOAD_MINOR
Call LOAD_BG
Call LOAD_SPECIES
Call PREP_CONTENTS

End Sub

Private Sub optSTD_Click()

optLOC.Value = False

Call LOAD_MAJOR
Call LOAD_MINOR
Call LOAD_BG
Call LOAD_SPECIES
Call PREP_CONTENTS

Exit Sub

End Sub
