VERSION 5.00
Begin VB.Form frmEFFORT 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10815
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   1  'Arrow
   Moveable        =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   10815
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGUIDE1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   9000
      MousePointer    =   1  'Arrow
      Picture         =   "EFFORT.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   142
      Top             =   6720
      Width           =   735
   End
   Begin VB.CommandButton cmdLEAVE 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   3480
      MousePointer    =   1  'Arrow
      Picture         =   "EFFORT.frx":2262
      Style           =   1  'Graphical
      TabIndex        =   139
      Top             =   6720
      Width           =   735
   End
   Begin VB.CommandButton cmdSTAY 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   1800
      Picture         =   "EFFORT.frx":24E4
      Style           =   1  'Graphical
      TabIndex        =   138
      Top             =   6720
      Width           =   735
   End
   Begin VB.CommandButton cmdCONFIRM 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      Picture         =   "EFFORT.frx":25EE
      Style           =   1  'Graphical
      TabIndex        =   136
      Top             =   3960
      Width           =   495
   End
   Begin VB.ListBox lstMINOR 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   3180
      Left            =   120
      TabIndex        =   134
      Top             =   480
      Width           =   4815
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   120
      TabIndex        =   5
      Top             =   2880
      Width           =   10575
      Begin VB.TextBox txtREC 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7920
         TabIndex        =   140
         Top             =   3360
         Width           =   2535
      End
      Begin VB.TextBox txtACT 
         Alignment       =   1  'Right Justify
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
         Index           =   0
         Left            =   360
         TabIndex        =   98
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtACT 
         Alignment       =   1  'Right Justify
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
         Index           =   1
         Left            =   360
         TabIndex        =   97
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txtACT 
         Alignment       =   1  'Right Justify
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
         Index           =   2
         Left            =   360
         TabIndex        =   96
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox txtACT 
         Alignment       =   1  'Right Justify
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
         Index           =   3
         Left            =   360
         TabIndex        =   95
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox txtACT 
         Alignment       =   1  'Right Justify
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
         Index           =   4
         Left            =   360
         TabIndex        =   94
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox txtACT 
         Alignment       =   1  'Right Justify
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
         Index           =   5
         Left            =   360
         TabIndex        =   93
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox txtACT 
         Alignment       =   1  'Right Justify
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
         Index           =   6
         Left            =   360
         TabIndex        =   92
         Top             =   2640
         Width           =   735
      End
      Begin VB.TextBox txtACT 
         Alignment       =   1  'Right Justify
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
         Index           =   7
         Left            =   360
         TabIndex        =   91
         Top             =   3000
         Width           =   735
      End
      Begin VB.TextBox txtACT 
         Alignment       =   1  'Right Justify
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
         Index           =   8
         Left            =   3000
         TabIndex        =   90
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtACT 
         Alignment       =   1  'Right Justify
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
         Index           =   9
         Left            =   3000
         TabIndex        =   89
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txtACT 
         Alignment       =   1  'Right Justify
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
         Index           =   10
         Left            =   3000
         TabIndex        =   88
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox txtACT 
         Alignment       =   1  'Right Justify
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
         Index           =   11
         Left            =   3000
         TabIndex        =   87
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox txtACT 
         Alignment       =   1  'Right Justify
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
         Index           =   12
         Left            =   3000
         TabIndex        =   86
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox txtACT 
         Alignment       =   1  'Right Justify
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
         Index           =   13
         Left            =   3000
         TabIndex        =   85
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox txtACT 
         Alignment       =   1  'Right Justify
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
         Index           =   14
         Left            =   3000
         TabIndex        =   84
         Top             =   2640
         Width           =   735
      End
      Begin VB.TextBox txtACT 
         Alignment       =   1  'Right Justify
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
         Index           =   15
         Left            =   3000
         TabIndex        =   83
         Top             =   3000
         Width           =   735
      End
      Begin VB.TextBox txtACT 
         Alignment       =   1  'Right Justify
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
         Index           =   16
         Left            =   5640
         TabIndex        =   82
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtACT 
         Alignment       =   1  'Right Justify
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
         Index           =   17
         Left            =   5640
         TabIndex        =   81
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txtACT 
         Alignment       =   1  'Right Justify
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
         Index           =   18
         Left            =   5640
         TabIndex        =   80
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox txtACT 
         Alignment       =   1  'Right Justify
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
         Index           =   19
         Left            =   5640
         TabIndex        =   79
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox txtACT 
         Alignment       =   1  'Right Justify
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
         Index           =   20
         Left            =   5640
         TabIndex        =   78
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox txtACT 
         Alignment       =   1  'Right Justify
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
         Index           =   21
         Left            =   5640
         TabIndex        =   77
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox txtACT 
         Alignment       =   1  'Right Justify
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
         Index           =   22
         Left            =   5640
         TabIndex        =   76
         Top             =   2640
         Width           =   735
      End
      Begin VB.TextBox txtACT 
         Alignment       =   1  'Right Justify
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
         Index           =   23
         Left            =   5640
         TabIndex        =   75
         Top             =   3000
         Width           =   735
      End
      Begin VB.TextBox txtACT 
         Alignment       =   1  'Right Justify
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
         Index           =   24
         Left            =   8280
         TabIndex        =   74
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtACT 
         Alignment       =   1  'Right Justify
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
         Index           =   25
         Left            =   8280
         TabIndex        =   73
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txtACT 
         Alignment       =   1  'Right Justify
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
         Index           =   26
         Left            =   8280
         TabIndex        =   72
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox txtACT 
         Alignment       =   1  'Right Justify
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
         Index           =   27
         Left            =   8280
         TabIndex        =   71
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox txtACT 
         Alignment       =   1  'Right Justify
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
         Index           =   28
         Left            =   8280
         TabIndex        =   70
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox txtACT 
         Alignment       =   1  'Right Justify
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
         Index           =   29
         Left            =   8280
         TabIndex        =   69
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox txtACT 
         Alignment       =   1  'Right Justify
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
         Index           =   30
         Left            =   8280
         TabIndex        =   68
         Top             =   2640
         Width           =   735
      End
      Begin VB.TextBox txtSMP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   0
         Left            =   1080
         TabIndex        =   67
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtSMP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   1
         Left            =   1080
         TabIndex        =   66
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txtSMP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   2
         Left            =   1080
         TabIndex        =   65
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox txtSMP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   3
         Left            =   1080
         TabIndex        =   64
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox txtSMP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   4
         Left            =   1080
         TabIndex        =   63
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox txtSMP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   5
         Left            =   1080
         TabIndex        =   62
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox txtSMP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   6
         Left            =   1080
         TabIndex        =   61
         Top             =   2640
         Width           =   735
      End
      Begin VB.TextBox txtSMP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   7
         Left            =   1080
         TabIndex        =   60
         Top             =   3000
         Width           =   735
      End
      Begin VB.TextBox txtSMP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   8
         Left            =   3720
         TabIndex        =   59
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtSMP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   9
         Left            =   3720
         TabIndex        =   58
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txtSMP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   10
         Left            =   3720
         TabIndex        =   57
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox txtSMP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   11
         Left            =   3720
         TabIndex        =   56
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox txtSMP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   12
         Left            =   3720
         TabIndex        =   55
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox txtSMP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   13
         Left            =   3720
         TabIndex        =   54
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox txtSMP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   14
         Left            =   3720
         TabIndex        =   53
         Top             =   2640
         Width           =   735
      End
      Begin VB.TextBox txtSMP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   15
         Left            =   3720
         TabIndex        =   52
         Top             =   3000
         Width           =   735
      End
      Begin VB.TextBox txtSMP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   16
         Left            =   6360
         TabIndex        =   51
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtSMP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   17
         Left            =   6360
         TabIndex        =   50
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txtSMP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   18
         Left            =   6360
         TabIndex        =   49
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox txtSMP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   19
         Left            =   6360
         TabIndex        =   48
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox txtSMP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   20
         Left            =   6360
         TabIndex        =   47
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox txtSMP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   21
         Left            =   6360
         TabIndex        =   46
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox txtSMP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   22
         Left            =   6360
         TabIndex        =   45
         Top             =   2640
         Width           =   735
      End
      Begin VB.TextBox txtSMP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   23
         Left            =   6360
         TabIndex        =   44
         Top             =   3000
         Width           =   735
      End
      Begin VB.TextBox txtSMP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   24
         Left            =   9000
         TabIndex        =   43
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtSMP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   25
         Left            =   9000
         TabIndex        =   42
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txtSMP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   26
         Left            =   9000
         TabIndex        =   41
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox txtSMP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   27
         Left            =   9000
         TabIndex        =   40
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox txtSMP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   28
         Left            =   9000
         TabIndex        =   39
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox txtSMP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   29
         Left            =   9000
         TabIndex        =   38
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox txtSMP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   30
         Left            =   9000
         TabIndex        =   37
         Top             =   2640
         Width           =   735
      End
      Begin VB.TextBox txtFRM 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   1800
         TabIndex        =   36
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtFRM 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   1800
         TabIndex        =   35
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txtFRM 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   1800
         TabIndex        =   34
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox txtFRM 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   1800
         TabIndex        =   33
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox txtFRM 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   1800
         TabIndex        =   32
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox txtFRM 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   1800
         TabIndex        =   31
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox txtFRM 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   1800
         TabIndex        =   30
         Top             =   2640
         Width           =   735
      End
      Begin VB.TextBox txtFRM 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   1800
         TabIndex        =   29
         Top             =   3000
         Width           =   735
      End
      Begin VB.TextBox txtFRM 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   4440
         TabIndex        =   28
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtFRM 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   4440
         TabIndex        =   27
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txtFRM 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   10
         Left            =   4440
         TabIndex        =   26
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox txtFRM 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   11
         Left            =   4440
         TabIndex        =   25
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox txtFRM 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   12
         Left            =   4440
         TabIndex        =   24
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox txtFRM 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   13
         Left            =   4440
         TabIndex        =   23
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox txtFRM 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   14
         Left            =   4440
         TabIndex        =   22
         Top             =   2640
         Width           =   735
      End
      Begin VB.TextBox txtFRM 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   15
         Left            =   4440
         TabIndex        =   21
         Top             =   3000
         Width           =   735
      End
      Begin VB.TextBox txtFRM 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   16
         Left            =   7080
         TabIndex        =   20
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtFRM 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   17
         Left            =   7080
         TabIndex        =   19
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txtFRM 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   18
         Left            =   7080
         TabIndex        =   18
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox txtFRM 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   19
         Left            =   7080
         TabIndex        =   17
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox txtFRM 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   20
         Left            =   7080
         TabIndex        =   16
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox txtFRM 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   21
         Left            =   7080
         TabIndex        =   15
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox txtFRM 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   22
         Left            =   7080
         TabIndex        =   14
         Top             =   2640
         Width           =   735
      End
      Begin VB.TextBox txtFRM 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   23
         Left            =   7080
         TabIndex        =   13
         Top             =   3000
         Width           =   735
      End
      Begin VB.TextBox txtFRM 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   24
         Left            =   9720
         TabIndex        =   12
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtFRM 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   25
         Left            =   9720
         TabIndex        =   11
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txtFRM 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   26
         Left            =   9720
         TabIndex        =   10
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox txtFRM 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   27
         Left            =   9720
         TabIndex        =   9
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox txtFRM 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   28
         Left            =   9720
         TabIndex        =   8
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox txtFRM 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   29
         Left            =   9720
         TabIndex        =   7
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox txtFRM 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   30
         Left            =   9720
         TabIndex        =   6
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label lblREC 
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
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
         Left            =   7920
         TabIndex        =   141
         Top             =   3120
         Width           =   2535
      End
      Begin VB.Label lblTIT 
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   8280
         TabIndex        =   133
         Top             =   120
         Width           =   2175
      End
      Begin VB.Label lblTIT 
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   5640
         TabIndex        =   132
         Top             =   120
         Width           =   2175
      End
      Begin VB.Label lblTIT 
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   3000
         TabIndex        =   131
         Top             =   120
         Width           =   2175
      End
      Begin VB.Label lblTIT 
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   130
         Top             =   120
         Width           =   2175
      End
      Begin VB.Label lblDAYS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   0
         Left            =   285
         TabIndex        =   129
         Top             =   480
         Width           =   90
      End
      Begin VB.Label lblDAYS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   1
         Left            =   285
         TabIndex        =   128
         Top             =   840
         Width           =   90
      End
      Begin VB.Label lblDAYS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   2
         Left            =   285
         TabIndex        =   127
         Top             =   1200
         Width           =   90
      End
      Begin VB.Label lblDAYS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   3
         Left            =   285
         TabIndex        =   126
         Top             =   1560
         Width           =   90
      End
      Begin VB.Label lblDAYS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   4
         Left            =   285
         TabIndex        =   125
         Top             =   1920
         Width           =   90
      End
      Begin VB.Label lblDAYS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   5
         Left            =   285
         TabIndex        =   124
         Top             =   2280
         Width           =   90
      End
      Begin VB.Label lblDAYS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   6
         Left            =   285
         TabIndex        =   123
         Top             =   2640
         Width           =   90
      End
      Begin VB.Label lblDAYS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   7
         Left            =   285
         TabIndex        =   122
         Top             =   3000
         Width           =   90
      End
      Begin VB.Label lblDAYS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   8
         Left            =   2925
         TabIndex        =   121
         Top             =   480
         Width           =   90
      End
      Begin VB.Label lblDAYS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   9
         Left            =   2925
         TabIndex        =   120
         Top             =   840
         Width           =   90
      End
      Begin VB.Label lblDAYS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   10
         Left            =   2925
         TabIndex        =   119
         Top             =   1200
         Width           =   90
      End
      Begin VB.Label lblDAYS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   11
         Left            =   2925
         TabIndex        =   118
         Top             =   1560
         Width           =   90
      End
      Begin VB.Label lblDAYS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   12
         Left            =   2925
         TabIndex        =   117
         Top             =   1920
         Width           =   90
      End
      Begin VB.Label lblDAYS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   13
         Left            =   2925
         TabIndex        =   116
         Top             =   2280
         Width           =   90
      End
      Begin VB.Label lblDAYS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   14
         Left            =   2925
         TabIndex        =   115
         Top             =   2640
         Width           =   90
      End
      Begin VB.Label lblDAYS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   15
         Left            =   2925
         TabIndex        =   114
         Top             =   3000
         Width           =   90
      End
      Begin VB.Label lblDAYS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   16
         Left            =   5565
         TabIndex        =   113
         Top             =   480
         Width           =   90
      End
      Begin VB.Label lblDAYS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   17
         Left            =   5565
         TabIndex        =   112
         Top             =   840
         Width           =   90
      End
      Begin VB.Label lblDAYS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   18
         Left            =   5565
         TabIndex        =   111
         Top             =   1200
         Width           =   90
      End
      Begin VB.Label lblDAYS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   19
         Left            =   5565
         TabIndex        =   110
         Top             =   1560
         Width           =   90
      End
      Begin VB.Label lblDAYS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   20
         Left            =   5565
         TabIndex        =   109
         Top             =   1920
         Width           =   90
      End
      Begin VB.Label lblDAYS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   21
         Left            =   5565
         TabIndex        =   108
         Top             =   2280
         Width           =   90
      End
      Begin VB.Label lblDAYS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   22
         Left            =   5565
         TabIndex        =   107
         Top             =   2640
         Width           =   90
      End
      Begin VB.Label lblDAYS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   23
         Left            =   5565
         TabIndex        =   106
         Top             =   3000
         Width           =   90
      End
      Begin VB.Label lblDAYS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   24
         Left            =   8205
         TabIndex        =   105
         Top             =   480
         Width           =   90
      End
      Begin VB.Label lblDAYS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   25
         Left            =   8205
         TabIndex        =   104
         Top             =   840
         Width           =   90
      End
      Begin VB.Label lblDAYS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   26
         Left            =   8205
         TabIndex        =   103
         Top             =   1200
         Width           =   90
      End
      Begin VB.Label lblDAYS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   27
         Left            =   8205
         TabIndex        =   102
         Top             =   1560
         Width           =   90
      End
      Begin VB.Label lblDAYS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   28
         Left            =   8205
         TabIndex        =   101
         Top             =   1920
         Width           =   90
      End
      Begin VB.Label lblDAYS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   29
         Left            =   8205
         TabIndex        =   100
         Top             =   2280
         Width           =   90
      End
      Begin VB.Label lblDAYS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   30
         Left            =   8205
         TabIndex        =   99
         Top             =   2640
         Width           =   90
      End
   End
   Begin VB.ListBox lstSIBG 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5820
      Left            =   120
      MultiSelect     =   2  'Extended
      TabIndex        =   4
      Top             =   480
      Width           =   10575
   End
   Begin VB.CommandButton cmdDEL 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   2640
      Picture         =   "EFFORT.frx":26F8
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6720
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
      Picture         =   "EFFORT.frx":297A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6720
      Width           =   735
   End
   Begin VB.CommandButton cmdPRINT 
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
      Left            =   960
      MousePointer    =   1  'Arrow
      Picture         =   "EFFORT.frx":476C
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6720
      Width           =   735
   End
   Begin VB.CommandButton cmdBACK 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   9840
      MousePointer    =   1  'Arrow
      Picture         =   "EFFORT.frx":49EE
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6720
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FFFF&
      Caption         =   " 08"
      Height          =   255
      Left            =   0
      TabIndex        =   143
      Top             =   7440
      Width           =   255
   End
   Begin VB.Label lblSELSIBG 
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Left            =   120
      TabIndex        =   137
      Top             =   120
      Width           =   10575
   End
   Begin VB.Label lblSELMN 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
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
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   135
      Top             =   120
      Width           =   4815
   End
End
Attribute VB_Name = "frmEFFORT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' TNJ Note: The code subroutines in this module have been ordered alphabetically

Private MNCODE(), MNAME(), MNCAL(), NMN, MNSEL()
Private SICODE(), SINAME(), SISEQ(), BGCODE(), BGNAME(), BGSEQ(), NSI, NBG
Private NAS, ASIC(), ASIN()
Private NFR, FRC(), FRN(), FRNO()
Private CURSIBG, CURSIBGC, OLDREC
Private OKMN(), OKSIBG()
Private SIMN()
Private Sub ASSO_SIMN()

Dim I, J, K, XXX, yyy, fnm, L

ReDim SIMN(1 To 10000)

For I = 1 To 10000
SIMN(I) = 0
Next I

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_ASSOSI.TXT"

Open fnm For Input As #1

Do Until EOF(1)

Line Input #1, XXX

J = Val(Mid(XXX, 1, 4)): K = Val(Mid(XXX, 37, 4))

For L = 1 To K
Line Input #1, yyy
I = Val(Mid(yyy, 6, 4)): SIMN(I) = J
Next L
  
Loop
  
Close #1

End Sub
Private Sub cmdBACK_Click()

lstSIBG.Visible = True

cmdBACK.MousePointer = 13

If Dir(APPROOT + "\ARTBAS\EFFORT\WORK.MDB") <> "" Then Kill APPROOT + "\ARTBAS\EFFORT\WORK.MDB"

Dim fnm, NN, XXX

NN = 0

fnm = APPROOT + "\ARTBAS\EFFORT\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_ESAMPLES.TXT"

If Dir(fnm) = "" Then GoTo CONT_RET

Open fnm For Input As #1

Do Until EOF(1)
Line Input #1, XXX
NN = NN + 1
Loop

Close #1

If NN = 0 Then Kill fnm
If NN = 1 And Left(XXX, 6) = "ZZZZZZ" Then Kill fnm

CONT_RET:

frmEFFORT.MousePointer = 13
Load frmARTB01
Unload frmEFFORT
frmARTB01.Show

End Sub
Private Sub cmdCONFIRM_Click()

cmdCONFIRM.MousePointer = 13

CURMNC = MNCODE(lstMINOR.ListIndex + 1)
CURMNN = MNAME(lstMINOR.ListIndex + 1)

frmEFFORT.Caption = monthtab(current_month) + " " + _
                    Format(current_year, "0000") + " - " + msgtab(33) + _
                    " - " + CURMNN

lblSELMN.Visible = False
lstMINOR.Visible = False
cmdCONFIRM.Visible = False

lstSIBG.Visible = True
lblSELSIBG.Visible = True

Call LOAD_ASSO
Call LOAD_FRAME

End Sub
Private Sub cmdDEL_Click()

txtREC.Text = "CANCELLED"

Dim response As Integer

Beep

response = MsgBox(msgtab(36), vbCritical + vbDefaultButton2 + vbOKCancel, " ")

If response = 2 Then Exit Sub

Dim dbn, XKEY, I

For I = 1 To CURCAL

txtACT(I - 1).Text = "": txtSMP(I - 1) = ""
txtFRM(I - 1).Text = ""

Next I

dbn = APPROOT + "\ARTBAS\EFFORT\WORK.MDB"

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(dbn)
Set prm_record = prm_database.OpenRecordset("ETAB")

With prm_record

.Index = "primarykey"

For I = 1 To CURCAL

XKEY = CURSIBGC + "+D" + Format(I, "00")

.Seek "=", XKEY

If .NoMatch = True Then GoTo CONT_LOOP

.Edit

![eact] = 0: ![esmp] = 0:  ![erec] = "???"

.Update

CONT_LOOP:

Next I

For I = 1 To CURCAL

txtACT(I - 1).Text = "": txtSMP(I - 1) = ""

Next I

End With

prm_record.Close
prm_database.Close

Call cmdEND_Click

End Sub

Private Sub cmdEND_Click()

lstSIBG.Visible = True

Dim resp

If Len(RTrim(txtREC.Text)) = 0 Or txtREC.Text = "???" Then
   txtREC.Text = "???"
   cmdSTAY.MousePointer = 1
   resp = MsgBox(msgtab(169), vbCritical + vbOKOnly, " ")
   Exit Sub
   End If

cmdSTAY.MousePointer = 13

cmdEND.MousePointer = 13

Call cmdSTAY_Click

cmdBACK.Visible = True

cmdEND.Visible = False
cmdPRINT.Visible = False
cmdSTAY.Visible = False
cmdDEL.Visible = False
cmdLEAVE.Visible = False

Frame1.Visible = False

lstSIBG.Enabled = True

NFR = 0

Call LOAD_FRAME

cmdEND.MousePointer = 1

End Sub

Private Sub cmdGUIDE1_Click()

HTYPE = "40"

If lstMINOR.Visible = False Then
   If Frame1.Visible = False Then HTYPE = "50"
   If Frame1.Visible = True Then HTYPE = "60"
   End If

HFNM = APPROOT + "\ARTBAS\HELP\" + current_language + "HELP" + HTYPE + ".rtf"

If Dir(HFNM) = "" Then Exit Sub

frmEFFORT.Enabled = False
Load frmGUIDE
frmGUIDE.Show

End Sub
Private Sub cmdLEAVE_Click()

lstSIBG.Visible = True

cmdLEAVE.MousePointer = 13

cmdBACK.Visible = True

cmdEND.Visible = False
cmdPRINT.Visible = False
cmdSTAY.Visible = False
cmdDEL.Visible = False
cmdLEAVE.Visible = False

Frame1.Visible = False

lstSIBG.Enabled = True

NFR = 0

Call LOAD_FRAME

cmdLEAVE.MousePointer = 1

End Sub
Private Sub cmdPRINT_Click()

Dim dbn

dbn = APPROOT + "\ARTBAS\EFFORT\WORK.MDB"

If Dir(dbn) = "" Then Exit Sub

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(dbn)
Set prm_record = prm_database.OpenRecordset("ETAB")

Printer.FontBold = True
Printer.FontName = "Courier"
Printer.FontName = "Courier New"
Printer.FontSize = 11

Dim I, J, pageno, lineno

pageno = 0

GoSub CHANGE_PAGE

With prm_record

.MoveFirst
.Index = "primarykey"

Do Until .EOF

If ![emnc] <> CURMNC Then GoTo CONT_READ
If Left(![ekey], 11) <> CURSIBGC Then GoTo CONT_READ

Printer.Print Tab(5); Right(![ekey], 2); _
              Tab(10); Right(Space(15) + LTrim(Format(![eact], "#####0.000")), 15); _
              Tab(25); Right(Space(15) + LTrim(Format(![esmp], "#####0.000")), 15); _
              Tab(40); Right(Space(15) + LTrim(Format(![efrm], "#####0.000")), 15)
                            
lineno = lineno + 1

If lineno > 55 Then GoSub CHANGE_PAGE

CONT_READ:

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

Printer.Print Tab(5); frmEFFORT.Caption

Printer.Print

Printer.Print Tab(5); CURSIBG
Printer.Print Tab(5); String(70, "-")

Printer.Print
Printer.Print Tab(5); msgtab(82) + ": " + txtREC.Text

Printer.Print

Printer.Print Tab(5); msgtab(83); _
              Tab(10); Right(Space(15) + RTrim(msgtab(84)), 15); _
              Tab(25); Right(Space(15) + RTrim(msgtab(85)), 15); _
              Tab(40); Right(Space(15) + RTrim(msgtab(86)), 15)
             
Printer.Print

Return

End Sub
Private Sub cmdSTAY_Click()

Dim resp

cmdSTAY.MousePointer = 13

If Len(RTrim(txtREC.Text)) = 0 Or txtREC.Text = "???" Then
   txtREC.Text = "???"
   cmdSTAY.MousePointer = 1
   resp = MsgBox(msgtab(169), vbCritical + vbOKOnly, " ")
   Exit Sub
   End If

Dim fnm, dbn, XKEY1, XKEY2, I, J, K

fnm = APPROOT + "\ARTBAS\EFFORT\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_ESAMPLES.TXT"

XKEY1 = FRC(lstSIBG.ListIndex + 1)

dbn = APPROOT + "\ARTBAS\EFFORT\WORK.MDB"

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(dbn)
Set prm_record = prm_database.OpenRecordset("ETAB")

With prm_record

.Index = "primarykey"

For I = 1 To CURCAL

If Len(txtACT(I - 1).Text) = 0 Then txtACT(I - 1).Text = 0
If Len(txtSMP(I - 1).Text) = 0 Then txtSMP(I - 1).Text = 0

XKEY2 = XKEY1 + "+D" + Format(I, "00")

.Seek "=", XKEY2

If .NoMatch = True Then

   .AddNew
   
   ![ekey] = XKEY2
   ![emnc] = CURMNC
   ![eact] = CDbl(txtACT(I - 1).Text)
   ![esmp] = CDbl(txtSMP(I - 1).Text)
   ![efrm] = FRNO(lstSIBG.ListIndex + 1)
   ![erec] = txtREC.Text
     
   .Update
   
   GoTo CONT_LOOP
   
   End If

.Edit

![emnc] = CURMNC
![eact] = CDbl(txtACT(I - 1).Text)
![esmp] = CDbl(txtSMP(I - 1).Text)
![efrm] = FRNO(lstSIBG.ListIndex + 1)
![erec] = txtREC.Text
   
.Update

CONT_LOOP:

Next I

Open fnm For Output As #1

.MoveFirst

Do Until .EOF

If ![ekey] = "ZZZZZZZZZZZZZZZ" Then GoTo CONT_READ
If ![eact] = 0 And ![esmp] = 0 Then GoTo CONT_READ
   
   Print #1, ![ekey] + " " + _
             Format(![emnc], "0000") + " " + _
             Format(![eact], "000000.000") + " " + _
             Format(![esmp], "000000.000") + " " + _
             Format(![efrm], "000000.000") + " " + ![erec]
             
             OKMN(![emnc]) = "+"
        
CONT_READ:
     
.MoveNext

Loop

End With

Close #1

prm_record.Close
prm_database.Close

For I = 1 To CURCAL

If Len(txtACT(I - 1).Text) <> 0 And txtACT(I - 1).Text = 0 Then txtACT(I - 1).Text = ""
If Len(txtSMP(I - 1).Text) <> 0 And txtSMP(I - 1).Text = 0 Then txtSMP(I - 1).Text = ""

If Len(txtACT(I - 1).Text) = 0 Then
   If Len(txtSMP(I - 1).Text) <> 0 Then txtACT(I - 1).Text = 0
   End If

Next I

cmdSTAY.MousePointer = 1

End Sub
Private Sub DELETE_ESTIM()

Dim fnm

fnm = APPROOT + "\ARTBAS\RESULTS\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "*.*"

If Dir(fnm) <> "" Then Kill fnm

End Sub
Private Sub Form_Load()

Set Picture = LoadPicture(APPROOT + "\ARTBAS\PICS_RUNTIME\SCREEN_08.JPG")

Call ASSO_SIMN

ReDim OKMN(1 To 10000)

Dim cal1, cal2, I, fnm

For I = 1 To 10000
OKMN(I) = "-"
Next I

fnm = APPROOT + "\ARTBAS\EFFORT\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_ESAMPLES.TXT"

If Dir(fnm) <> "" Then Call SETUP_EFF

If Dir(fnm) = "" Then Call SETUP_NEWEFF

frmEFFORT.Caption = monthtab(current_month) + " " + _
                    Format(current_year, "0000") + " - " + msgtab(33)

lstSIBG.Visible = False
lblSELSIBG.Visible = False
Frame1.Visible = False

cmdEND.Visible = False
cmdPRINT.Visible = False
cmdDEL.Visible = False
cmdSTAY.Visible = False
cmdLEAVE.Visible = False

cmdBACK.ToolTipText = msgtab(49)
cmdEND.ToolTipText = msgtab(51)
cmdPRINT.ToolTipText = msgtab(52)
cmdCONFIRM.ToolTipText = msgtab(36)

cmdEND.ToolTipText = msgtab(77)
cmdPRINT.ToolTipText = msgtab(78)
cmdSTAY.ToolTipText = msgtab(79)
cmdDEL.ToolTipText = msgtab(80)
cmdLEAVE.ToolTipText = msgtab(81)
cmdGUIDE1.ToolTipText = msgtab(243)
lblSELMN.Caption = msgtab(72)
lblSELSIBG.Caption = msgtab(73)
lblREC.Caption = msgtab(82)

For I = 0 To 30

If I + 1 > CURCAL Then

txtACT(I).Visible = False
txtSMP(I).Visible = False
txtFRM(I).Visible = False
lblDAYS(I).Visible = False

End If

Next I

For I = 0 To 30

lblDAYS(I).Caption = I + 1
txtFRM(I).Enabled = False

Next I

For I = 0 To 3
lblTIT(I).Caption = msgtab(71)
Next I

For I = 0 To CURCAL - 1
txtACT(I).TabIndex = 2 * (I + 1) - 2
txtSMP(I).TabIndex = 2 * (I + 1) - 1
Next I

Call LOAD_STRATA

End Sub
Private Sub LOAD_ASSO()

Dim I, J, K, XXX, yyy, fnm

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_ASSOSI.TXT"

Open fnm For Input As #1

Do Until EOF(1)

Line Input #1, XXX

If CURMNC <> CDbl(Mid(XXX, 1, 4)) Then
   K = CDbl(Mid(XXX, 37, 4))
   For I = 1 To K
   Line Input #1, yyy
   Next I
   GoTo CONT_LOOP
   End If

NAS = CDbl(Mid(XXX, 37, 4))

ReDim ASIC(1 To NAS), ASIN(1 To NAS)

For I = 1 To NAS
Line Input #1, yyy
J = CDbl(Mid(yyy, 6, 4))
ASIC(I) = J: ASIN(I) = Mid(yyy, 11, 30)
Next I

Close #1

Exit Sub

CONT_LOOP:

Loop

End Sub
Private Sub LOAD_BG()

Dim J, K, XXX, yyy, fnm

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_BG.TXT"

Open fnm For Input As #1

NBG = 0

Do Until EOF(1)

Line Input #1, XXX

NBG = NBG + 1

ReDim Preserve BGCODE(1 To NBG), BGNAME(1 To NBG)

BGCODE(NBG) = CDbl(Left(XXX, 4))
BGNAME(NBG) = Mid(XXX, 6, 30)

Loop

Close #1

End Sub
Private Sub LOAD_FRAME()

lblSELSIBG.Caption = msgtab(73)

Call LOAD_BG

Dim SF(), SFN(), BF()

ReDim SF(1 To 10000), BF(1 To 10000), SFN(1 To 10000)

Dim I

For I = 1 To NBG
BF(BGCODE(I)) = BGNAME(I)
Next I

For I = 1 To 10000
SF(I) = 0
Next I

For I = 1 To NAS
SF(ASIC(I)) = ASIC(I): SFN(ASIC(I)) = ASIN(I)
Next I

Dim J, K, XXX, yyy, fnm

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_FRAME.TXT"

Open fnm For Input As #1

NFR = 0

Do Until EOF(1)

Line Input #1, XXX

K = CDbl(Mid(XXX, 2, 4))

If SF(K) = 0 Then GoTo CONT_LOOP
If CDbl(Mid(XXX, 26, 15)) = 0 Then GoTo CONT_LOOP

NFR = NFR + 1

ReDim Preserve FRC(1 To NFR), FRN(1 To NFR), FRNO(1 To NFR), OKSIBG(1 To NFR)

FRC(NFR) = Left(XXX, 11)
FRNO(NFR) = CDbl(Mid(XXX, 26, 15))
OKSIBG(NFR) = "-"

K = CDbl(Mid(XXX, 2, 4))
J = SF(K)
K = CDbl(Mid(XXX, 8, 4))

FRN(NFR) = SFN(J) + " " + BF(K)

CONT_LOOP:

Loop

Close #1

If NFR = 0 Then Exit Sub

fnm = APPROOT + "\ARTBAS\EFFORT\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_ESAMPLES.TXT"

If Dir(fnm) = "" Then GoTo NO_ACTION

Open fnm For Input As #1

Do Until EOF(1)

Line Input #1, XXX

XXX = Left(XXX, 11)

For I = 1 To NFR
If XXX = FRC(I) Then
   OKSIBG(I) = "+"
   GoTo CONT_READ
   End If
Next I

CONT_READ:

Loop

Close #1

NO_ACTION:

lstSIBG.Clear

For I = 1 To NFR
lstSIBG.AddItem FRN(I) + " " + OKSIBG(I)
Next I

lstSIBG.ItemData(lstSIBG.NewIndex) = I

End Sub
Private Sub LOAD_STRATA()

Dim I, XXX, fnm

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_MINOR.TXT"

NMN = 0

lstMINOR.Clear

Open fnm For Input As #1

Do Until EOF(1)

Line Input #1, XXX

NMN = NMN + 1

ReDim Preserve MNAME(1 To NMN), MNCODE(1 To NMN), MNCAL(1 To NMN), MNSEL(1 To NMN)

MNCODE(NMN) = Val(Mid(XXX, 1, 4))
MNAME(NMN) = Mid(XXX, 6, 30)
MNCAL(NMN) = Mid(XXX, 37, 31)

I = NMN

lstMINOR.AddItem MNAME(I) + " " + OKMN(MNCODE(NMN))
MNSEL(I) = 0
lstMINOR.ItemData(lstMINOR.NewIndex) = I

Loop

Close #1

lstMINOR.ListIndex = 0

End Sub
Private Sub lstSIBG_Click()

lstSIBG.Visible = False

Call DELETE_ESTIM

Dim I, fnm, dbn

I = lstSIBG.ListIndex

CURSIBG = FRN(I + 1): CURSIBGC = FRC(I + 1)

lblSELSIBG.Caption = RTrim(msgtab(73)) + " : " + _
                   RTrim(Left(CURSIBG, 30)) + " + " + _
                   LTrim(Right(CURSIBG, 30))

lstSIBG.Enabled = False
Frame1.Visible = True

cmdBACK.Visible = False

cmdEND.Visible = True
cmdPRINT.Visible = True
cmdSTAY.Visible = True
cmdDEL.Visible = True
cmdLEAVE.Visible = True

fnm = APPROOT + "\ARTBAS\EFFORT\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_ESAMPLES.TXT"

If Dir(fnm) <> "" Then GoTo CR_FROM_FILE

For I = 1 To CURCAL

txtACT(I - 1).Text = "": txtSMP(I - 1).Text = ""
txtFRM(I - 1).Text = Format(FRNO(lstSIBG.ListIndex + 1), "#######0")

Next I

Exit Sub

CR_FROM_FILE:

dbn = APPROOT + "\ARTBAS\EFFORT\WORK.MDB"

Dim prm_database As Database, prm_record As Recordset, XKEY

Set prm_database = OpenDatabase(dbn)
Set prm_record = prm_database.OpenRecordset("ETAB")

With prm_record

.Index = "primarykey"

For I = 1 To CURCAL

txtACT(I - 1).Text = "": txtSMP(I - 1).Text = ""
txtFRM(I - 1).Text = Format(FRNO(lstSIBG.ListIndex + 1), "#######0")

Next I

For I = 1 To CURCAL

XKEY = CURSIBGC + "+D" + Format(I, "00")

.Seek "=", XKEY

If .NoMatch = True Then GoTo CONT_LOOP

txtACT(I - 1).Text = ![eact]: txtSMP(I - 1).Text = ![esmp]
txtFRM(I - 1).Text = Format(FRNO(lstSIBG.ListIndex + 1), "#######0")

If Len(txtACT(I - 1).Text) <> 0 And txtACT(I - 1).Text = 0 Then txtACT(I - 1).Text = ""
If Len(txtSMP(I - 1).Text) <> 0 And txtSMP(I - 1).Text = 0 Then txtSMP(I - 1).Text = ""

If Len(txtACT(I - 1).Text) = 0 Then
   If Len(txtSMP(I - 1).Text) <> 0 Then txtACT(I - 1).Text = 0
   End If

If Len(RTrim(![erec])) <> 0 And Left(![erec], 3) <> "???" Then
   txtREC.Text = ![erec]
   End If
   
CONT_LOOP:

Next I

If txtREC.Text = "CANCELLED" Then txtREC.Text = "???"

txtREC.Refresh

prm_record.Close
prm_database.Close

End With

End Sub
Private Sub txtACT_Change(Index As Integer)

If IsNumeric(txtACT(Index).Text) = False Then txtACT(Index).Text = ""

If IsNumeric(txtACT(Index).Text) = True Then
   If txtACT(Index).Text < 0 Then txtACT(Index).Text = -txtACT(Index).Text
   End If

End Sub
Private Sub txtACT_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

If Len(txtACT(Index).Text) = 0 Then
   txtSMP(Index).Text = ""
   Exit Sub
   End If

If IsNumeric(txtACT(Index).Text) = True And txtACT(Index).Text = 0 Then
   txtSMP(Index).Text = txtFRM(Index).Text
   End If

End Sub
Private Sub txtREC_Change()

Dim resp

If Len(txtREC.Text) > 15 Then
Beep

resp = MsgBox("Max 15 chrs", vbOKOnly)

txtREC.Text = Left(txtREC.Text, 15)
'txtREC.Text = ""
Exit Sub
End If

End Sub
Private Sub txtSMP_Change(Index As Integer)

If IsNumeric(txtSMP(Index).Text) = False Then txtSMP(Index).Text = ""
   
If IsNumeric(txtSMP(Index).Text) = True Then
   If txtSMP(Index).Text < 0 Then txtSMP(Index).Text = -txtSMP(Index).Text
   End If
   
End Sub
Private Sub SETUP_EFF()

Dim fnm, dbn, XXX, xcode, xmnc, xact, xsmp, xfrm, xrec

fnm = APPROOT + "\ARTBAS\EFFORT\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_ESAMPLES.TXT"

Open fnm For Input As #1

FileCopy APPROOT + "\ARTBAS\STRUS\EFFORT.MDB", APPROOT + "\ARTBAS\EFFORT\WORK.MDB"

dbn = APPROOT + "\ARTBAS\EFFORT\WORK.MDB"

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(dbn)
Set prm_record = prm_database.OpenRecordset("ETAB")

With prm_record

.Index = "primarykey"

Do Until EOF(1)

Line Input #1, XXX

.AddNew

![ekey] = Left(XXX, 15)

xmnc = Val(Mid(XXX, 2, 4))

![emnc] = SIMN(xmnc)
![eact] = CDbl(Mid(XXX, 22, 10))
![esmp] = CDbl(Mid(XXX, 33, 10))
![efrm] = CDbl(Mid(XXX, 44, 10))
![erec] = Mid(XXX, 55, 15)

If ![eact] <> 0 And ![esmp] = 0 Then ![esmp] = ![efrm]

OKMN(![emnc]) = "+"

.Update

Loop

Close #1

End With

prm_record.Close
prm_database.Close

End Sub
Private Sub SETUP_NEWEFF()

Dim fnm, dbn

fnm = APPROOT + "\ARTBAS\EFFORT\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_ESAMPLES.TXT"

Open fnm For Output As #1

dbn = APPROOT + "\ARTBAS\STRUS\EFFORT.MDB"

FileCopy dbn, APPROOT + "\ARTBAS\EFFORT\WORK.MDB"

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(dbn)
Set prm_record = prm_database.OpenRecordset("ETAB")

With prm_record

.Index = "primarykey"

.MoveFirst

Do Until .EOF

.Edit
![emnc] = CURMNC
.Update

Print #1, ![ekey], ![emnc], ![eact], ![esmp], ![efrm], ![erec]

.MoveNext

Loop

Close #1

End With

prm_record.Close
prm_database.Close

End Sub
