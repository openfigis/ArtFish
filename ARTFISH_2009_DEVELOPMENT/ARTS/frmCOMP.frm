VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmCOMP 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7140
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   10185
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   3  'I-Beam
   Moveable        =   0   'False
   ScaleHeight     =   7140
   ScaleWidth      =   10185
   StartUpPosition =   2  'CenterScreen
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
      Left            =   8520
      MousePointer    =   1  'Arrow
      Picture         =   "frmCOMP.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   109
      Top             =   6240
      Width           =   735
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      ForeColor       =   &H80000008&
      Height          =   4095
      Left            =   120
      TabIndex        =   23
      Top             =   240
      Width           =   4815
      Begin VB.OptionButton optQUIT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
         Caption         =   "Option1"
         Height          =   255
         Left            =   2640
         TabIndex        =   108
         Top             =   3720
         Width           =   2055
      End
      Begin VB.OptionButton optV 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Option1"
         Height          =   255
         Index           =   12
         Left            =   2280
         TabIndex        =   106
         Top             =   3720
         Width           =   255
      End
      Begin VB.OptionButton optV 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Option1"
         Height          =   255
         Index           =   11
         Left            =   2280
         TabIndex        =   105
         Top             =   3360
         Width           =   255
      End
      Begin VB.OptionButton optV 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Option1"
         Height          =   255
         Index           =   10
         Left            =   2280
         TabIndex        =   104
         Top             =   3120
         Width           =   255
      End
      Begin VB.OptionButton optV 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Option1"
         Height          =   255
         Index           =   9
         Left            =   2280
         TabIndex        =   103
         Top             =   2880
         Width           =   255
      End
      Begin VB.OptionButton optV 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Option1"
         Height          =   255
         Index           =   8
         Left            =   2280
         TabIndex        =   102
         Top             =   2640
         Width           =   255
      End
      Begin VB.OptionButton optV 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Option1"
         Height          =   255
         Index           =   7
         Left            =   2280
         TabIndex        =   101
         Top             =   2400
         Width           =   255
      End
      Begin VB.OptionButton optV 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Option1"
         Height          =   255
         Index           =   6
         Left            =   2280
         TabIndex        =   100
         Top             =   2160
         Width           =   255
      End
      Begin VB.OptionButton optV 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Option1"
         Height          =   255
         Index           =   5
         Left            =   2280
         TabIndex        =   99
         Top             =   1920
         Width           =   255
      End
      Begin VB.OptionButton optV 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Option1"
         Height          =   255
         Index           =   4
         Left            =   2280
         TabIndex        =   98
         Top             =   1680
         Width           =   255
      End
      Begin VB.OptionButton optV 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Option1"
         Height          =   255
         Index           =   3
         Left            =   2280
         TabIndex        =   97
         Top             =   1440
         Width           =   255
      End
      Begin VB.OptionButton optV 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Option1"
         Height          =   255
         Index           =   2
         Left            =   2280
         TabIndex        =   96
         Top             =   1200
         Width           =   255
      End
      Begin VB.OptionButton optV 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Option1"
         Height          =   255
         Index           =   1
         Left            =   2280
         TabIndex        =   95
         Top             =   960
         Width           =   255
      End
      Begin VB.OptionButton optV 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Option1"
         Height          =   255
         Index           =   0
         Left            =   2280
         TabIndex        =   94
         Top             =   720
         Width           =   255
      End
      Begin VB.OptionButton optP 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Option1"
         Height          =   255
         Index           =   12
         Left            =   1920
         TabIndex        =   93
         Top             =   3720
         Width           =   255
      End
      Begin VB.OptionButton optP 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Option1"
         Height          =   255
         Index           =   11
         Left            =   1920
         TabIndex        =   92
         Top             =   3360
         Width           =   255
      End
      Begin VB.OptionButton optP 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Option1"
         Height          =   255
         Index           =   10
         Left            =   1920
         TabIndex        =   91
         Top             =   3120
         Width           =   255
      End
      Begin VB.OptionButton optP 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Option1"
         Height          =   255
         Index           =   9
         Left            =   1920
         TabIndex        =   90
         Top             =   2880
         Width           =   255
      End
      Begin VB.OptionButton optP 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Option1"
         Height          =   255
         Index           =   8
         Left            =   1920
         TabIndex        =   89
         Top             =   2640
         Width           =   255
      End
      Begin VB.OptionButton optP 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Option1"
         Height          =   255
         Index           =   7
         Left            =   1920
         TabIndex        =   88
         Top             =   2400
         Width           =   255
      End
      Begin VB.OptionButton optP 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Option1"
         Height          =   255
         Index           =   6
         Left            =   1920
         TabIndex        =   87
         Top             =   2160
         Width           =   255
      End
      Begin VB.OptionButton optP 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Option1"
         Height          =   255
         Index           =   5
         Left            =   1920
         TabIndex        =   86
         Top             =   1920
         Width           =   255
      End
      Begin VB.OptionButton optP 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Option1"
         Height          =   255
         Index           =   4
         Left            =   1920
         TabIndex        =   85
         Top             =   1680
         Width           =   255
      End
      Begin VB.OptionButton optP 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Option1"
         Height          =   255
         Index           =   3
         Left            =   1920
         TabIndex        =   84
         Top             =   1440
         Width           =   255
      End
      Begin VB.OptionButton optP 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Option1"
         Height          =   255
         Index           =   2
         Left            =   1920
         TabIndex        =   83
         Top             =   1200
         Width           =   255
      End
      Begin VB.OptionButton optP 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Option1"
         Height          =   255
         Index           =   1
         Left            =   1920
         TabIndex        =   82
         Top             =   960
         Width           =   255
      End
      Begin VB.OptionButton optP 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Option1"
         Height          =   255
         Index           =   0
         Left            =   1920
         TabIndex        =   81
         Top             =   720
         Width           =   255
      End
      Begin VB.OptionButton optU 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Option1"
         Height          =   255
         Index           =   12
         Left            =   1560
         TabIndex        =   80
         Top             =   3720
         Width           =   255
      End
      Begin VB.OptionButton optU 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Option1"
         Height          =   255
         Index           =   11
         Left            =   1560
         TabIndex        =   79
         Top             =   3360
         Width           =   255
      End
      Begin VB.OptionButton optU 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Option1"
         Height          =   255
         Index           =   10
         Left            =   1560
         TabIndex        =   78
         Top             =   3120
         Width           =   255
      End
      Begin VB.OptionButton optU 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Option1"
         Height          =   255
         Index           =   9
         Left            =   1560
         TabIndex        =   77
         Top             =   2880
         Width           =   255
      End
      Begin VB.OptionButton optU 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Option1"
         Height          =   255
         Index           =   8
         Left            =   1560
         TabIndex        =   76
         Top             =   2640
         Width           =   255
      End
      Begin VB.OptionButton optU 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Option1"
         Height          =   255
         Index           =   7
         Left            =   1560
         TabIndex        =   75
         Top             =   2400
         Width           =   255
      End
      Begin VB.OptionButton optU 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Option1"
         Height          =   255
         Index           =   6
         Left            =   1560
         TabIndex        =   74
         Top             =   2160
         Width           =   255
      End
      Begin VB.OptionButton optU 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Option1"
         Height          =   255
         Index           =   5
         Left            =   1560
         TabIndex        =   73
         Top             =   1920
         Width           =   255
      End
      Begin VB.OptionButton optU 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Option1"
         Height          =   255
         Index           =   4
         Left            =   1560
         TabIndex        =   72
         Top             =   1680
         Width           =   255
      End
      Begin VB.OptionButton optU 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Option1"
         Height          =   255
         Index           =   3
         Left            =   1560
         TabIndex        =   71
         Top             =   1440
         Width           =   255
      End
      Begin VB.OptionButton optU 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Option1"
         Height          =   255
         Index           =   2
         Left            =   1560
         TabIndex        =   70
         Top             =   1200
         Width           =   255
      End
      Begin VB.OptionButton optU 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Option1"
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   69
         Top             =   960
         Width           =   255
      End
      Begin VB.OptionButton optU 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Option1"
         Height          =   255
         Index           =   0
         Left            =   1560
         TabIndex        =   68
         Top             =   720
         Width           =   255
      End
      Begin VB.OptionButton optE 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Option1"
         Height          =   255
         Index           =   12
         Left            =   1200
         TabIndex        =   67
         Top             =   3720
         Width           =   255
      End
      Begin VB.OptionButton optE 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Option1"
         Height          =   255
         Index           =   11
         Left            =   1200
         TabIndex        =   66
         Top             =   3360
         Width           =   255
      End
      Begin VB.OptionButton optE 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Option1"
         Height          =   255
         Index           =   10
         Left            =   1200
         TabIndex        =   65
         Top             =   3120
         Width           =   255
      End
      Begin VB.OptionButton optE 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Option1"
         Height          =   255
         Index           =   9
         Left            =   1200
         TabIndex        =   64
         Top             =   2880
         Width           =   255
      End
      Begin VB.OptionButton optE 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Option1"
         Height          =   255
         Index           =   8
         Left            =   1200
         TabIndex        =   63
         Top             =   2640
         Width           =   255
      End
      Begin VB.OptionButton optE 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Option1"
         Height          =   255
         Index           =   7
         Left            =   1200
         TabIndex        =   62
         Top             =   2400
         Width           =   255
      End
      Begin VB.OptionButton optE 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Option1"
         Height          =   255
         Index           =   6
         Left            =   1200
         TabIndex        =   61
         Top             =   2160
         Width           =   255
      End
      Begin VB.OptionButton optE 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Option1"
         Height          =   255
         Index           =   5
         Left            =   1200
         TabIndex        =   60
         Top             =   1920
         Width           =   255
      End
      Begin VB.OptionButton optE 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Option1"
         Height          =   255
         Index           =   4
         Left            =   1200
         TabIndex        =   59
         Top             =   1680
         Width           =   255
      End
      Begin VB.OptionButton optE 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Option1"
         Height          =   255
         Index           =   3
         Left            =   1200
         TabIndex        =   58
         Top             =   1440
         Width           =   255
      End
      Begin VB.OptionButton optE 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Option1"
         Height          =   255
         Index           =   2
         Left            =   1200
         TabIndex        =   57
         Top             =   1200
         Width           =   255
      End
      Begin VB.OptionButton optE 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Option1"
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   56
         Top             =   960
         Width           =   255
      End
      Begin VB.OptionButton optE 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Option1"
         Height          =   255
         Index           =   0
         Left            =   1200
         TabIndex        =   55
         Top             =   720
         Width           =   255
      End
      Begin VB.OptionButton optC 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Option1"
         Height          =   255
         Index           =   12
         Left            =   840
         TabIndex        =   54
         Top             =   3720
         Width           =   255
      End
      Begin VB.OptionButton optC 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Option1"
         Height          =   255
         Index           =   11
         Left            =   840
         TabIndex        =   53
         Top             =   3360
         Width           =   255
      End
      Begin VB.OptionButton optC 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Option1"
         Height          =   255
         Index           =   10
         Left            =   840
         TabIndex        =   52
         Top             =   3120
         Width           =   255
      End
      Begin VB.OptionButton optC 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Option1"
         Height          =   255
         Index           =   9
         Left            =   840
         TabIndex        =   51
         Top             =   2880
         Width           =   255
      End
      Begin VB.OptionButton optC 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Option1"
         Height          =   255
         Index           =   8
         Left            =   840
         TabIndex        =   50
         Top             =   2640
         Width           =   255
      End
      Begin VB.OptionButton optC 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Option1"
         Height          =   255
         Index           =   7
         Left            =   840
         TabIndex        =   49
         Top             =   2400
         Width           =   255
      End
      Begin VB.OptionButton optC 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Option1"
         Height          =   255
         Index           =   6
         Left            =   840
         TabIndex        =   48
         Top             =   2160
         Width           =   255
      End
      Begin VB.OptionButton optC 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Option1"
         Height          =   255
         Index           =   5
         Left            =   840
         TabIndex        =   47
         Top             =   1920
         Width           =   255
      End
      Begin VB.OptionButton optC 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Option1"
         Height          =   255
         Index           =   4
         Left            =   840
         TabIndex        =   46
         Top             =   1680
         Width           =   255
      End
      Begin VB.OptionButton optC 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Option1"
         Height          =   255
         Index           =   3
         Left            =   840
         TabIndex        =   45
         Top             =   1440
         Width           =   255
      End
      Begin VB.OptionButton optC 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Option1"
         Height          =   255
         Index           =   2
         Left            =   840
         TabIndex        =   44
         Top             =   1200
         Width           =   255
      End
      Begin VB.OptionButton optC 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Option1"
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   43
         Top             =   960
         Width           =   255
      End
      Begin VB.OptionButton optC 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Option1"
         Height          =   255
         Index           =   0
         Left            =   840
         TabIndex        =   42
         Top             =   720
         Width           =   255
      End
      Begin VB.Label lblRANK 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ranking criteria"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   2520
         TabIndex        =   107
         Top             =   720
         Width           =   2175
         WordWrap        =   -1  'True
      End
      Begin VB.Line Line11 
         X1              =   4680
         X2              =   4680
         Y1              =   240
         Y2              =   480
      End
      Begin VB.Line Line10 
         X1              =   2400
         X2              =   4680
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Line Line9 
         X1              =   2400
         X2              =   2400
         Y1              =   480
         Y2              =   600
      End
      Begin VB.Line Line8 
         X1              =   3480
         X2              =   3480
         Y1              =   240
         Y2              =   360
      End
      Begin VB.Line Line7 
         X1              =   2040
         X2              =   3480
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Line Line6 
         X1              =   2040
         X2              =   2040
         Y1              =   600
         Y2              =   360
      End
      Begin VB.Line Line5 
         X1              =   2640
         X2              =   1680
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line Line4 
         X1              =   1680
         X2              =   1680
         Y1              =   600
         Y2              =   240
      End
      Begin VB.Line Line3 
         X1              =   1320
         X2              =   1320
         Y1              =   240
         Y2              =   600
      End
      Begin VB.Line Line2 
         X1              =   960
         X2              =   960
         Y1              =   240
         Y2              =   600
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   960
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Label lblV 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Value"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   4320
         TabIndex        =   41
         Top             =   0
         Width           =   855
      End
      Begin VB.Label lblP 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Price"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   3120
         TabIndex        =   40
         Top             =   0
         Width           =   855
      End
      Begin VB.Label lblU 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "CPUE"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   2160
         TabIndex        =   39
         Top             =   0
         Width           =   615
      End
      Begin VB.Label lblE 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Effort"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   1200
         TabIndex        =   38
         Top             =   0
         Width           =   735
      End
      Begin VB.Label lblC 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Catch"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   0
         TabIndex        =   37
         Top             =   0
         Width           =   975
      End
      Begin VB.Label lblMO 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   36
         Top             =   3720
         Width           =   495
      End
      Begin VB.Label lblMO 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "12"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   11
         Left            =   360
         TabIndex        =   35
         Top             =   3360
         Width           =   255
      End
      Begin VB.Label lblMO 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "11"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   10
         Left            =   360
         TabIndex        =   34
         Top             =   3120
         Width           =   255
      End
      Begin VB.Label lblMO 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   9
         Left            =   360
         TabIndex        =   33
         Top             =   2880
         Width           =   255
      End
      Begin VB.Label lblMO 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "09"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   8
         Left            =   360
         TabIndex        =   32
         Top             =   2640
         Width           =   255
      End
      Begin VB.Label lblMO 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "08"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   7
         Left            =   360
         TabIndex        =   31
         Top             =   2400
         Width           =   255
      End
      Begin VB.Label lblMO 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "07"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   6
         Left            =   360
         TabIndex        =   30
         Top             =   2160
         Width           =   255
      End
      Begin VB.Label lblMO 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "06"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   29
         Top             =   1920
         Width           =   255
      End
      Begin VB.Label lblMO 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "05"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   28
         Top             =   1680
         Width           =   255
      End
      Begin VB.Label lblMO 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "04"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   27
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label lblMO 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "03"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   26
         Top             =   1200
         Width           =   255
      End
      Begin VB.Label lblMO 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "02"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   25
         Top             =   960
         Width           =   255
      End
      Begin VB.Label lblMO 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "01"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   24
         Top             =   720
         Width           =   255
      End
   End
   Begin VB.CommandButton cmdRANK 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   1800
      Picture         =   "frmCOMP.frx":2262
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   6240
      Width           =   735
   End
   Begin VB.CommandButton cmdREP 
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
      MousePointer    =   1  'Arrow
      Picture         =   "frmCOMP.frx":285C
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   6240
      Width           =   735
   End
   Begin VB.CommandButton cmdGROUP 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   960
      Picture         =   "frmCOMP.frx":2ADE
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6240
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4095
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   3615
      Begin VB.OptionButton optMN1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Option1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   3375
      End
      Begin VB.OptionButton optMN2 
         BackColor       =   &H0080FF80&
         Caption         =   "Option1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   3375
      End
      Begin VB.OptionButton optMN3 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Option1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   3375
      End
      Begin VB.OptionButton optMN4 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Option1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   3375
      End
      Begin VB.OptionButton optMJ1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Option1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1680
         Width           =   3375
      End
      Begin VB.OptionButton optMJ2 
         BackColor       =   &H0080FF80&
         Caption         =   "Option1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1920
         Width           =   3375
      End
      Begin VB.OptionButton optMJ3 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Option1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   2160
         Width           =   3375
      End
      Begin VB.OptionButton optMJ4 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Option1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   2400
         Width           =   3375
      End
      Begin VB.OptionButton optGT1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Option1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   3000
         Width           =   3375
      End
      Begin VB.OptionButton optGT2 
         BackColor       =   &H0080FF80&
         Caption         =   "Option1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   3240
         Width           =   3375
      End
      Begin VB.OptionButton optGT3 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Option1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   3480
         Width           =   3375
      End
      Begin VB.OptionButton optGT4 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Option1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   3720
         Width           =   3375
      End
      Begin VB.Label lblMN 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   120
         Width           =   3375
      End
      Begin VB.Label lblMJ 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1440
         Width           =   3375
      End
      Begin VB.Label lblGT 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   2760
         Width           =   3375
      End
   End
   Begin VB.CommandButton cmdEXIT 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   9720
      MousePointer    =   1  'Arrow
      Picture         =   "frmCOMP.frx":2D60
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6600
      Width           =   375
   End
   Begin VB.CommandButton cmdRETURN 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   9360
      MousePointer    =   1  'Arrow
      Picture         =   "frmCOMP.frx":2FE2
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6240
      Width           =   735
   End
   Begin RichTextLib.RichTextBox rtsSEL 
      Height          =   4110
      Left            =   5040
      TabIndex        =   19
      Top             =   240
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   7250
      _Version        =   393217
      BackColor       =   12648447
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      MousePointer    =   1
      TextRTF         =   $"frmCOMP.frx":3264
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComctlLib.ProgressBar pgbCOMP 
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   4440
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      Height          =   255
      Left            =   2640
      TabIndex        =   111
      Top             =   6720
      Width           =   5775
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FFFF&
      Caption         =   "03"
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
      TabIndex        =   110
      Top             =   6960
      Width           =   255
   End
End
Attribute VB_Name = "frmCOMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WCOMP, WLEV, resp, WELM, WPER, WMY, ADDFISH, ADDFISH1, ADDFISH2
Private RANK_TOTAL
Private FFW, FFN
Private Sub cmdEXIT_Click()

Call write_parms

End

End Sub
Private Sub cmdGROUP_Click()

AW_OPTION = 0

COMPGEN = "YES"

Label2.Visible = True

Open APPROOT + "\ARTS\CONTROL\COMPUTE.TXT" For Output As #1
Print #1, "YESCOMPUTE"
Close #1

Frame1.Visible = True

optMN1.Value = False
optMN2.Value = False
optMN3.Value = False
optMN4.Value = False

optMJ1.Value = False
optMJ2.Value = False
optMJ3.Value = False
optMJ4.Value = False

optGT1.Value = False
optGT2.Value = False
optGT3.Value = False
optGT4.Value = False

End Sub

Private Sub cmdGUIDE_Click()

HTYPE = "30"

HFNM = APPROOT + "\ARTS\HELP\" + current_language + "HELP" + HTYPE + ".rtf"

If Dir(HFNM) = "" Then Exit Sub

frmCOMP.Enabled = False
Load frmGUIDE
frmGUIDE.Show

End Sub

Private Sub cmdRANK_Click()

COMPGEN = "YES"
COMPFISH = "NOTOK"

Label2.Visible = False

Frame2.Visible = True

cmdREP.Visible = False
cmdGROUP.Visible = False
cmdRANK.Visible = False

optQUIT.Value = False

Dim I

For I = 1 To 13

optC(I - 1).Value = False
optE(I - 1).Value = False
optU(I - 1).Value = False
optP(I - 1).Value = False
optV(I - 1).Value = False

Next I

For I = 1 To 13

If I <= 12 Then
   If VALID_MONTHS(I) = "-" Then GoTo NEXT_I
   End If
   
optC(I - 1).Enabled = True
optE(I - 1).Enabled = True
optU(I - 1).Enabled = True
optP(I - 1).Enabled = True
optV(I - 1).Enabled = True

NEXT_I:

Next I

optQUIT.Enabled = True

End Sub
Private Sub cmdREP_Click()

Load frmREPORTS
Unload frmCOMP
frmREPORTS.Show

End Sub
Private Sub cmdRETURN_Click()

Load frmSEL
Unload frmCOMP
frmSEL.Show

End Sub
Private Sub Form_Load()

Label2.Caption = msgtab(121)

Label2.Visible = False

Set Picture = LoadPicture(APPROOT + "\ARTS\PICS_RUNTIME\SCREEN_03.JPG")

rtsSEL.FileName = APPROOT + "\ARTS\WORK\WSEL.TXT"
rtsSEL.Refresh

Dim J

For J = 1 To 12

If VALID_MONTHS(J) = "X" Then GoTo NEXT_J

optC(J - 1).Enabled = False
optE(J - 1).Enabled = False
optU(J - 1).Enabled = False
optP(J - 1).Enabled = False
optV(J - 1).Enabled = False

NEXT_J:

Next J

RANKYN = "N"

lblMO(12).Caption = CURY

pgbCOMP.Visible = False

Frame1.Visible = False
Frame2.Visible = False

lblMN.Caption = msgtab(49)
lblMJ.Caption = msgtab(48)
lblGT.Caption = msgtab(47)

optMN1.Caption = msgtab(50)
optMN2.Caption = msgtab(51)
optMN3.Caption = msgtab(52)
optMN4.Caption = msgtab(53)

optMJ1.Caption = msgtab(50)
optMJ2.Caption = msgtab(51)
optMJ3.Caption = msgtab(52)
optMJ4.Caption = msgtab(53)

optGT1.Caption = msgtab(50)
optGT2.Caption = msgtab(51)
optGT3.Caption = msgtab(52)
optGT4.Caption = msgtab(53)

cmdREP.ToolTipText = msgtab(54)
cmdGROUP.ToolTipText = msgtab(55)
cmdRETURN.ToolTipText = msgtab(56)
cmdEXIT.ToolTipText = msgtab(57)

cmdGUIDE.ToolTipText = msgtab(6)

cmdRANK.ToolTipText = msgtab(84)

lblC.Caption = msgtab(63)
lblE.Caption = msgtab(64)
lblU.Caption = msgtab(65)
lblP.Caption = msgtab(66)
lblV.Caption = msgtab(67)

optQUIT.Caption = msgtab(88)

lblRANK.Caption = Chr(13) + msgtab(85) + Chr(13) + Chr(13) + msgtab(87)

frmCOMP.MousePointer = 1

frmCOMP.Caption = msgtab(15) + ": " + Format(CURY, "0000") + " - " + _
                 msgtab(45)

End Sub
Private Sub optC_Click(Index As Integer)

Dim I

I = Index

If I = 12 Then
   WELM = "C"
   WPER = CURY
   WMY = "Y"
   Call RANK_PROC_CLEAN
   Exit Sub
   End If

WELM = "C"
WPER = I + 1
WMY = "M"
Call RANK_PROC_CLEAN
Exit Sub
   
End Sub
Private Sub optE_Click(Index As Integer)

Dim I

I = Index

If I = 12 Then
   WELM = "E"
   WPER = CURY
   WMY = "Y"
   Call RANK_PROC_CLEAN
   Exit Sub
   End If

WELM = "E"
WPER = I + 1
WMY = "M"
Call RANK_PROC_CLEAN
Exit Sub

End Sub
Private Sub optGT1_Click()

COMPFISH = "NOTOK"

cmdGROUP.Enabled = False

WLEV = lblGT.Caption
WCOMP = optGT1.Caption

Call WRITE_LOG

optMN1.Enabled = False
optMN2.Enabled = False
optMN3.Enabled = False
optMN4.Enabled = False

optMJ1.Enabled = False
optMJ2.Enabled = False
optMJ3.Enabled = False
optMJ4.Enabled = False

optGT1.Enabled = False
optGT2.Enabled = False
optGT3.Enabled = False
optGT4.Enabled = False

Call GROUPT_BYGT
Call GROUPS_BYGT

Call BYMN_BYBG_BYSP
Call TOTAL_CATCH
Call TOTAL_EFFORT

resp = MsgBox(msgtab(40), vbOKOnly, " ")

pgbCOMP.Visible = False

Frame1.Visible = False

optMN1.Enabled = True
optMN2.Enabled = True
optMN3.Enabled = True
optMN4.Enabled = True

optMJ1.Enabled = True
optMJ2.Enabled = True
optMJ3.Enabled = True
optMJ4.Enabled = True

optGT1.Enabled = True
optGT2.Enabled = True
optGT3.Enabled = True
optGT4.Enabled = True

End Sub
Private Sub optGT2_Click()

COMPFISH = "OK"

cmdGROUP.Enabled = False

WLEV = lblGT.Caption
WCOMP = optGT2.Caption

Call WRITE_LOG

optMN1.Enabled = False
optMN2.Enabled = False
optMN3.Enabled = False
optMN4.Enabled = False

optMJ1.Enabled = False
optMJ2.Enabled = False
optMJ3.Enabled = False
optMJ4.Enabled = False

optGT1.Enabled = False
optGT2.Enabled = False
optGT3.Enabled = False
optGT4.Enabled = False

Call GROUPT_BYGT
Call GROUPS_BYGT

Call BYMN_BYSP_BYBG
Call TOTAL_CATCH
Call TOTAL_EFFORT

resp = MsgBox(msgtab(40), vbOKOnly, " ")

pgbCOMP.Visible = False

Frame1.Visible = False

optMN1.Enabled = True
optMN2.Enabled = True
optMN3.Enabled = True
optMN4.Enabled = True

optMJ1.Enabled = True
optMJ2.Enabled = True
optMJ3.Enabled = True
optMJ4.Enabled = True

optGT1.Enabled = True
optGT2.Enabled = True
optGT3.Enabled = True
optGT4.Enabled = True

End Sub
Private Sub optGT3_Click()

COMPFISH = "NOTOK"

cmdGROUP.Enabled = False

WLEV = lblGT.Caption
WCOMP = optGT3.Caption

Call WRITE_LOG

optMN1.Enabled = False
optMN2.Enabled = False
optMN3.Enabled = False
optMN4.Enabled = False

optMJ1.Enabled = False
optMJ2.Enabled = False
optMJ3.Enabled = False
optMJ4.Enabled = False

optGT1.Enabled = False
optGT2.Enabled = False
optGT3.Enabled = False
optGT4.Enabled = False

Call GROUPT_BYGT
Call GROUPS_BYGT

Dim dbn1, dbn2

dbn1 = APPROOT + "\ARTS\WORK\WT" + Format(CURY, "0000") + ".MDB"
dbn2 = APPROOT + "\ARTS\WORK\WGS" + Format(CURY, "0000") + ".MDB"

If Dir(dbn1) <> "" Then FileCopy dbn1, dbn2

Call TOTAL_CATCH
Call TOTAL_EFFORT

resp = MsgBox(msgtab(40), vbOKOnly, " ")

pgbCOMP.Visible = False

Frame1.Visible = False

optMN1.Enabled = True
optMN2.Enabled = True
optMN3.Enabled = True
optMN4.Enabled = True

optMJ1.Enabled = True
optMJ2.Enabled = True
optMJ3.Enabled = True
optMJ4.Enabled = True

optGT1.Enabled = True
optGT2.Enabled = True
optGT3.Enabled = True
optGT4.Enabled = True

End Sub
Private Sub optGT4_Click()

COMPFISH = "NOTOK"

cmdGROUP.Enabled = False

WLEV = lblGT.Caption
WCOMP = optGT4.Caption

Call WRITE_LOG

optMN1.Enabled = False
optMN2.Enabled = False
optMN3.Enabled = False
optMN4.Enabled = False

optMJ1.Enabled = False
optMJ2.Enabled = False
optMJ3.Enabled = False
optMJ4.Enabled = False

optGT1.Enabled = False
optGT2.Enabled = False
optGT3.Enabled = False
optGT4.Enabled = False

Call GROUPT_BYGT
Call GROUPS_BYGT

Call BYMN_BYSP
Call TOTAL_CATCH
Call TOTAL_EFFORT

resp = MsgBox(msgtab(40), vbOKOnly, " ")

pgbCOMP.Visible = False

Frame1.Visible = False

optMN1.Enabled = True
optMN2.Enabled = True
optMN3.Enabled = True
optMN4.Enabled = True

optMJ1.Enabled = True
optMJ2.Enabled = True
optMJ3.Enabled = True
optMJ4.Enabled = True

optGT1.Enabled = True
optGT2.Enabled = True
optGT3.Enabled = True
optGT4.Enabled = True

End Sub
Private Sub optMJ1_Click()

COMPFISH = "NOTOK"

cmdGROUP.Enabled = False

WLEV = lblMJ.Caption
WCOMP = optMJ1.Caption

Call WRITE_LOG

optMN1.Enabled = False
optMN2.Enabled = False
optMN3.Enabled = False
optMN4.Enabled = False

optMJ1.Enabled = False
optMJ2.Enabled = False
optMJ3.Enabled = False
optMJ4.Enabled = False

optGT1.Enabled = False
optGT2.Enabled = False
optGT3.Enabled = False
optGT4.Enabled = False

Call GROUPT_BYMJ
Call GROUPS_BYMJ

Call BYMN_BYBG_BYSP
Call TOTAL_CATCH
Call TOTAL_EFFORT

resp = MsgBox(msgtab(40), vbOKOnly, " ")

pgbCOMP.Visible = False

Frame1.Visible = False

optMN1.Enabled = True
optMN2.Enabled = True
optMN3.Enabled = True
optMN4.Enabled = True

optMJ1.Enabled = True
optMJ2.Enabled = True
optMJ3.Enabled = True
optMJ4.Enabled = True

optGT1.Enabled = True
optGT2.Enabled = True
optGT3.Enabled = True
optGT4.Enabled = True

End Sub
Private Sub optMJ2_Click()

COMPFISH = "OK"

cmdGROUP.Enabled = False

WLEV = lblMJ.Caption
WCOMP = optMJ2.Caption

Call WRITE_LOG

optMN1.Enabled = False
optMN2.Enabled = False
optMN3.Enabled = False
optMN4.Enabled = False

optMJ1.Enabled = False
optMJ2.Enabled = False
optMJ3.Enabled = False
optMJ4.Enabled = False

optGT1.Enabled = False
optGT2.Enabled = False
optGT3.Enabled = False
optGT4.Enabled = False

Call GROUPT_BYMJ
Call GROUPS_BYMJ

Call BYMN_BYSP_BYBG
Call TOTAL_CATCH
Call TOTAL_EFFORT

resp = MsgBox(msgtab(40), vbOKOnly, " ")

pgbCOMP.Visible = False

Frame1.Visible = False

optMN1.Enabled = True
optMN2.Enabled = True
optMN3.Enabled = True
optMN4.Enabled = True

optMJ1.Enabled = True
optMJ2.Enabled = True
optMJ3.Enabled = True
optMJ4.Enabled = True

optGT1.Enabled = True
optGT2.Enabled = True
optGT3.Enabled = True
optGT4.Enabled = True

End Sub
Private Sub optMJ3_Click()

COMPFISH = "NOTOK"

cmdGROUP.Enabled = False

WLEV = lblMJ.Caption
WCOMP = optMJ3.Caption

Call WRITE_LOG

optMN1.Enabled = False
optMN2.Enabled = False
optMN3.Enabled = False
optMN4.Enabled = False

optMJ1.Enabled = False
optMJ2.Enabled = False
optMJ3.Enabled = False
optMJ4.Enabled = False

optGT1.Enabled = False
optGT2.Enabled = False
optGT3.Enabled = False
optGT4.Enabled = False

Call GROUPT_BYMJ
Call GROUPS_BYMJ

Dim dbn1, dbn2

dbn1 = APPROOT + "\ARTS\WORK\WT" + Format(CURY, "0000") + ".MDB"
dbn2 = APPROOT + "\ARTS\WORK\WGS" + Format(CURY, "0000") + ".MDB"

If Dir(dbn1) <> "" Then FileCopy dbn1, dbn2

Call TOTAL_CATCH
Call TOTAL_EFFORT

resp = MsgBox(msgtab(40), vbOKOnly, " ")

pgbCOMP.Visible = False

Frame1.Visible = False

optMN1.Enabled = True
optMN2.Enabled = True
optMN3.Enabled = True
optMN4.Enabled = True

optMJ1.Enabled = True
optMJ2.Enabled = True
optMJ3.Enabled = True
optMJ4.Enabled = True

optGT1.Enabled = True
optGT2.Enabled = True
optGT3.Enabled = True
optGT4.Enabled = True

End Sub
Private Sub optMJ4_Click()

COMPFISH = "NOTOK"

cmdGROUP.Enabled = False

WLEV = lblMJ.Caption
WCOMP = optMJ4.Caption

Call WRITE_LOG

optMN1.Enabled = False
optMN2.Enabled = False
optMN3.Enabled = False
optMN4.Enabled = False

optMJ1.Enabled = False
optMJ2.Enabled = False
optMJ3.Enabled = False
optMJ4.Enabled = False

optGT1.Enabled = False
optGT2.Enabled = False
optGT3.Enabled = False
optGT4.Enabled = False

Call GROUPT_BYMJ
Call GROUPS_BYMJ

Call BYMN_BYSP
Call TOTAL_CATCH
Call TOTAL_EFFORT

resp = MsgBox(msgtab(40), vbOKOnly, " ")

pgbCOMP.Visible = False

Frame1.Visible = False

optMN1.Enabled = True
optMN2.Enabled = True
optMN3.Enabled = True
optMN4.Enabled = True

optMJ1.Enabled = True
optMJ2.Enabled = True
optMJ3.Enabled = True
optMJ4.Enabled = True

optGT1.Enabled = True
optGT2.Enabled = True
optGT3.Enabled = True
optGT4.Enabled = True


End Sub
Private Sub optMN1_Click()

COMPFISH = "NOTOK"

cmdGROUP.Enabled = False

WLEV = lblMN.Caption
WCOMP = optMN1.Caption

Call WRITE_LOG

optMN1.Enabled = False
optMN2.Enabled = False
optMN3.Enabled = False
optMN4.Enabled = False

optMJ1.Enabled = False
optMJ2.Enabled = False
optMJ3.Enabled = False
optMJ4.Enabled = False

optGT1.Enabled = False
optGT2.Enabled = False
optGT3.Enabled = False
optGT4.Enabled = False

Dim dbn1, dbn2

dbn1 = APPROOT + "\ARTS\WORK\WT" + Format(CURY, "0000") + ".MDB"
dbn2 = APPROOT + "\ARTS\WORK\WGT" + Format(CURY, "0000") + ".MDB"

If Dir(dbn1) <> "" Then FileCopy dbn1, dbn2

dbn1 = APPROOT + "\ARTS\WORK\WS" + Format(CURY, "0000") + ".MDB"
dbn2 = APPROOT + "\ARTS\WORK\WGS" + Format(CURY, "0000") + ".MDB"

If Dir(dbn1) <> "" Then FileCopy dbn1, dbn2

Call BYMN_BYBG_BYSP
Call TOTAL_CATCH
Call TOTAL_EFFORT

resp = MsgBox(msgtab(40), vbOKOnly, " ")

pgbCOMP.Visible = False

Frame1.Visible = False

optMN1.Enabled = True
optMN2.Enabled = True
optMN3.Enabled = True
optMN4.Enabled = True

optMJ1.Enabled = True
optMJ2.Enabled = True
optMJ3.Enabled = True
optMJ4.Enabled = True

optGT1.Enabled = True
optGT2.Enabled = True
optGT3.Enabled = True
optGT4.Enabled = True

End Sub
Private Sub BYMN_BYSP_BYBG()

Dim I, ix, FFC, FFE, FFP, FFU, FFV, RR, cc, PP

Dim dbn1, dbn2, XKEY1, XKEY2, NREC, IREC

dbn1 = APPROOT + "\ARTS\WORK\WT" + Format(CURY, "0000") + ".MDB"
dbn2 = APPROOT + "\ARTS\WORK\WGT" + Format(CURY, "0000") + ".MDB"

If Dir(dbn1) <> "" Then FileCopy dbn1, dbn2

dbn1 = APPROOT + "\ARTS\WORK\WS" + Format(CURY, "0000") + ".MDB"
dbn2 = APPROOT + "\ARTS\WORK\WGS" + Format(CURY, "0000") + ".MDB"

FileCopy APPROOT + "\ARTS\STRUS\ARTS.MDB", dbn2

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(dbn1)
Set prm_record = prm_database.OpenRecordset("ASITAB")

Dim prm2_database As Database, prm2_record As Recordset

Set prm2_database = OpenDatabase(dbn2)
Set prm2_record = prm2_database.OpenRecordset("ASITAB")

prm2_record.Index = "primarykey"

With prm_record

NREC = .RecordCount: IREC = 0

pgbCOMP.Min = 0: pgbCOMP.Max = NREC
pgbCOMP.Visible = True

.Index = "primarykey"

.MoveFirst

Do Until .EOF

IREC = IREC + 1: pgbCOMP.Value = IREC

XKEY1 = ![akey]
XKEY2 = Left(XKEY1, 11) + Right(XKEY1, 6) + Mid(XKEY1, 12, 6)

prm2_record.AddNew

prm2_record![akey] = XKEY2

For I = 1 To 13

ix = Format(I, "00")

FFC = "C" + ix:  prm2_record.Fields(FFC) = .Fields(FFC)
FFE = "E" + ix:  prm2_record.Fields(FFE) = .Fields(FFE)
FFU = "U" + ix:  prm2_record.Fields(FFU) = .Fields(FFU)
FFP = "P" + ix:  prm2_record.Fields(FFP) = .Fields(FFP)
FFV = "V" + ix:  prm2_record.Fields(FFV) = .Fields(FFV)
FFW = "W" + ix:  prm2_record.Fields(FFW) = .Fields(FFW)
FFN = "F" + ix:  prm2_record.Fields(FFN) = .Fields(FFN)

Next I

prm2_record![RANK] = ![RANK]
prm2_record![PER] = ![PER]
prm2_record![CUM] = ![CUM]
prm2_record![ADDFISH] = ![ADDFISH]

prm2_record.Update

'SPECIES TOTALS
'==============

XKEY2 = Left(XKEY2, 18) + "B0000"

prm2_record.Seek "=", XKEY2

If prm2_record.NoMatch = False Then GoTo UPDATE_OLD

prm2_record.AddNew

prm2_record![akey] = XKEY2

For I = 1 To 13

ix = Format(I, "00")

FFC = "C" + ix:  prm2_record.Fields(FFC) = .Fields(FFC)
FFE = "E" + ix:  prm2_record.Fields(FFE) = .Fields(FFE)
FFU = "U" + ix:  prm2_record.Fields(FFU) = .Fields(FFU)
FFP = "P" + ix:  prm2_record.Fields(FFP) = .Fields(FFP)
FFV = "V" + ix:  prm2_record.Fields(FFV) = .Fields(FFV)
FFW = "W" + ix:  prm2_record.Fields(FFW) = .Fields(FFW)
FFN = "F" + ix:  prm2_record.Fields(FFN) = .Fields(FFN)

Next I

prm2_record![RANK] = -999
prm2_record![PER] = -999
prm2_record![CUM] = -999
prm2_record![ADDFISH] = ![ADDFISH]

prm2_record.Update

GoTo NEXT_REC

UPDATE_OLD:

prm2_record.Edit

For I = 1 To 13

ix = Format(I, "00")

FFC = "C" + ix:  prm2_record.Fields(FFC) = prm2_record.Fields(FFC) + .Fields(FFC)
FFE = "E" + ix:  prm2_record.Fields(FFE) = prm2_record.Fields(FFE) + .Fields(FFE)
FFU = "U" + ix:  prm2_record.Fields(FFU) = prm2_record.Fields(FFU) + .Fields(FFU)
FFP = "P" + ix:  prm2_record.Fields(FFP) = prm2_record.Fields(FFP) + .Fields(FFP)
FFV = "V" + ix:  prm2_record.Fields(FFV) = prm2_record.Fields(FFV) + .Fields(FFV)
FFW = "W" + ix:  prm2_record.Fields(FFW) = prm2_record.Fields(FFW) + .Fields(FFW)
FFN = "F" + ix:  prm2_record.Fields(FFN) = prm2_record.Fields(FFN) + .Fields(FFN)

prm2_record.Fields(FFU) = 0
prm2_record.Fields(FFP) = 0
prm2_record.Fields(FFW) = 0

If prm2_record.Fields(FFE) <> 0 Then
   prm2_record.Fields(FFU) = prm2_record.Fields(FFC) / prm2_record.Fields(FFE)
   End If

If prm2_record.Fields(FFC) <> 0 Then
   prm2_record.Fields(FFP) = prm2_record.Fields(FFV) / prm2_record.Fields(FFC)
   End If

If prm2_record.Fields(FFN) <> 0 Then
   prm2_record.Fields(FFW) = prm2_record.Fields(FFC) / prm2_record.Fields(FFN)
   End If

If ![ADDFISH] = "NO" Then prm2_record![ADDFISH] = "NO"

If prm2_record![ADDFISH] = "YES" And ![ADDFISH] = "NO" Then prm2_record![ADDFISH] = "NO"

Next I

prm2_record![RANK] = -999
prm2_record![PER] = -999
prm2_record![CUM] = -999

prm2_record.Update

'END SPECIES TOTALS
'==================

NEXT_REC:

.MoveNext

Loop

End With

prm_record.Close
prm_database.Close

prm2_record.Close
prm2_database.Close

End Sub
Private Sub optMN2_Click()

COMPFISH = "OK"

cmdGROUP.Enabled = False

WLEV = lblMN.Caption
WCOMP = optMN2.Caption

Call WRITE_LOG

optMN1.Enabled = False
optMN2.Enabled = False
optMN3.Enabled = False
optMN4.Enabled = False

optMJ1.Enabled = False
optMJ2.Enabled = False
optMJ3.Enabled = False
optMJ4.Enabled = False

optGT1.Enabled = False
optGT2.Enabled = False
optGT3.Enabled = False
optGT4.Enabled = False

Call BYMN_BYSP_BYBG
Call TOTAL_CATCH
Call TOTAL_EFFORT

resp = MsgBox(msgtab(40), vbOKOnly, " ")

pgbCOMP.Visible = False

Frame1.Visible = False

optMN1.Enabled = True
optMN2.Enabled = True
optMN3.Enabled = True
optMN4.Enabled = True

optMJ1.Enabled = True
optMJ2.Enabled = True
optMJ3.Enabled = True
optMJ4.Enabled = True

optGT1.Enabled = True
optGT2.Enabled = True
optGT3.Enabled = True
optGT4.Enabled = True

End Sub
Private Sub TOTAL_EFFORT()

Dim TE()

ReDim TE(1 To 13)

Dim I, ix, FFC, FFE, FFP, FFU, FFV, RR, cc, PP, FFF

For I = 1 To 13
TE(I) = 0
Next I

Dim dbn1, dbn2, XKEY1, XKEY2, NREC, IREC

dbn1 = APPROOT + "\ARTS\WORK\WT" + Format(CURY, "0000") + ".MDB"
dbn2 = APPROOT + "\ARTS\WORK\WGS" + Format(CURY, "0000") + ".MDB"

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(dbn1)
Set prm_record = prm_database.OpenRecordset("ASITAB")

Dim prm2_database As Database, prm2_record As Recordset

Set prm2_database = OpenDatabase(dbn2)
Set prm2_record = prm2_database.OpenRecordset("ASITAB")

With prm_record

NREC = .RecordCount: IREC = 0

pgbCOMP.Min = 0: pgbCOMP.Max = NREC
pgbCOMP.Visible = True

.Index = "primarykey"

.MoveFirst

Do Until .EOF

IREC = IREC + 1: pgbCOMP.Value = IREC

For I = 1 To 13

ix = Format(I, "00")

FFE = "E" + ix:  TE(I) = TE(I) + .Fields(FFE)

Next I

.MoveNext

Loop

End With

prm_record.Close
prm_database.Close

prm2_record.Index = "primarykey"

XKEY2 = " GT"

prm2_record.Seek "=", XKEY2

prm2_record.Edit

For I = 1 To 13

FFC = "C" + Format(I, "00")
FFE = "E" + Format(I, "00")
FFU = "U" + Format(I, "00")
FFP = "P" + Format(I, "00")
FFV = "V" + Format(I, "00")
FFW = "W" + Format(I, "00")
FFN = "F" + Format(I, "00")

prm2_record.Fields(FFE) = TE(I)

prm2_record![RANK] = -999
prm2_record![CUM] = -999
prm2_record![PER] = -999

prm2_record.Fields(FFU) = 0
prm2_record.Fields(FFP) = 0
prm2_record.Fields(FFW) = 0

If prm2_record.Fields(FFE) <> 0 Then
   prm2_record.Fields(FFU) = prm2_record.Fields(FFC) / prm2_record.Fields(FFE)
   End If

If prm2_record.Fields(FFC) <> 0 Then
   prm2_record.Fields(FFP) = prm2_record.Fields(FFV) / prm2_record.Fields(FFC)
   End If

If prm2_record.Fields(FFN) <> 0 Then
   prm2_record.Fields(FFW) = prm2_record.Fields(FFC) / prm2_record.Fields(FFN)
   End If

Next I

If RANK_TOTAL <> "Y" Then GoTo NO_RANK

If WMY = "M" Then FFF = WELM + Format(WPER, "00")
If WMY = "Y" Then FFF = WELM + "13"

prm2_record![RANK] = -prm2_record.Fields(FFF)
prm2_record![PER] = 100
prm2_record![CUM] = 100

NO_RANK:

prm2_record.Update

prm2_record.Close
prm2_database.Close

End Sub
Private Sub BYMN_BYBG_BYSP()

Dim I, ix, FFC, FFE, FFP, FFU, FFV, RR, cc, PP

Dim dbn1, dbn2, XKEY1, XKEY2, NREC, IREC

dbn1 = APPROOT + "\ARTS\WORK\WGT" + Format(CURY, "0000") + ".MDB"
dbn2 = APPROOT + "\ARTS\WORK\WGS" + Format(CURY, "0000") + ".MDB"

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(dbn1)
Set prm_record = prm_database.OpenRecordset("ASITAB")

Dim prm2_database As Database, prm2_record As Recordset

Set prm2_database = OpenDatabase(dbn2)
Set prm2_record = prm2_database.OpenRecordset("ASITAB")

prm2_record.Index = "primarykey"

With prm_record

NREC = .RecordCount: IREC = 0

pgbCOMP.Min = 0: pgbCOMP.Max = NREC
pgbCOMP.Visible = True

.Index = "primarykey"

.MoveFirst

Do Until .EOF

IREC = IREC + 1: pgbCOMP.Value = IREC

XKEY1 = ![akey]
XKEY2 = XKEY1

prm2_record.AddNew

prm2_record![akey] = XKEY2

For I = 1 To 13

ix = Format(I, "00")

FFC = "C" + ix:  prm2_record.Fields(FFC) = .Fields(FFC)
FFE = "E" + ix:  prm2_record.Fields(FFE) = .Fields(FFE)
FFU = "U" + ix:  prm2_record.Fields(FFU) = .Fields(FFU)
FFP = "P" + ix:  prm2_record.Fields(FFP) = .Fields(FFP)
FFV = "V" + ix:  prm2_record.Fields(FFV) = .Fields(FFV)
FFW = "W" + ix:  prm2_record.Fields(FFW) = .Fields(FFW)
FFN = "F" + ix:  prm2_record.Fields(FFN) = .Fields(FFN)

Next I

prm2_record![RANK] = -999
prm2_record![PER] = -999
prm2_record![CUM] = -999
prm2_record![ADDFISH] = "NO"

prm2_record.Update

NEXT_REC:

.MoveNext

Loop

End With

prm_record.Close
prm_database.Close

prm2_record.Close
prm2_database.Close

End Sub
Private Sub TOTAL_CATCH()

Dim TC(), TV(), TF()

ReDim TC(1 To 13), TV(1 To 13), TF(1 To 13)

Dim I, ix, FFC, FFE, FFP, FFU, FFV, RR, cc, PP

For I = 1 To 13
TC(I) = 0: TV(I) = 0: TF(I) = 0
Next I

Dim dbn1, dbn2, XKEY1, XKEY2, NREC, IREC

dbn1 = APPROOT + "\ARTS\WORK\WS" + Format(CURY, "0000") + ".MDB"
dbn2 = APPROOT + "\ARTS\WORK\WGS" + Format(CURY, "0000") + ".MDB"

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(dbn1)
Set prm_record = prm_database.OpenRecordset("ASITAB")

Dim prm2_database As Database, prm2_record As Recordset

Set prm2_database = OpenDatabase(dbn2)
Set prm2_record = prm2_database.OpenRecordset("ASITAB")

With prm_record

NREC = .RecordCount: IREC = 0

pgbCOMP.Min = 0: pgbCOMP.Max = NREC
pgbCOMP.Visible = True

.Index = "primarykey"

.MoveFirst

ADDFISH = "YES"

Do Until .EOF

IREC = IREC + 1: pgbCOMP.Value = IREC

If ![ADDFISH] = "NO" Then ADDFISH = "NO"

For I = 1 To 13

ix = Format(I, "00")

FFC = "C" + ix:  FFV = "V" + ix: FFN = "F" + ix

TC(I) = TC(I) + .Fields(FFC)
TV(I) = TV(I) + .Fields(FFV)
TF(I) = TF(I) + .Fields(FFN)

Next I

.MoveNext

Loop

End With

prm_record.Close
prm_database.Close

prm2_record.Index = "primarykey"

XKEY2 = " GT"

prm2_record.AddNew

prm2_record![akey] = " GT"

For I = 1 To 13

FFC = "C" + Format(I, "00")
FFE = "E" + Format(I, "00")
FFU = "U" + Format(I, "00")
FFP = "P" + Format(I, "00")
FFV = "V" + Format(I, "00")
FFW = "W" + Format(I, "00")
FFN = "F" + Format(I, "00")

prm2_record.Fields(FFC) = TC(I)
prm2_record.Fields(FFV) = TV(I)
prm2_record.Fields(FFN) = TF(I)

prm2_record![RANK] = -999
prm2_record![CUM] = -999
prm2_record![PER] = -999
prm2_record![ADDFISH] = ADDFISH

prm2_record.Fields(FFU) = 0
prm2_record.Fields(FFP) = 0
prm2_record.Fields(FFW) = 0

Next I

prm2_record.Update

prm2_record.Close
prm2_database.Close

End Sub
Private Sub optMN3_Click()

COMPFISH = "NOTOK"

cmdGROUP.Enabled = False

WLEV = lblMN.Caption
WCOMP = optMN3.Caption

Call WRITE_LOG

optMN1.Enabled = False
optMN2.Enabled = False
optMN3.Enabled = False
optMN4.Enabled = False

optMJ1.Enabled = False
optMJ2.Enabled = False
optMJ3.Enabled = False
optMJ4.Enabled = False

optGT1.Enabled = False
optGT2.Enabled = False
optGT3.Enabled = False
optGT4.Enabled = False

Dim dbn1, dbn2

dbn1 = APPROOT + "\ARTS\WORK\WT" + Format(CURY, "0000") + ".MDB"
dbn2 = APPROOT + "\ARTS\WORK\WGS" + Format(CURY, "0000") + ".MDB"

If Dir(dbn1) <> "" Then FileCopy dbn1, dbn2

Call TOTAL_CATCH
Call TOTAL_EFFORT

resp = MsgBox(msgtab(40), vbOKOnly, " ")

pgbCOMP.Visible = False

Frame1.Visible = False

optMN1.Enabled = True
optMN2.Enabled = True
optMN3.Enabled = True
optMN4.Enabled = True

optMJ1.Enabled = True
optMJ2.Enabled = True
optMJ3.Enabled = True
optMJ4.Enabled = True

optGT1.Enabled = True
optGT2.Enabled = True
optGT3.Enabled = True
optGT4.Enabled = True

End Sub
Private Sub optMN4_Click()
COMPFISH = "NOTOK"

cmdGROUP.Enabled = False

WLEV = lblMN.Caption
WCOMP = optMN4.Caption

Call WRITE_LOG

optMN1.Enabled = False
optMN2.Enabled = False
optMN3.Enabled = False
optMN4.Enabled = False

optMJ1.Enabled = False
optMJ2.Enabled = False
optMJ3.Enabled = False
optMJ4.Enabled = False

optGT1.Enabled = False
optGT2.Enabled = False
optGT3.Enabled = False
optGT4.Enabled = False

Call BYMN_BYSP
Call TOTAL_CATCH
Call TOTAL_EFFORT

resp = MsgBox(msgtab(40), vbOKOnly, " ")

pgbCOMP.Visible = False

Frame1.Visible = False

optMN1.Enabled = True
optMN2.Enabled = True
optMN3.Enabled = True
optMN4.Enabled = True

optMJ1.Enabled = True
optMJ2.Enabled = True
optMJ3.Enabled = True
optMJ4.Enabled = True

optGT1.Enabled = True
optGT2.Enabled = True
optGT3.Enabled = True
optGT4.Enabled = True

End Sub
Private Sub BYMN_BYSP()

Dim I, ix, FFC, FFE, FFP, FFU, FFV, RR, cc, PP

Dim dbn1, dbn2, XKEY1, XKEY2, NREC, IREC

dbn1 = APPROOT + "\ARTS\WORK\WS" + Format(CURY, "0000") + ".MDB"
dbn2 = APPROOT + "\ARTS\WORK\WGS" + Format(CURY, "0000") + ".MDB"

FileCopy APPROOT + "\ARTS\STRUS\ARTS.MDB", dbn2

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(dbn1)
Set prm_record = prm_database.OpenRecordset("ASITAB")

Dim prm2_database As Database, prm2_record As Recordset

Set prm2_database = OpenDatabase(dbn2)
Set prm2_record = prm2_database.OpenRecordset("ASITAB")

prm2_record.Index = "primarykey"

With prm_record

NREC = .RecordCount: IREC = 0

pgbCOMP.Min = 0: pgbCOMP.Max = NREC
pgbCOMP.Visible = True

.Index = "primarykey"

.MoveFirst

Do Until .EOF

IREC = IREC + 1: pgbCOMP.Value = IREC

XKEY1 = ![akey]
XKEY2 = Left(XKEY1, 11) + Right(XKEY1, 6) + "+B0000"

'SPECIES TOTALS
'==============

XKEY2 = Left(XKEY2, 18) + "B0000"

prm2_record.Seek "=", XKEY2

If prm2_record.NoMatch = False Then GoTo UPDATE_OLD

prm2_record.AddNew

prm2_record![akey] = XKEY2

For I = 1 To 13

ix = Format(I, "00")

FFC = "C" + ix:  prm2_record.Fields(FFC) = .Fields(FFC)
FFE = "E" + ix:  prm2_record.Fields(FFE) = .Fields(FFE)
FFU = "U" + ix:  prm2_record.Fields(FFU) = .Fields(FFU)
FFP = "P" + ix:  prm2_record.Fields(FFP) = .Fields(FFP)
FFV = "V" + ix:  prm2_record.Fields(FFV) = .Fields(FFV)
FFW = "W" + ix:  prm2_record.Fields(FFW) = .Fields(FFW)
FFN = "F" + ix:  prm2_record.Fields(FFN) = .Fields(FFN)

Next I

prm2_record![RANK] = 0
prm2_record![PER] = 0
prm2_record![CUM] = 0

prm2_record.Update

GoTo NEXT_REC

UPDATE_OLD:

prm2_record.Edit

For I = 1 To 13

ix = Format(I, "00")

FFC = "C" + ix:  prm2_record.Fields(FFC) = prm2_record.Fields(FFC) + .Fields(FFC)
FFE = "E" + ix:  prm2_record.Fields(FFE) = prm2_record.Fields(FFE) + .Fields(FFE)
FFU = "U" + ix:  prm2_record.Fields(FFU) = prm2_record.Fields(FFU) + .Fields(FFU)
FFP = "P" + ix:  prm2_record.Fields(FFP) = prm2_record.Fields(FFP) + .Fields(FFP)
FFV = "V" + ix:  prm2_record.Fields(FFV) = prm2_record.Fields(FFV) + .Fields(FFV)
FFW = "W" + ix:  prm2_record.Fields(FFW) = prm2_record.Fields(FFW) + .Fields(FFW)
FFN = "F" + ix:  prm2_record.Fields(FFN) = prm2_record.Fields(FFN) + .Fields(FFN)

prm2_record.Fields(FFU) = 0
prm2_record.Fields(FFP) = 0
prm2_record.Fields(FFW) = 0

If prm2_record.Fields(FFE) <> 0 Then
   prm2_record.Fields(FFU) = prm2_record.Fields(FFC) / prm2_record.Fields(FFE)
   End If

If prm2_record.Fields(FFC) <> 0 Then
   prm2_record.Fields(FFP) = prm2_record.Fields(FFV) / prm2_record.Fields(FFC)
   End If

If prm2_record.Fields(FFN) <> 0 Then
   prm2_record.Fields(FFW) = prm2_record.Fields(FFC) / prm2_record.Fields(FFN)
   End If

Next I

prm2_record![RANK] = 0
prm2_record![PER] = 0
prm2_record![CUM] = 0

prm2_record.Update

'END SPECIES TOTALS
'==================

NEXT_REC:

.MoveNext

Loop

End With

prm_record.Close
prm_database.Close

prm2_record.Close
prm2_database.Close

End Sub
Private Sub GROUPT_BYMJ()

Dim I, ix, FFC, FFE, FFP, FFU, FFV, RR, cc, PP

Dim dbn1, dbn2, XKEY1, XKEY2, NREC, IREC

dbn1 = APPROOT + "\ARTS\WORK\WT" + Format(CURY, "0000") + ".MDB"
dbn2 = APPROOT + "\ARTS\WORK\WGT" + Format(CURY, "0000") + ".MDB"

FileCopy APPROOT + "\ARTS\STRUS\ARTS.MDB", dbn2

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(dbn1)
Set prm_record = prm_database.OpenRecordset("ASITAB")

Dim prm2_database As Database, prm2_record As Recordset

Set prm2_database = OpenDatabase(dbn2)
Set prm2_record = prm2_database.OpenRecordset("ASITAB")

prm2_record.Index = "primarykey"

With prm_record

NREC = .RecordCount: IREC = 0

pgbCOMP.Min = 0: pgbCOMP.Max = NREC
pgbCOMP.Visible = True

.Index = "primarykey"

.MoveFirst

Do Until .EOF

IREC = IREC + 1: pgbCOMP.Value = IREC

XKEY1 = ![akey]
XKEY2 = Left(XKEY1, 5) + "+M0000" + Right(XKEY1, 12)

prm2_record.Seek "=", XKEY2

If prm2_record.NoMatch = False Then GoTo UPDATE_OLD

prm2_record.AddNew

prm2_record![akey] = XKEY2

For I = 1 To 13

ix = Format(I, "00")

FFC = "C" + ix:  prm2_record.Fields(FFC) = .Fields(FFC)
FFE = "E" + ix:  prm2_record.Fields(FFE) = .Fields(FFE)
FFU = "U" + ix:  prm2_record.Fields(FFU) = .Fields(FFU)
FFP = "P" + ix:  prm2_record.Fields(FFP) = .Fields(FFP)
FFV = "V" + ix:  prm2_record.Fields(FFV) = .Fields(FFV)
FFW = "W" + ix:  prm2_record.Fields(FFW) = .Fields(FFW)
FFN = "F" + ix:  prm2_record.Fields(FFN) = .Fields(FFN)

Next I

prm2_record![RANK] = 0
prm2_record![PER] = 0
prm2_record![CUM] = 0

prm2_record.Update

GoTo NEXT_REC

UPDATE_OLD:

prm2_record.Edit

For I = 1 To 13

ix = Format(I, "00")

FFC = "C" + ix:  prm2_record.Fields(FFC) = prm2_record.Fields(FFC) + .Fields(FFC)
FFE = "E" + ix:  prm2_record.Fields(FFE) = prm2_record.Fields(FFE) + .Fields(FFE)
FFU = "U" + ix:  prm2_record.Fields(FFU) = prm2_record.Fields(FFU) + .Fields(FFU)
FFP = "P" + ix:  prm2_record.Fields(FFP) = prm2_record.Fields(FFP) + .Fields(FFP)
FFV = "V" + ix:  prm2_record.Fields(FFV) = prm2_record.Fields(FFV) + .Fields(FFV)
FFW = "W" + ix:  prm2_record.Fields(FFW) = prm2_record.Fields(FFW) + .Fields(FFW)
FFN = "F" + ix:  prm2_record.Fields(FFN) = prm2_record.Fields(FFN) + .Fields(FFN)

prm2_record.Fields(FFU) = 0
prm2_record.Fields(FFP) = 0
prm2_record.Fields(FFW) = 0

If prm2_record.Fields(FFE) <> 0 Then
   prm2_record.Fields(FFU) = prm2_record.Fields(FFC) / prm2_record.Fields(FFE)
   End If

If prm2_record.Fields(FFC) <> 0 Then
   prm2_record.Fields(FFP) = prm2_record.Fields(FFV) / prm2_record.Fields(FFC)
   End If

If prm2_record.Fields(FFN) <> 0 Then
   prm2_record.Fields(FFW) = prm2_record.Fields(FFC) / prm2_record.Fields(FFN)
   End If

Next I

prm2_record![RANK] = 0
prm2_record![PER] = 0
prm2_record![CUM] = 0

prm2_record.Update

NEXT_REC:

.MoveNext

Loop

End With

prm_record.Close
prm_database.Close

prm2_record.Close
prm2_database.Close

FileCopy dbn2, dbn1

End Sub
Private Sub GROUPS_BYMJ()

Dim I, ix, FFC, FFE, FFP, FFU, FFV, RR, cc, PP

Dim dbn1, dbn2, XKEY1, XKEY2, NREC, IREC

dbn1 = APPROOT + "\ARTS\WORK\WS" + Format(CURY, "0000") + ".MDB"
dbn2 = APPROOT + "\ARTS\WORK\WGS" + Format(CURY, "0000") + ".MDB"

FileCopy APPROOT + "\ARTS\STRUS\ARTS.MDB", dbn2

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(dbn1)
Set prm_record = prm_database.OpenRecordset("ASITAB")

Dim prm2_database As Database, prm2_record As Recordset

Set prm2_database = OpenDatabase(dbn2)
Set prm2_record = prm2_database.OpenRecordset("ASITAB")

prm2_record.Index = "primarykey"

With prm_record

NREC = .RecordCount: IREC = 0

pgbCOMP.Min = 0: pgbCOMP.Max = NREC
pgbCOMP.Visible = True

.Index = "primarykey"

.MoveFirst

Do Until .EOF

IREC = IREC + 1: pgbCOMP.Value = IREC

XKEY1 = ![akey]
XKEY2 = Left(XKEY1, 5) + "+M0000" + Right(XKEY1, 12)

prm2_record.Seek "=", XKEY2

If prm2_record.NoMatch = False Then GoTo UPDATE_OLD

prm2_record.AddNew

prm2_record![akey] = XKEY2

For I = 1 To 13

ix = Format(I, "00")

FFC = "C" + ix:  prm2_record.Fields(FFC) = .Fields(FFC)
FFE = "E" + ix:  prm2_record.Fields(FFE) = .Fields(FFE)
FFU = "U" + ix:  prm2_record.Fields(FFU) = .Fields(FFU)
FFP = "P" + ix:  prm2_record.Fields(FFP) = .Fields(FFP)
FFV = "V" + ix:  prm2_record.Fields(FFV) = .Fields(FFV)
FFW = "W" + ix:  prm2_record.Fields(FFW) = .Fields(FFW)
FFN = "F" + ix:  prm2_record.Fields(FFN) = .Fields(FFN)

Next I

prm2_record![RANK] = 0
prm2_record![PER] = 0
prm2_record![CUM] = 0

prm2_record![ADDFISH] = ![ADDFISH]

prm2_record.Update

GoTo NEXT_REC

UPDATE_OLD:

prm2_record.Edit

For I = 1 To 13

ix = Format(I, "00")

FFC = "C" + ix:  prm2_record.Fields(FFC) = prm2_record.Fields(FFC) + .Fields(FFC)
FFE = "E" + ix:  prm2_record.Fields(FFE) = prm2_record.Fields(FFE) + .Fields(FFE)
FFU = "U" + ix:  prm2_record.Fields(FFU) = prm2_record.Fields(FFU) + .Fields(FFU)
FFP = "P" + ix:  prm2_record.Fields(FFP) = prm2_record.Fields(FFP) + .Fields(FFP)
FFV = "V" + ix:  prm2_record.Fields(FFV) = prm2_record.Fields(FFV) + .Fields(FFV)
FFW = "W" + ix:  prm2_record.Fields(FFW) = prm2_record.Fields(FFW) + .Fields(FFW)
FFN = "F" + ix:  prm2_record.Fields(FFN) = prm2_record.Fields(FFN) + .Fields(FFN)

prm2_record.Fields(FFU) = 0
prm2_record.Fields(FFP) = 0
prm2_record.Fields(FFW) = 0

If prm2_record.Fields(FFE) <> 0 Then
   prm2_record.Fields(FFU) = prm2_record.Fields(FFC) / prm2_record.Fields(FFE)
   End If

If prm2_record.Fields(FFC) <> 0 Then
   prm2_record.Fields(FFP) = prm2_record.Fields(FFV) / prm2_record.Fields(FFC)
   End If

If prm2_record.Fields(FFN) <> 0 Then
   prm2_record.Fields(FFW) = prm2_record.Fields(FFC) / prm2_record.Fields(FFN)
   End If

Next I

prm2_record![RANK] = 0
prm2_record![PER] = 0
prm2_record![CUM] = 0

If ![ADDFISH] = "NO" Then prm2_record![ADDFISH] = "NO"

If prm2_record![ADDFISH] = "YES" And ![ADDFISH] = "NO" Then prm2_record![ADDFISH] = "NO"

prm2_record.Update

NEXT_REC:

.MoveNext

Loop

End With

prm_record.Close
prm_database.Close

prm2_record.Close
prm2_database.Close

FileCopy dbn2, dbn1

End Sub
Private Sub GROUPT_BYGT()

Dim I, ix, FFC, FFE, FFP, FFU, FFV, RR, cc, PP

Dim dbn1, dbn2, XKEY1, XKEY2, NREC, IREC

dbn1 = APPROOT + "\ARTS\WORK\WT" + Format(CURY, "0000") + ".MDB"
dbn2 = APPROOT + "\ARTS\WORK\WGT" + Format(CURY, "0000") + ".MDB"

FileCopy APPROOT + "\ARTS\STRUS\ARTS.MDB", dbn2

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(dbn1)
Set prm_record = prm_database.OpenRecordset("ASITAB")

Dim prm2_database As Database, prm2_record As Recordset

Set prm2_database = OpenDatabase(dbn2)
Set prm2_record = prm2_database.OpenRecordset("ASITAB")

prm2_record.Index = "primarykey"

With prm_record

NREC = .RecordCount: IREC = 0

pgbCOMP.Min = 0: pgbCOMP.Max = NREC
pgbCOMP.Visible = True

.Index = "primarykey"

.MoveFirst

Do Until .EOF

IREC = IREC + 1: pgbCOMP.Value = IREC

XKEY1 = ![akey]
XKEY2 = "J0000+M0000" + Right(XKEY1, 12)

prm2_record.Seek "=", XKEY2

If prm2_record.NoMatch = False Then GoTo UPDATE_OLD

prm2_record.AddNew

prm2_record![akey] = XKEY2

For I = 1 To 13

ix = Format(I, "00")

FFC = "C" + ix:  prm2_record.Fields(FFC) = .Fields(FFC)
FFE = "E" + ix:  prm2_record.Fields(FFE) = .Fields(FFE)
FFU = "U" + ix:  prm2_record.Fields(FFU) = .Fields(FFU)
FFP = "P" + ix:  prm2_record.Fields(FFP) = .Fields(FFP)
FFV = "V" + ix:  prm2_record.Fields(FFV) = .Fields(FFV)
FFW = "W" + ix:  prm2_record.Fields(FFW) = .Fields(FFW)
FFN = "F" + ix:  prm2_record.Fields(FFN) = .Fields(FFN)

Next I

prm2_record![RANK] = 0
prm2_record![PER] = 0
prm2_record![CUM] = 0

prm2_record.Update

GoTo NEXT_REC

UPDATE_OLD:

prm2_record.Edit

For I = 1 To 13

ix = Format(I, "00")

FFC = "C" + ix:  prm2_record.Fields(FFC) = prm2_record.Fields(FFC) + .Fields(FFC)
FFE = "E" + ix:  prm2_record.Fields(FFE) = prm2_record.Fields(FFE) + .Fields(FFE)
FFU = "U" + ix:  prm2_record.Fields(FFU) = prm2_record.Fields(FFU) + .Fields(FFU)
FFP = "P" + ix:  prm2_record.Fields(FFP) = prm2_record.Fields(FFP) + .Fields(FFP)
FFV = "V" + ix:  prm2_record.Fields(FFV) = prm2_record.Fields(FFV) + .Fields(FFV)
FFW = "W" + ix:  prm2_record.Fields(FFW) = prm2_record.Fields(FFW) + .Fields(FFW)
FFN = "F" + ix:  prm2_record.Fields(FFN) = prm2_record.Fields(FFN) + .Fields(FFN)

prm2_record.Fields(FFU) = 0
prm2_record.Fields(FFP) = 0
prm2_record.Fields(FFW) = 0

If prm2_record.Fields(FFE) <> 0 Then
   prm2_record.Fields(FFU) = prm2_record.Fields(FFC) / prm2_record.Fields(FFE)
   End If

If prm2_record.Fields(FFC) <> 0 Then
   prm2_record.Fields(FFP) = prm2_record.Fields(FFV) / prm2_record.Fields(FFC)
   End If

If prm2_record.Fields(FFN) <> 0 Then
   prm2_record.Fields(FFW) = prm2_record.Fields(FFC) / prm2_record.Fields(FFN)
   End If

Next I

prm2_record![RANK] = 0
prm2_record![PER] = 0
prm2_record![CUM] = 0

prm2_record.Update

NEXT_REC:

.MoveNext

Loop

End With

prm_record.Close
prm_database.Close

prm2_record.Close
prm2_database.Close

FileCopy dbn2, dbn1

End Sub
Private Sub GROUPS_BYGT()

Dim I, ix, FFC, FFE, FFP, FFU, FFV, RR, cc, PP

Dim dbn1, dbn2, XKEY1, XKEY2, NREC, IREC

dbn1 = APPROOT + "\ARTS\WORK\WS" + Format(CURY, "0000") + ".MDB"
dbn2 = APPROOT + "\ARTS\WORK\WGS" + Format(CURY, "0000") + ".MDB"

FileCopy APPROOT + "\ARTS\STRUS\ARTS.MDB", dbn2

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(dbn1)
Set prm_record = prm_database.OpenRecordset("ASITAB")

Dim prm2_database As Database, prm2_record As Recordset

Set prm2_database = OpenDatabase(dbn2)
Set prm2_record = prm2_database.OpenRecordset("ASITAB")

prm2_record.Index = "primarykey"

With prm_record

NREC = .RecordCount: IREC = 0

pgbCOMP.Min = 0: pgbCOMP.Max = NREC
pgbCOMP.Visible = True

.Index = "primarykey"

.MoveFirst

Do Until .EOF

IREC = IREC + 1: pgbCOMP.Value = IREC

XKEY1 = ![akey]
XKEY2 = "J0000+M0000" + Right(XKEY1, 12)

prm2_record.Seek "=", XKEY2

If prm2_record.NoMatch = False Then GoTo UPDATE_OLD

prm2_record.AddNew

prm2_record![akey] = XKEY2

For I = 1 To 13

ix = Format(I, "00")

FFC = "C" + ix:  prm2_record.Fields(FFC) = .Fields(FFC)
FFE = "E" + ix:  prm2_record.Fields(FFE) = .Fields(FFE)
FFU = "U" + ix:  prm2_record.Fields(FFU) = .Fields(FFU)
FFP = "P" + ix:  prm2_record.Fields(FFP) = .Fields(FFP)
FFV = "V" + ix:  prm2_record.Fields(FFV) = .Fields(FFV)
FFW = "W" + ix:  prm2_record.Fields(FFW) = .Fields(FFW)
FFN = "F" + ix:  prm2_record.Fields(FFN) = .Fields(FFN)

Next I

prm2_record![RANK] = 0
prm2_record![PER] = 0
prm2_record![CUM] = 0

prm2_record![ADDFISH] = ![ADDFISH]

prm2_record.Update

GoTo NEXT_REC

UPDATE_OLD:

prm2_record.Edit

For I = 1 To 13

ix = Format(I, "00")

FFC = "C" + ix:  prm2_record.Fields(FFC) = prm2_record.Fields(FFC) + .Fields(FFC)
FFE = "E" + ix:  prm2_record.Fields(FFE) = prm2_record.Fields(FFE) + .Fields(FFE)
FFU = "U" + ix:  prm2_record.Fields(FFU) = prm2_record.Fields(FFU) + .Fields(FFU)
FFP = "P" + ix:  prm2_record.Fields(FFP) = prm2_record.Fields(FFP) + .Fields(FFP)
FFV = "V" + ix:  prm2_record.Fields(FFV) = prm2_record.Fields(FFV) + .Fields(FFV)
FFW = "W" + ix:  prm2_record.Fields(FFW) = prm2_record.Fields(FFW) + .Fields(FFW)
FFN = "F" + ix:  prm2_record.Fields(FFN) = prm2_record.Fields(FFN) + .Fields(FFN)

prm2_record.Fields(FFU) = 0
prm2_record.Fields(FFP) = 0
prm2_record.Fields(FFW) = 0

If prm2_record.Fields(FFE) <> 0 Then
   prm2_record.Fields(FFU) = prm2_record.Fields(FFC) / prm2_record.Fields(FFE)
   End If

If prm2_record.Fields(FFC) <> 0 Then
   prm2_record.Fields(FFP) = prm2_record.Fields(FFV) / prm2_record.Fields(FFC)
   End If

If prm2_record.Fields(FFN) <> 0 Then
   prm2_record.Fields(FFW) = prm2_record.Fields(FFC) / prm2_record.Fields(FFN)
   End If

Next I

prm2_record![RANK] = 0
prm2_record![PER] = 0
prm2_record![CUM] = 0

If ![ADDFISH] = "NO" Then prm2_record![ADDFISH] = "NO"

If prm2_record![ADDFISH] = "YES" And ![ADDFISH] = "NO" Then prm2_record![ADDFISH] = "NO"

prm2_record.Update

NEXT_REC:

.MoveNext

Loop

End With

prm_record.Close
prm_database.Close

prm2_record.Close
prm2_database.Close

FileCopy dbn2, dbn1

End Sub
Private Sub WRITE_LOG()

Dim fnm

fnm = APPROOT + "\ARTS\WORK\WSEL.TXT"

Open fnm For Append As #1

Print #1, " "
Print #1, msgtab(45)
Print #1, String(40, "=")
Print #1, WLEV
Print #1, WCOMP

Close #1

rtsSEL.FileName = APPROOT + "\ARTS\WORK\WSEL.TXT"
rtsSEL.Refresh

End Sub
Private Sub optP_Click(Index As Integer)

Dim I

I = Index

If I = 12 Then
   WELM = "P"
   WPER = CURY
   WMY = "Y"
   Call RANK_PROC_CLEAN
   Exit Sub
   End If

WELM = "P"
WPER = I + 1
WMY = "M"
Call RANK_PROC_CLEAN
Exit Sub

End Sub
Private Sub optQUIT_Click()

Frame2.Visible = False

cmdREP.Visible = True
cmdGROUP.Visible = True
cmdRANK.Visible = True

End Sub
Private Sub RANK_PROC_CLEAN()

optQUIT.Value = False
optQUIT.Enabled = False

Dim I

If WELM = "C" Then

   For I = 1 To 13
   
   optE(I - 1).Enabled = False
   optU(I - 1).Enabled = False
   optP(I - 1).Enabled = False
   optV(I - 1).Enabled = False
   
   If I = WPER Then GoTo NEXT_C
   optC(I - 1).Enabled = False

NEXT_C:
   
   Next I
   
   End If
   
If WELM = "E" Then

   For I = 1 To 13
   optC(I - 1).Enabled = False
   optU(I - 1).Enabled = False
   optP(I - 1).Enabled = False
   optV(I - 1).Enabled = False
   If I = WPER Then GoTo NEXT_E
   optE(I - 1).Enabled = False

NEXT_E:
   
   Next I
   
   End If

If WELM = "U" Then

   For I = 1 To 13
   optC(I - 1).Enabled = False
   optE(I - 1).Enabled = False
   optP(I - 1).Enabled = False
   optV(I - 1).Enabled = False
   If I = WPER Then GoTo NEXT_U
   optU(I - 1).Enabled = False

NEXT_U:
   
   Next I
   
   End If

If WELM = "P" Then

   For I = 1 To 13
   optC(I - 1).Enabled = False
   optE(I - 1).Enabled = False
   optU(I - 1).Enabled = False
   optV(I - 1).Enabled = False
   If I = WPER Then GoTo NEXT_P
   optP(I - 1).Enabled = False

NEXT_P:
   
   Next I
   
   End If

If WELM = "V" Then

   For I = 1 To 13
   optC(I - 1).Enabled = False
   optE(I - 1).Enabled = False
   optU(I - 1).Enabled = False
   optP(I - 1).Enabled = False
   If I = WPER Then GoTo NEXT_V
   optV(I - 1).Enabled = False

NEXT_V:
   
   Next I
   
   End If

optQUIT.Enabled = False

Call RANK_MAIN

End Sub
Private Sub optU_Click(Index As Integer)

Dim I

I = Index

If I = 12 Then
   WELM = "U"
   WPER = CURY
   WMY = "Y"
   Call RANK_PROC_CLEAN
   Exit Sub
   End If

WELM = "U"
WPER = I + 1
WMY = "M"
Call RANK_PROC_CLEAN
Exit Sub

End Sub
Private Sub optV_Click(Index As Integer)

Dim I

I = Index

If I = 12 Then
   WELM = "V"
   WPER = CURY
   WMY = "Y"
   Call RANK_PROC_CLEAN
   Exit Sub
   End If

WELM = "V"
WPER = I + 1
WMY = "M"
Call RANK_PROC_CLEAN
Exit Sub

End Sub
Private Sub RANK_MAIN()

Dim DBN, XKEY, TEST_STRING, RES_TEST

DBN = APPROOT + "\ARTS\WORK\WGS" + Format(CURY, "0000") + ".MDB"

If Dir(DBN) = "" Then Exit Sub

pgbCOMP.Visible = True

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(DBN)
Set prm_record = prm_database.OpenRecordset("ASITAB")

With prm_record

.Index = "primarykey"

.MoveFirst

Dim NREC, IREC, FFF, XXX, WTOT, WPERC, WCUM

IREC = 0

NREC = .RecordCount

pgbCOMP.Min = 0: pgbCOMP.Max = NREC

Do Until .EOF

IREC = IREC + 1

pgbCOMP.Value = IREC

If ![RANK] = -999 Then .Delete

.MoveNext

Loop

.MoveFirst

IREC = 0: WTOT = 0

Do Until .EOF

IREC = IREC + 1

pgbCOMP.Value = IREC

If WMY = "M" Then FFF = WELM + Format(WPER, "00")
If WMY = "Y" Then FFF = WELM + "13"

.Edit

![RANK] = -(.Fields(FFF)): WTOT = WTOT + .Fields(FFF)

If WELM = "C" Then XXX = RTrim(msgtab(63))
If WELM = "E" Then XXX = RTrim(msgtab(64))
If WELM = "U" Then XXX = RTrim(msgtab(65))
If WELM = "P" Then XXX = RTrim(msgtab(66))
If WELM = "V" Then XXX = RTrim(msgtab(67))

If WMY = "Y" Then XXX = XXX + " " + Format(CURY, "0000")
If WMY = "M" Then XXX = XXX + " " + Format(WPER, "00") + "/" + Format(CURY, "0000")

![CRIT] = FFF + ":" + XXX

.Update

.MoveNext

Loop

.Index = "RANK"

.MoveFirst

IREC = 0

Do Until .EOF

IREC = IREC + 1

pgbCOMP.Value = IREC

.Edit

![PER] = 0

If WTOT <> 0 Then ![PER] = 100 * Abs(![RANK]) / WTOT

If IREC = 1 Then
   WCUM = ![PER]: ![CUM] = ![PER]
   End If

If IREC > 1 Then
   WCUM = WCUM + ![PER]: ![CUM] = WCUM
   End If

XKEY = ![akey]

If WELM = "U" Or WELM = "P" Then
   ![PER] = 0: ![CUM] = 0
   End If

If WELM = "E" Then
   RES_TEST = InStr(XKEY, "S0000")
   If RES_TEST = 0 Then
      ![PER] = 0: ![CUM] = 0
      End If
   End If

.Update

.MoveNext

Loop

End With

prm_record.Close
prm_database.Close

pgbCOMP.Visible = False
Frame2.Visible = False
cmdREP.Visible = True


'ADD GRAND TOTALS
'================

RANK_TOTAL = "Y"

Call TOTAL_CATCH
Call TOTAL_EFFORT

RANK_TOTAL = "N"

Dim fnm

fnm = APPROOT + "\ARTS\WORK\WSEL.TXT"

Open fnm For Append As #1

Print #1, " "

Print #1, msgtab(85)
Print #1, String(40, "=")
Print #1, XXX

RANK_CRIT = XXX

Close #1

rtsSEL.FileName = fnm
rtsSEL.Refresh

RANKYN = "Y"

pgbCOMP.Visible = False

End Sub
