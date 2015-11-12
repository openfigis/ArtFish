VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmLAND 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12015
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   1  'Arrow
   Moveable        =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   12015
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6975
      Left            =   7560
      TabIndex        =   8
      Top             =   600
      Width           =   4335
      Begin VB.CommandButton cmdGUIDE2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   3480
         MousePointer    =   1  'Arrow
         Picture         =   "frmLAND.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   65
         Top             =   6120
         Width           =   735
      End
      Begin VB.TextBox txtREC 
         Appearance      =   0  'Flat
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
         Height          =   285
         Left            =   720
         TabIndex        =   20
         Top             =   6600
         Width           =   2655
      End
      Begin VB.ListBox lstSUM 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   6.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   2940
         ItemData        =   "frmLAND.frx":2262
         Left            =   120
         List            =   "frmLAND.frx":2264
         TabIndex        =   19
         Top             =   360
         Width           =   4095
      End
      Begin VB.TextBox txtNOU 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         TabIndex        =   18
         Top             =   5160
         Width           =   1335
      End
      Begin VB.TextBox txtDUR 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         TabIndex        =   17
         Top             =   5400
         Width           =   1335
      End
      Begin VB.TextBox txtWSMP 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   1680
         TabIndex        =   16
         Top             =   5640
         Width           =   1335
      End
      Begin VB.TextBox txtWTOT 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1680
         TabIndex        =   15
         Top             =   5880
         Width           =   1335
      End
      Begin VB.CommandButton cmdADD 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         MousePointer    =   1  'Arrow
         Picture         =   "frmLAND.frx":2266
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   4320
         Width           =   495
      End
      Begin VB.TextBox txtDAY 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         TabIndex        =   13
         Top             =   4920
         Width           =   1335
      End
      Begin VB.CommandButton cmdSTAY 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         Picture         =   "frmLAND.frx":24E8
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   3720
         Width           =   495
      End
      Begin VB.CommandButton cmdLEAVE 
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         MousePointer    =   1  'Arrow
         Picture         =   "frmLAND.frx":25F2
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   6360
         Width           =   495
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
         Height          =   495
         Left            =   120
         MousePointer    =   1  'Arrow
         Picture         =   "frmLAND.frx":2874
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   5400
         Width           =   495
      End
      Begin VB.CommandButton cmdDEL 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   3720
         Picture         =   "frmLAND.frx":2AF6
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   4320
         Width           =   495
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
         TabIndex        =   64
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
         Index           =   29
         Left            =   8205
         TabIndex        =   63
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
         Index           =   28
         Left            =   8205
         TabIndex        =   62
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
         Index           =   27
         Left            =   8205
         TabIndex        =   61
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
         Index           =   26
         Left            =   8205
         TabIndex        =   60
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
         Index           =   25
         Left            =   8205
         TabIndex        =   59
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
         Index           =   24
         Left            =   8205
         TabIndex        =   58
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
         Index           =   23
         Left            =   5565
         TabIndex        =   57
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
         Index           =   22
         Left            =   5565
         TabIndex        =   56
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
         Index           =   21
         Left            =   5565
         TabIndex        =   55
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
         Index           =   20
         Left            =   5565
         TabIndex        =   54
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
         Index           =   19
         Left            =   5565
         TabIndex        =   53
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
         Index           =   18
         Left            =   5565
         TabIndex        =   52
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
         Index           =   17
         Left            =   5565
         TabIndex        =   51
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
         Index           =   16
         Left            =   5565
         TabIndex        =   50
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
         Index           =   15
         Left            =   2925
         TabIndex        =   49
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
         Index           =   14
         Left            =   2925
         TabIndex        =   48
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
         Index           =   13
         Left            =   2925
         TabIndex        =   47
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
         Index           =   12
         Left            =   2925
         TabIndex        =   46
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
         Index           =   11
         Left            =   2925
         TabIndex        =   45
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
         Index           =   10
         Left            =   2925
         TabIndex        =   44
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
         Index           =   9
         Left            =   2925
         TabIndex        =   43
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
         Index           =   8
         Left            =   2925
         TabIndex        =   42
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
         Index           =   7
         Left            =   285
         TabIndex        =   41
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
         Index           =   6
         Left            =   285
         TabIndex        =   40
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
         Index           =   5
         Left            =   285
         TabIndex        =   39
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
         Index           =   4
         Left            =   285
         TabIndex        =   38
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
         Index           =   3
         Left            =   285
         TabIndex        =   37
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
         Index           =   2
         Left            =   285
         TabIndex        =   36
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
         Index           =   1
         Left            =   285
         TabIndex        =   35
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
         Index           =   0
         Left            =   285
         TabIndex        =   34
         Top             =   480
         Width           =   90
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
         Left            =   720
         TabIndex        =   33
         Top             =   6360
         Width           =   2655
      End
      Begin VB.Label lblNOU 
         Alignment       =   1  'Right Justify
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
         Left            =   120
         TabIndex        =   32
         Top             =   5160
         Width           =   1455
      End
      Begin VB.Label lblDUR 
         Alignment       =   1  'Right Justify
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
         Left            =   120
         TabIndex        =   31
         Top             =   5400
         Width           =   1455
      End
      Begin VB.Label lblSMP 
         Alignment       =   1  'Right Justify
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
         Left            =   120
         TabIndex        =   30
         Top             =   5640
         Width           =   1455
      End
      Begin VB.Label lblTOT 
         Alignment       =   1  'Right Justify
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
         Left            =   120
         TabIndex        =   29
         Top             =   5880
         Width           =   1455
      End
      Begin VB.Label lblDOC 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "000000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   495
         Left            =   1680
         TabIndex        =   28
         Top             =   4320
         Width           =   1335
      End
      Begin VB.Label lblDAY 
         Alignment       =   1  'Right Justify
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
         Left            =   120
         TabIndex        =   27
         Top             =   4920
         Width           =   1455
      End
      Begin VB.Label lblSUM 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   6.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   120
         Width           =   5055
      End
      Begin VB.Label lblDAY2 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00808000&
         Height          =   255
         Left            =   3120
         TabIndex        =   25
         Top             =   4920
         Width           =   855
      End
      Begin VB.Label lblTOT2 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00808000&
         Height          =   255
         Left            =   3120
         TabIndex        =   24
         Top             =   5880
         Width           =   855
      End
      Begin VB.Label lblSMP2 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00808000&
         Height          =   255
         Left            =   3120
         TabIndex        =   23
         Top             =   5640
         Width           =   855
      End
      Begin VB.Label lblDUR2 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00808000&
         Height          =   255
         Left            =   3120
         TabIndex        =   22
         Top             =   5400
         Width           =   855
      End
      Begin VB.Label lblNOU2 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00808000&
         Height          =   255
         Left            =   3120
         TabIndex        =   21
         Top             =   5160
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdGUIDE1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   9960
      MousePointer    =   1  'Arrow
      Picture         =   "frmLAND.frx":2D78
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6840
      Width           =   735
   End
   Begin VB.CommandButton cmdBACK 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   10920
      MousePointer    =   1  'Arrow
      Picture         =   "frmLAND.frx":4FDA
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6840
      Width           =   735
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
      Height          =   6060
      Left            =   120
      MultiSelect     =   2  'Extended
      TabIndex        =   3
      Top             =   480
      Width           =   11775
   End
   Begin VB.CommandButton cmdCONFIRM 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      Picture         =   "frmLAND.frx":525C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3960
      Width           =   495
   End
   Begin VB.Data dtaGEN 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\ARTBAS\STRUS\Dspecies.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "STAB"
      Top             =   0
      Width           =   2175
   End
   Begin VB.ListBox lstMINOR 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3180
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   4335
   End
   Begin MSDBGrid.DBGrid dbgGEN 
      Bindings        =   "frmLAND.frx":5366
      Height          =   6975
      Left            =   120
      OleObjectBlob   =   "frmLAND.frx":5388
      TabIndex        =   0
      Top             =   600
      Width           =   7455
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FFFF&
      Caption         =   " 09"
      Height          =   255
      Left            =   0
      TabIndex        =   66
      Top             =   7560
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
      TabIndex        =   6
      Top             =   120
      Width           =   11775
   End
   Begin VB.Label lblSELMN 
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
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
      TabIndex        =   5
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "frmLAND"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private MNCODE(), MNAME(), MNCAL(), NMN, MNSEL()
Private SICODE(), SINAME(), SISEQ(), BGCODE(), BGNAME(), BGSEQ(), NSI, NBG
Private NAS, ASIC(), ASIN()
Private NFR, FRC(), FRN(), FRNO()
Private CURSIBG, CURSIBGC, OLDREC
Private OKMN(), OKSIBG()
Private SIMN(), NSP, SPEC(), SPEN(), SPSEQ(), SUMERR, ERRTOT
Private DECF, LSTAB(), NLST, CURSEL
Private NEDIT, LTC(), LTL(), LTF(), LTP(), LTV()
Private CURDOC, CURNOU, CURDUR, CURDAY, CURSMP, CURTOT, CURREC
Private FIND_STRING, WKEY_CODE, WKEY_CHAR
Private Sub cmdGUIDE1_Click()

HTYPE = "70"

If lstMINOR.Visible = False Then HTYPE = "80"
   
HFNM = APPROOT + "\ARTBAS\HELP\" + current_language + "HELP" + HTYPE + ".rtf"

If Dir(HFNM) = "" Then Exit Sub

frmLAND.Enabled = False
Load frmGUIDE
frmGUIDE.Show

End Sub
Private Sub cmdGUIDE2_Click()

HTYPE = "90"

If dbgGEN.Visible = True Then HTYPE = "A0"
   
HFNM = APPROOT + "\ARTBAS\HELP\" + current_language + "HELP" + HTYPE + ".rtf"

If Dir(HFNM) = "" Then Exit Sub

frmLAND.Enabled = False
Load frmGUIDE
frmGUIDE.Show

End Sub
Private Sub cmdADD_Click()

cmdLEAVE.Enabled = False

lblDAY2.Visible = False
lblNOU2.Visible = False
lblDUR2.Visible = False
lblSMP2.Visible = False
lblTOT2.Visible = False

NEDIT = 0

If Len(txtWSMP.Text) = 0 Then
   cmdLEAVE.Enabled = True

   Dim resp
   
   resp = MsgBox(msgtab(99), vbCritical, " ")
   
   Exit Sub
   End If

dbgGEN.Visible = True

If IsNumeric(txtWSMP.Text) = True And CDbl(txtWSMP.Text) = 0# Then
    
   resp = MsgBox(msgtab(99), vbCritical, " ")
   cmdLEAVE.Enabled = True

   Exit Sub
   End If

If IsNumeric(txtWSMP.Text) = True And CDbl(txtWSMP.Text) < 0.001 Then
   txtWSMP.Text = 0.001: txtWTOT.Text = 0.001
   End If
   
If Dir(APPROOT + "\ARTBAS\LANDINGS\WORK.MDB") = "" Then
   FileCopy APPROOT + "\ARTBAS\STRUS\LSAMPLES.MDB", APPROOT + "\ARTBAS\LANDINGS\WORK.MDB"
   FileCopy APPROOT + "\ARTBAS\STRUS\LSPECIES.MDB", APPROOT + "\ARTBAS\LANDINGS\WORK2.MDB"
   End If
 
If Len(txtWTOT.Text) = 0 Then txtWTOT.Text = 0
 
If CDbl(txtWTOT.Text) < CDbl(txtWSMP.Text) Then txtWTOT.Text = txtWSMP.Text

If Len(txtDAY.Text) = 0 Then txtDAY.Text = 1
If Len(txtNOU.Text) = 0 Then txtNOU.Text = 1
If Len(txtDUR.Text) = 0 Then txtDUR.Text = 1
If Len(txtREC.Text) = 0 Then txtREC.Text = "???"

Dim I, J, K, fnm, dbn

dbn = APPROOT + "\ARTBAS\LANDINGS\WORK.MDB"

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(dbn)
Set prm_record = prm_database.OpenRecordset("LTAB")

With prm_record

.Index = "primarykey"

.AddNew

lblDOC.Caption = Format(![LDOC], "000000")
lblDOC.Visible = True

![LDAY] = CDbl(txtDAY.Text)
![LNOU] = CDbl(txtNOU.Text)
![LDUR] = CDbl(txtDUR.Text)
![LSMP] = CDbl(txtWSMP.Text)

If Len(txtWTOT.Text) = 0 Then txtWTOT.Text = txtWSMP.Text

![ltot] = CDbl(txtWTOT.Text)

If ![LSMP] < 0.001 Then
   ![LSMP] = 0.001: ![ltot] = 0.001
   End If

![LREC] = Left(txtREC.Text, 15)
![LMNC] = CURMNC
![lsbc] = CURSIBGC

.Update

End With

Dim WWW, ZZZ

WWW = CDbl(txtWSMP.Text): WWW = Format(WWW, "#####0.00")
WWW = Right(Space(15) + LTrim(WWW), 9)

ZZZ = WWW

WWW = CDbl(txtWTOT.Text): WWW = Format(WWW, "#####0.00")
WWW = Right(Space(15) + LTrim(WWW), 9)

lstSUM.Enabled = False

NLST = NLST + 1

ReDim Preserve LSTAB(1 To NLST)

LSTAB(NLST) = Left(lblDOC.Caption, 6) + " " + _
                Format(txtDAY.Text, "00") + _
                Right(Space(10) + LTrim(Format(txtNOU, "##0.00")), 7) + " " + _
                Right(Space(10) + LTrim(Format(txtDUR, "##0.00")), 7) + " " + _
                ZZZ + " " + WWW

lstSUM.AddItem LSTAB(NLST)

lstSUM.ItemData(lstSUM.NewIndex) = lblDOC.Caption

CURSEL = lstSUM.NewIndex + 1

If lstSUM.ListCount > 5 Then lstSUM.TopIndex = lstSUM.ListCount - 5

If CDbl(txtWSMP.Text) = 0.001 Then
   cmdLEAVE.Enabled = True
   cmdPRINT.Visible = False
   lstSUM.Enabled = True
   dbgGEN.Visible = False
   txtWSMP.Text = ""
   txtWTOT.Text = ""
   Exit Sub
   End If
  
cmdPRINT.Visible = False
cmdADD.Visible = False
cmdSTAY.Visible = True

Call CREATE_SPEDB

dtaGEN.DatabaseName = APPROOT + "\ARTBAS\LANDINGS\WORKSP.MDB"
dtaGEN.Refresh

End Sub
Private Sub cmdBACK_Click()

cmdBACK.MousePointer = 13

dtaGEN.DatabaseName = APPROOT + "\ARTBAS\STRUS\DSPECIES.MDB"
dtaGEN.Refresh

If Dir(APPROOT + "\ARTBAS\LANDINGS\WORK.MDB") <> "" Then Kill APPROOT + "\ARTBAS\LANDINGS\WORK.MDB"
If Dir(APPROOT + "\ARTBAS\LANDINGS\WORK2.MDB") <> "" Then Kill APPROOT + "\ARTBAS\LANDINGS\WORK2.MDB"
If Dir(APPROOT + "\ARTBAS\LANDINGS\WORKSP.MDB") <> "" Then Kill APPROOT + "\ARTBAS\LANDINGS\WORKSP.MDB"

Dim fnm, NN, XXX

NN = 0

fnm = APPROOT + "\ARTBAS\LANDINGS\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_LSAMPLES.TXT"

If Dir(fnm) = "" Then GoTo CONT_RET

Open fnm For Input As #1

Do Until EOF(1)
Line Input #1, XXX
NN = NN + 1
Loop

Close #1

If NN = 0 Then Kill fnm

CONT_RET:

frmLAND.MousePointer = 13
Load frmARTB01
Unload frmLAND
frmARTB01.Show

End Sub
Private Sub cmdCONFIRM_Click()

CURMNC = MNCODE(lstMINOR.ListIndex + 1)
CURMNN = MNAME(lstMINOR.ListIndex + 1)

frmLAND.Caption = monthtab(current_month) + " " + _
                    Format(current_year, "0000") + " - " + msgtab(87) + _
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

cmdSTAY.Visible = False
dbgGEN.Visible = False

Dim dbn, I, XKEY
Dim prm_database As Database, prm_record As Recordset

dbn = APPROOT + "\ARTBAS\LANDINGS\WORK2.MDB"

Set prm_database = OpenDatabase(dbn)
Set prm_record = prm_database.OpenRecordset("STAB")

With prm_record

.Index = "primarykey"

For I = 1 To NSP

XKEY = "D" + Left(lblDOC.Caption, 6) + "+S" + SPEC(I)

.Seek "=", XKEY

If .NoMatch = True Then GoTo CONT_LOOP

.Edit

![slan] = 0
![snof] = 0
![spri] = 0
![sval] = 0

.Update

CONT_LOOP:

Next I

End With

prm_record.Close
prm_database.Close

dbn = APPROOT + "\ARTBAS\LANDINGS\WORK.MDB"

Set prm_database = OpenDatabase(dbn)
Set prm_record = prm_database.OpenRecordset("LTAB")

With prm_record

.Index = "primarykey"

XKEY = Left(lblDOC.Caption, 6)

.Seek "=", XKEY

If .NoMatch = True Then
   prm_record.Close
   prm_database.Close
   Exit Sub
   End If
   
.Edit

![LDAY] = 0
![LSMP] = 0

.Update

End With

prm_record.Close
prm_database.Close

Call cmdLEAVE_Click

End Sub

Private Sub cmdLEAVE_Click()

lstSIBG.Visible = True

cmdLEAVE.MousePointer = 13

lblDOC.Visible = False
lstSUM.Enabled = True

cmdBACK.Visible = True

cmdPRINT.Visible = False
cmdSTAY.Visible = False
cmdLEAVE.Visible = False

Frame1.Visible = False
dbgGEN.Visible = False

lstSIBG.Enabled = True

NFR = 0

Call DUMP_WORK
Call DUMP_WORK2

Call LOAD_FRAME

cmdLEAVE.MousePointer = 1

End Sub
Private Sub cmdPRINT_Click()

Dim dbn

dbn = APPROOT + "\ARTBAS\LANDINGS\WORKSP.MDB"

If Dir(dbn) = "" Then Exit Sub

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(dbn)
Set prm_record = prm_database.OpenRecordset("STAB")

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

If ![slan] = 0 Then GoTo CONT_READ

Printer.Print Tab(5); Left(RTrim(![sdes]) + String(30, "."), 30); _
              Tab(36); Right(Space(10) + LTrim(Format(![slan], "#####0.000")), 10); _
              Tab(47); Right(Space(4) + LTrim(Format(![snof], "###0")), 4); _
              Tab(52); Right(Space(12) + LTrim(Format(![spri], "#####0.000")), 12); _
              Tab(65); Right(Space(13) + LTrim(Format(![sval], "########0.000")), 13)
                                         
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

Printer.Print Tab(5); frmLAND.Caption

Printer.Print

Printer.Print Tab(5); lblSELSIBG.Caption
Printer.Print Tab(5); String(80, "-")

Printer.Print

Dim PPP

Printer.Print

PPP = Left(msgtab(120) + String(15, "."), 15) + " " + CURDOC
Printer.Print Tab(5); PPP

Printer.Print

PPP = Left(msgtab(82) + String(15, "."), 15) + " " + CURREC
Printer.Print Tab(5); PPP

PPP = Left(msgtab(83) + String(15, "."), 15) + " " + CURDAY
Printer.Print Tab(5); PPP

PPP = Left(msgtab(88) + String(15, "."), 15) + " " + CURNOU
Printer.Print Tab(5); PPP

PPP = Left(msgtab(89) + String(15, "."), 15) + " " + CURDUR
Printer.Print Tab(5); PPP

PPP = Left(msgtab(90) + String(15, "."), 15) + " " + CURSMP
Printer.Print Tab(5); PPP

PPP = Left(msgtab(91) + String(15, "."), 15) + " " + CURTOT
Printer.Print Tab(5); PPP

Printer.Print

Printer.Print Tab(5); msgtab(47); _
              Tab(36); Right(Space(10) + RTrim(UNW), 10); _
              Tab(47); msgtab(93); _
              Tab(52); Right(Space(12) + RTrim(UNM) + "/" + RTrim(UNW), 12); _
              Tab(65); Right(Space(13) + RTrim(UNM), 13)
                               
Printer.Print

Return

End Sub
Private Sub cmdSTAY_Click()

cmdLEAVE.Enabled = True

Dim resp

If Len(RTrim(txtREC.Text)) = 0 Or txtREC.Text = "???" Then
   txtREC.Text = "???"
   resp = MsgBox(msgtab(169), vbCritical + vbOKOnly, " ")
   Exit Sub
   End If

cmdSTAY.MousePointer = 13

If Len(txtWSMP.Text) = 0 Then
   
   cmdSTAY.MousePointer = 1
      
   resp = MsgBox(msgtab(99), vbCritical, " ")
   
   Exit Sub
   End If

dbgGEN.Visible = True

If IsNumeric(txtWSMP.Text) = True And CDbl(txtWSMP.Text) = 0# Then
   cmdSTAY.MousePointer = 1
   resp = MsgBox(msgtab(99), vbCritical, " ")
   
   Exit Sub
   End If

If IsNumeric(txtWSMP.Text) = True And CDbl(txtWSMP.Text) < 0.001 Then
   txtWSMP.Text = 0.001: txtWTOT.Text = 0.001
   End If
    
If Len(txtWTOT.Text) = 0 Then txtWTOT.Text = 0
    
If CDbl(txtWTOT.Text) < CDbl(txtWSMP.Text) Then txtWTOT.Text = txtWSMP.Text
    
If Len(txtDAY.Text) = 0 Then txtDAY.Text = 1
If Len(txtNOU.Text) = 0 Then txtNOU.Text = 1
If Len(txtDUR.Text) = 0 Then txtDUR.Text = 1
If Len(txtREC.Text) = 0 Then txtREC.Text = "???"

Dim dbn, ddd

dbn = APPROOT + "\ARTBAS\LANDINGS\WORK.MDB"

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(dbn)
Set prm_record = prm_database.OpenRecordset("LTAB")

With prm_record

.Index = "primarykey"

ddd = CDbl(lblDOC.Caption)

.Seek "=", ddd

If .NoMatch = True Then Exit Sub

.Edit

![LDAY] = CDbl(txtDAY.Text)
![LNOU] = CDbl(txtNOU.Text)
![LDUR] = CDbl(txtDUR.Text)
![LSMP] = CDbl(txtWSMP.Text)

If Len(txtWTOT.Text) = 0 Then ![ltot] = ![LSMP]
If Len(txtWTOT.Text) <> 0 Then ![ltot] = CDbl(txtWTOT.Text)

If ![ltot] < ![LSMP] Then ![ltot] = ![LSMP]

If ![LSMP] < 0.001 Then
   ![LSMP] = 0.001: ![ltot] = 0.001
   End If

![LREC] = Left(txtREC.Text, 15)

.Update

End With

prm_record.Close
prm_database.Close

Call COPY_SPECIES

Call CONTROL_TOTALS

If SUMERR = "Y" Then
   cmdLEAVE.Visible = False
   cmdPRINT.Visible = False
   cmdADD.Visible = False
   dbgGEN.Refresh
   Dim msg
   msg = msgtab(101) + Chr(13) + msgtab(102) + Str(ERRTOT) + Chr(13) + msgtab(103) + txtWSMP.Text
   
   resp = MsgBox(msg, vbCritical, " ")
   cmdSTAY.MousePointer = 1
   Exit Sub
   
   End If

cmdDEL.Visible = False
cmdPRINT.Visible = True
cmdLEAVE.Visible = True
cmdADD.Visible = True
cmdSTAY.Visible = False
lstSUM.Enabled = True

dbgGEN.Visible = False
dbgGEN.Refresh

lstSUM.Refresh

Dim WWW, ZZZ

WWW = CDbl(txtWSMP.Text)
WWW = Format(WWW, "#####0.00")
WWW = Right(Space(10) + LTrim(WWW), 9)

ZZZ = WWW

If Len(txtWTOT.Text) = 0 Then txtWTOT.Text = txtWSMP.Text

WWW = CDbl(txtWTOT.Text)
WWW = Format(WWW, "#####0.00")
WWW = Right(Space(10) + LTrim(WWW), 9)

LSTAB(CURSEL) = Left(lblDOC.Caption, 6) + " " + _
                Format(txtDAY.Text, "00") + _
                Right(Space(10) + LTrim(Format(txtNOU, "##0.00")), 7) + " " + _
                Right(Space(10) + LTrim(Format(txtDUR, "##0.00")), 7) + " " + _
                ZZZ + " " + WWW

Call CREATE_SUMLIST

'lblDOC.Visible = False

CURDOC = Left(lblDOC.Caption, 6)
CURSMP = txtWSMP.Text
CURTOT = txtWTOT.Text
CURDAY = txtDAY.Text
CURNOU = txtNOU.Text
CURDUR = txtDUR.Text
CURREC = txtREC.Text

lblDAY2.Visible = False
lblNOU2.Visible = False
lblDUR2.Visible = False
lblSMP2.Visible = False
lblTOT2.Visible = False

txtWSMP.Text = ""
txtWTOT.Text = ""

cmdSTAY.MousePointer = 1

End Sub
Private Sub DBGGEN_AfterColEdit(ByVal ColIndex As Integer)

On Error GoTo EXIT_SUB

Dim XXX

If dbgGEN.Columns(4).Value <> 0 And dbgGEN.Columns(2).Value <> 0 Then
   dbgGEN.Columns(5).Value = dbgGEN.Columns(4).Value * dbgGEN.Columns(2).Value
   End If
  
If dbgGEN.Columns(2).Value <> 0 And dbgGEN.Columns(5).Value <> 0 Then

   If dbgGEN.Columns(4).Value = 0 Then
      dbgGEN.Columns(4).Value = dbgGEN.Columns(5).Value / dbgGEN.Columns(2).Value
      End If

   End If

XXX = dbgGEN.Columns(4).Value
XXX = Int(XXX * 1000) / 1000
dbgGEN.Columns(4).Value = XXX

If dbgGEN.Columns(2).Value = 0 Then
   dbgGEN.Columns(3).Value = 0
   dbgGEN.Columns(4).Value = 0
   dbgGEN.Columns(5).Value = 0
   End If

If dbgGEN.Columns(2).Value > 99999999 Then dbgGEN.Columns(2).Value = 0
If dbgGEN.Columns(3).Value > 99999999 Then dbgGEN.Columns(3).Value = 0
If dbgGEN.Columns(4).Value > 999999 Then dbgGEN.Columns(4).Value = 0
If dbgGEN.Columns(5).Value > 999999999 Then
   dbgGEN.Columns(2).Value = 0
   dbgGEN.Columns(3).Value = 0
   dbgGEN.Columns(4).Value = 0
   dbgGEN.Columns(5).Value = 0
   
   Dim resp
   
   resp = MsgBox(msgtab(244), vbCritical + vbOKOnly, " ")
   
   End If

Dim I

NEDIT = NEDIT + 1: I = NEDIT

ReDim Preserve LTC(1 To NEDIT), LTL(1 To NEDIT), LTF(1 To NEDIT), LTP(1 To NEDIT), LTV(1 To NEDIT)

LTC(I) = dbgGEN.Columns(0).Value
LTL(I) = dbgGEN.Columns(2).Value
LTF(I) = dbgGEN.Columns(3).Value
LTP(I) = dbgGEN.Columns(4).Value
LTV(I) = dbgGEN.Columns(5).Value

Exit Sub

EXIT_SUB:

dbgGEN.Columns(2).Value = 0
dbgGEN.Columns(3).Value = 0
dbgGEN.Columns(4).Value = 0
dbgGEN.Columns(5).Value = 0

End Sub
Private Sub dbgGEN_Change()

If dbgGEN.Col = 4 Then
   dbgGEN.Columns(5).Value = 0
   Exit Sub
   End If

If dbgGEN.Col = 5 Then
   dbgGEN.Columns(4).Value = 0
   Exit Sub
   End If

End Sub
Private Sub dbgGEN_Click()

FIND_STRING = ""

End Sub
Private Sub DBGGEN_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn And dbgGEN.Col = 1 Then
   dbgGEN.Col = 1
   FIND_STRING = ""
   Exit Sub
   End If

If dbgGEN.Col <> 1 Then
   FIND_STRING = ""
   Exit Sub
   End If
 
WKEY_CODE = KeyCode

If KeyCode = vbKeyPageUp Or _
   KeyCode = vbKeyPageDown Or _
   KeyCode = vbKeyEnd Or _
   KeyCode = vbKeyHome Or _
   KeyCode = vbKeyLeft Or _
   KeyCode = vbKeyUp Or _
   KeyCode = vbKeyRight Or _
   KeyCode = vbKeyDown Then
   
   FIND_STRING = ""
   
   Exit Sub
   
   End If

Call FIND_CHAR_STR

End Sub
Private Sub FIND_CHAR_STR()

Dim CURX, KEYX, I, XXX, yyy, FIND_FLAG, POS_CUR, POS_KEY, DIFF_POS, BMKC, L, WM

KEYX = WKEY_CODE

FIND_STRING = FIND_STRING + UCase(Chr(KEYX))

If Len(FIND_STRING) >= 1 Then GoSub FIND_CHAR_KEY

If POS_KEY = 0 Then
   FIND_STRING = ""
   dbgGEN.Col = 2
   Exit Sub
   End If
   
POS_KEY = POS_KEY - 1

dtaGEN.Recordset.MoveFirst

For I = 1 To POS_KEY

dtaGEN.Recordset.MoveNext

Next I

Exit Sub

FIND_CHAR_KEY:

POS_KEY = 0

FIND_STRING = UCase(FIND_STRING)
FIND_STRING = LTrim(RTrim(FIND_STRING))
L = Len(FIND_STRING)

Dim FCHR

FCHR = Left(FIND_STRING, 1)

For I = 1 To NSP
XXX = UCase(SPEN(I))

If FCHR <> "!" Then
   If Left(XXX, L) = FIND_STRING Then
      POS_KEY = I
      Return
      End If
GoTo NEXT_I2
End If

If FCHR = "!" Then
   WM = InStr(XXX, Right(FIND_STRING, L - 1))
   If WM <> 0 Then
   POS_KEY = I
   Return
   End If
GoTo NEXT_I2
End If

NEXT_I2:

Next I

FIND_STRING = ""

Return

End Sub
Private Sub DBGGEN_KeyUp(KeyCode As Integer, Shift As Integer)

If IsNumeric(dbgGEN.Columns(2).Value) = False Then
   dbgGEN.Columns(2).Value = 0
   End If

If IsNumeric(dbgGEN.Columns(3).Value) = False Then
   dbgGEN.Columns(3).Value = 0
   End If

If IsNumeric(dbgGEN.Columns(4).Value) = False Then
   dbgGEN.Columns(4).Value = 0
   End If

If IsNumeric(dbgGEN.Columns(5).Value) = False Then
   dbgGEN.Columns(5).Value = 0
   End If

End Sub
Private Sub Form_Load()

Set Picture = LoadPicture(APPROOT + "\ARTBAS\PICS_RUNTIME\SCREEN_09.JPG")

FIND_STRING = ""

txtDAY.Text = 1
txtNOU.Text = 1
txtDUR.Text = 1
txtWSMP.Text = 0
txtWTOT.Text = 0
txtREC = "???"

lblDAY2.Visible = False
lblNOU2.Visible = False
lblDUR2.Visible = False
lblSMP2.Visible = False
lblTOT2.Visible = False

lblSUM.Caption = msgtab(118)

DECF = "000000000000.000000"

Call LOAD_SPECIES
Call CREATE_SPEDB

lblDOC.Visible = False

Call ASSO_SIMN

ReDim OKMN(1 To 10000)

OLDREC = ""

Dim cal1, cal2, I, fnm, XXX

For I = 1 To 10000
OKMN(I) = "-"
Next I

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_UNITS.TXT"

' I don't understand why this code is here since if we are in the Landings input form,
' the UNITS file for the month should have already been created.
' To avoid problems, I leave the code in for now - TNJ
If Dir(fnm) = "" Then FileCopy APPROOT + "\ARTBAS\STRUS\DEFAULT_UNITS.TXT", fnm

Open fnm For Input As #1

Line Input #1, XXX: UNW = Left(XXX, 8)
Line Input #1, XXX: UNM = Left(XXX, 8)

Close #1

fnm = APPROOT + "\ARTBAS\LANDINGS\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_LSAMPLES.TXT"

If Dir(fnm) <> "" Then
   Call SETUP_LAND
   Call SETUP_LAND2
   End If
   
frmLAND.Caption = monthtab(current_month) + " " + _
                    Format(current_year, "0000") + " - " + msgtab(87)

dtaGEN.Visible = False
dbgGEN.Visible = False
lstSIBG.Visible = False
lblSELSIBG.Visible = False
Frame1.Visible = False

cmdDEL.ToolTipText = msgtab(112)
cmdADD.ToolTipText = msgtab(92)
cmdPRINT.Visible = False
cmdSTAY.Visible = False
cmdLEAVE.Visible = False
cmdDEL.Visible = False

cmdBACK.ToolTipText = msgtab(49)
cmdPRINT.ToolTipText = msgtab(52)
cmdCONFIRM.ToolTipText = msgtab(36)

cmdPRINT.ToolTipText = msgtab(78)
cmdSTAY.ToolTipText = msgtab(36)
cmdLEAVE.ToolTipText = msgtab(100)
cmdGUIDE1.ToolTipText = msgtab(243)
cmdGUIDE2.ToolTipText = msgtab(243)

lblSELMN.Caption = msgtab(72)
lblSELSIBG.Caption = msgtab(73)
lblREC.Caption = msgtab(82)
lblNOU.Caption = msgtab(88)
lblDUR.Caption = msgtab(89)
lblSMP.Caption = msgtab(90)
lblTOT.Caption = msgtab(91)
lblDAY.Caption = msgtab(98)

dbgGEN.Columns(1).Caption = msgtab(47)
dbgGEN.Columns(2).Caption = msgtab(97)
dbgGEN.Columns(3).Caption = msgtab(93)
dbgGEN.Columns(4).Caption = msgtab(94)
dbgGEN.Columns(5).Caption = msgtab(95)

Call LOAD_STRATA

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
Private Sub LOAD_ASSO()

Dim I, J, K, XXX, yyy, fnm, L

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

fnm = APPROOT + "\ARTBAS\LANDINGS\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_LSAMPLES.TXT"

If Dir(fnm) = "" Then GoTo NO_ACTION

Open fnm For Input As #1

Do Until EOF(1)

Line Input #1, XXX

XXX = Mid(XXX, 90, 11)

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
Private Sub lstSIBG_Click()

lstSIBG.Visible = False

lblDAY2.Visible = False
lblNOU2.Visible = False
lblDUR2.Visible = False
lblSMP2.Visible = False
lblTOT2.Visible = False

cmdDEL.Visible = False

Call DELETE_ESTIM

NLST = 0

txtWSMP.Text = ""
txtWTOT.Text = ""

cmdADD.Visible = True
dbgGEN.Visible = False
cmdPRINT.Visible = False

Dim I, fnm, dbn

I = lstSIBG.ListIndex

CURSIBG = FRN(I + 1): CURSIBGC = FRC(I + 1)

lblSELSIBG.Caption = RTrim(msgtab(73)) + " : " + _
                   RTrim(Left(CURSIBG, 30)) + " + " + _
                   LTrim(Right(CURSIBG, 30))

lstSIBG.Enabled = False
Frame1.Visible = True

cmdBACK.Visible = False
cmdLEAVE.Visible = True

fnm = APPROOT + "\ARTBAS\LANDINGS\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_LSAMPLES.TXT"

If Dir(fnm) <> "" Then GoTo CR_FROM_FILE

Exit Sub

CR_FROM_FILE:

dbn = APPROOT + "\ARTBAS\LANDINGS\WORK.MDB"

Dim prm_database As Database, prm_record As Recordset, XKEY

Set prm_database = OpenDatabase(dbn)
Set prm_record = prm_database.OpenRecordset("LTAB")

With prm_record

.Index = "primarykey"

.MoveFirst

Do Until .EOF

If ![lsbc] <> CURSIBGC Then GoTo CONT_READ
If ![LDAY] = 0 Then GoTo CONT_READ

NLST = NLST + 1

ReDim Preserve LSTAB(1 To NLST)

Dim WWW, ZZZ, ddd

ddd = Format(![LDOC], "000000")

WWW = ![LSMP]: WWW = Format(WWW, "#####0.00")
WWW = Right(Space(10) + LTrim(WWW), 9)

ZZZ = WWW

WWW = ![ltot]: WWW = Format(WWW, "#####0.00")
WWW = Right(Space(10) + LTrim(WWW), 9)

LSTAB(NLST) = ddd + " " + _
              Format(![LDAY], "00") + _
              Right(Space(10) + LTrim(Format(![LNOU], "##0.00")), 7) + " " + _
              Right(Space(10) + LTrim(Format(![LDUR], "##0.00")), 7) + " " + _
              ZZZ + " " + WWW

CONT_READ:

.MoveNext

Loop

prm_record.Close
prm_database.Close

End With

Call CREATE_SUMLIST

End Sub
Private Sub CREATE_SUMLIST()

Dim I

lstSUM.Clear

For I = 1 To NLST

lstSUM.AddItem LSTAB(I)

lstSUM.ItemData(lstSUM.NewIndex) = Val(Left(LSTAB(I), 6))

If lstSUM.ListCount > 5 Then lstSUM.TopIndex = lstSUM.ListCount - 5

Next I

End Sub
Private Sub lstSUM_Click()

NEDIT = 0

cmdLEAVE.Visible = False
cmdSTAY.Visible = True
dbgGEN.Visible = True
cmdADD.Visible = False
cmdDEL.Visible = True
lstSUM.Enabled = False
cmdPRINT.Visible = False

Dim ddd, dbn

ddd = lstSUM.ItemData(lstSUM.ListIndex)

CURSEL = lstSUM.ListIndex + 1

lblDOC.Caption = Format(ddd, "000000")
lblDOC.Visible = True

dbn = APPROOT + "\ARTBAS\LANDINGS\WORK.MDB"

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(dbn)
Set prm_record = prm_database.OpenRecordset("LTAB")

With prm_record

.Index = "primarykey"

.Seek "=", ddd

If .NoMatch = True Then Exit Sub

txtDAY.Text = ![LDAY]
txtNOU.Text = ![LNOU]
txtDUR.Text = ![LDUR]
txtWSMP.Text = ![LSMP]
txtWTOT.Text = ![ltot]
txtREC.Text = ![LREC]

lblDAY2.Visible = True
lblNOU2.Visible = True
lblDUR2.Visible = True
lblSMP2.Visible = True
lblTOT2.Visible = True

lblDAY2.Caption = ![LDAY]
lblNOU2.Caption = ![LNOU]
lblDUR2.Caption = ![LDUR]
lblSMP2.Caption = ![LSMP]
lblTOT2.Caption = ![ltot]

End With

prm_record.Close
prm_database.Close

dtaGEN.DatabaseName = APPROOT + "\ARTBAS\STRUS\DSPECIES.MDB"
dtaGEN.Refresh

Call READY_SPECIES

dtaGEN.DatabaseName = APPROOT + "\ARTBAS\LANDINGS\WORKSP.MDB"
dtaGEN.Refresh

End Sub
Private Sub txtDAY_Change()
If IsNumeric(txtDAY.Text) = False Then txtDAY.Text = ""
If IsNumeric(txtDAY.Text) = True And Val(txtDAY.Text) > CURCAL Then txtDAY.Text = ""
If IsNumeric(txtDAY.Text) = True And Val(txtDAY.Text) < 1 Then txtDAY.Text = ""
End Sub
Private Sub txtDUR_Change()
If IsNumeric(txtDUR.Text) = False Then txtDUR.Text = ""
End Sub
Private Sub txtNOU_Change()
If IsNumeric(txtNOU.Text) = False Then txtNOU.Text = ""
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
Private Sub DUMP_WORK()

Dim fnm, dbn

dbn = APPROOT + "\ARTBAS\LANDINGS\WORK.MDB"

If Dir(dbn) = "" Then
   Exit Sub
   End If
   
fnm = APPROOT + "\ARTBAS\LANDINGS\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_LSAMPLES.TXT"

Open fnm For Output As #1
   
Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(dbn)
Set prm_record = prm_database.OpenRecordset("LTAB")

With prm_record

.MoveFirst

Do Until .EOF

If ![LDAY] = 0 Then GoTo CONT_READ

Print #1, Format(![LDOC], "000000") + " " + _
          Format(![LDAY], "00") + " " + _
          Format(![LNOU], "0000.000") + " " + _
          Format(![LDUR], "0000.000") + " " + _
          Format(![LSMP], DECF) + " " + _
          Format(![ltot], DECF) + " " + _
          Left(![LREC] + Space(15), 15) + " " + _
          Format(![LMNC], "0000") + " " + _
          ![lsbc]

CONT_READ:

.MoveNext

Loop

Close #1

End With

prm_record.Close
prm_database.Close

End Sub
Private Sub SETUP_LAND()

Dim fnm, dbn, XXX, xcode, xmnc, xact, xsmp, xfrm, xrec

fnm = APPROOT + "\ARTBAS\LANDINGS\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_LSAMPLES.TXT"

Open fnm For Input As #1

FileCopy APPROOT + "\ARTBAS\STRUS\LSAMPLES.MDB", APPROOT + "\ARTBAS\LANDINGS\WORK.MDB"

dbn = APPROOT + "\ARTBAS\LANDINGS\WORK.MDB"

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(dbn)
Set prm_record = prm_database.OpenRecordset("LTAB")

With prm_record

.Index = "primarykey"

Do Until EOF(1)

Line Input #1, XXX

.AddNew

![LDOC] = CDbl(Left(XXX, 6))
![LDAY] = CDbl(Mid(XXX, 8, 2))
![LNOU] = CDbl(Mid(XXX, 11, 8))
![LDUR] = CDbl(Mid(XXX, 20, 8))
![LSMP] = CDbl(Mid(XXX, 29, 19))
![ltot] = CDbl(Mid(XXX, 49, 19))
![LREC] = Mid(XXX, 69, 15)
![LMNC] = CDbl(Mid(XXX, 85, 4))
![lsbc] = Mid(XXX, 90, 11)

OKMN(![LMNC]) = "+"

.Update

Loop

Close #1

End With

prm_record.Close
prm_database.Close

End Sub
Private Sub ASSO_SIMN()

Dim I, J, K, XXX, yyy, fnm

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

For I = 1 To K
Line Input #1, yyy
I = Val(Mid(yyy, 6, 4)): SIMN(I) = J
Next I
  
Loop
  
Close #1

End Sub
Private Sub txtWSMP_Change()

If IsNumeric(txtWSMP.Text) = False Then txtWSMP.Text = ""

End Sub
Private Sub LOAD_SPECIES()

Dim fnm, I, J, XXX

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_SPECIES.TXT"
      
Open fnm For Input As #1
NSP = 0

Do Until EOF(1)

Line Input #1, XXX

NSP = NSP + 1

ReDim Preserve SPEC(1 To NSP), SPEN(1 To NSP), SPSEQ(1 To NSP)

SPEC(NSP) = Left(XXX, 4): SPEN(NSP) = Mid(XXX, 6, 30)
SPSEQ(NSP) = Mid(XXX, 69, 6)

Loop

Close #1

End Sub
Private Sub CREATE_SPEDB()

Dim dbn, I

dtaGEN.DatabaseName = APPROOT + "\ARTBAS\STRUS\SPEAUX.MDB"
dtaGEN.Refresh

FileCopy APPROOT + "\ARTBAS\STRUS\DSPECIES.MDB", APPROOT + "\ARTBAS\LANDINGS\WORKSP.MDB"

'dtaGEN.DatabaseName = APPROOT + "\ARTBAS\LANDINGS\WORKSP.MDB"
'dtaGEN.Refresh

dbn = APPROOT + "\ARTBAS\LANDINGS\WORKSP.MDB"

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(dbn)
Set prm_record = prm_database.OpenRecordset("STAB")

With prm_record

For I = 1 To NSP

.AddNew

![SCODE] = SPEC(I)
![sdes] = SPEN(I)
![ssort] = SPSEQ(I)
![slan] = 0
![snof] = 0
![spri] = 0
![sval] = 0

.Update

Next I

End With

prm_record.Close
prm_database.Close

dtaGEN.DatabaseName = APPROOT + "\ARTBAS\LANDINGS\WORKSP.MDB"
dtaGEN.Refresh

End Sub
Private Sub COPY_SPECIES()

Dim I, J, dbn, XKEY

dbn = APPROOT + "\ARTBAS\LANDINGS\WORKSP.MDB"

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(dbn)
Set prm_record = prm_database.OpenRecordset("STAB")

With prm_record

.Index = "primarykey"

For I = 1 To NEDIT

XKEY = LTC(I)

.Seek "=", XKEY

If .NoMatch = True Then End

.Edit

![slan] = LTL(I)
![snof] = LTF(I)
![spri] = LTP(I)
![sval] = LTV(I)

If ![slan] <> 0 Then
   If ![spri] <> 0 And ![sval] = 0 Then ![sval] = ![slan] * ![spri]
   If ![spri] = 0 And ![sval] <> 0 Then ![spri] = ![sval] / ![slan]
   End If

.Update

Next I

NEDIT = 0

.MoveFirst

ReDim LTC(1 To NSP), LTL(1 To NSP), LTL(1 To NSP), LTF(1 To NSP), LTP(1 To NSP), LTV(1 To NSP)

I = 0

Do Until .EOF

I = I + 1

LTC(I) = ![SCODE]
LTL(I) = ![slan]
LTF(I) = ![snof]
LTP(I) = ![spri]
LTV(I) = ![sval]

.MoveNext

Loop

End With

prm_record.Close
prm_database.Close

dbn = APPROOT + "\ARTBAS\LANDINGS\WORK2.MDB"

Set prm_database = OpenDatabase(dbn)
Set prm_record = prm_database.OpenRecordset("STAB")

With prm_record

.Index = "primarykey"

For I = 1 To NSP

XKEY = "D" + Left(lblDOC.Caption, 6) + "+S" + LTC(I)

.Seek "=", XKEY

If .NoMatch = True Then

   .AddNew
   
   ![skey] = XKEY
   ![slan] = LTL(I)
   ![snof] = LTF(I)
   ![spri] = LTP(I)
   ![sval] = LTV(I)
   ![smnc] = CURMNC
   ![ssbc] = CURSIBGC
   
   .Update
   
   GoTo CONT_LOOP
   
   End If

.Edit

![slan] = LTL(I)
![snof] = LTF(I)
![spri] = LTP(I)
![sval] = LTV(I)
![smnc] = CURMNC
![ssbc] = CURSIBGC
   
.Update

CONT_LOOP:

Next I

End With

prm_record.Close
prm_database.Close

End Sub
Private Sub DUMP_WORK2()

Dim fnm, dbn

dbn = APPROOT + "\ARTBAS\LANDINGS\WORK2.MDB"

If Dir(dbn) = "" Then Exit Sub

fnm = APPROOT + "\ARTBAS\LANDINGS\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_LSPECIES.TXT"

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(dbn)
Set prm_record = prm_database.OpenRecordset("STAB")

With prm_record

If .RecordCount = 0 Then Exit Sub

Open fnm For Output As #1

.MoveFirst

Do Until .EOF

If ![slan] <> 0 Then
    Print #1, ![skey] + " " + _
          Format(![slan], DECF) + " " + _
          Format(![snof], DECF) + " " + _
          Format(![spri], DECF) + " " + _
          Format(![sval], DECF) + " " + _
          Format(![smnc], "0000") + " " + _
          ![ssbc]
    
    End If
    
.MoveNext

Loop

Close #1

End With

prm_record.Close
prm_database.Close

End Sub
Private Sub SETUP_LAND2()

Dim fnm, dbn, XXX, xcode, xmnc, xact, xsmp, xfrm, xrec

fnm = APPROOT + "\ARTBAS\LANDINGS\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_LSPECIES.TXT"

If Dir(fnm) = "" Then
   FileCopy APPROOT + "\ARTBAS\STRUS\LSPECIES.MDB", APPROOT + "\ARTBAS\LANDINGS\WORK2.MDB"
   Exit Sub
   End If
   
Open fnm For Input As #1

FileCopy APPROOT + "\ARTBAS\STRUS\LSPECIES.MDB", APPROOT + "\ARTBAS\LANDINGS\WORK2.MDB"

dbn = APPROOT + "\ARTBAS\LANDINGS\WORK2.MDB"

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(dbn)
Set prm_record = prm_database.OpenRecordset("STAB")

With prm_record

.Index = "primarykey"

Do Until EOF(1)

Line Input #1, XXX

.AddNew

![skey] = Left(XXX, 13)
![slan] = CDbl(Mid(XXX, 15, 19))
![snof] = CDbl(Mid(XXX, 35, 19))
![spri] = CDbl(Mid(XXX, 55, 19))
![sval] = CDbl(Mid(XXX, 75, 19))
![smnc] = Val(Mid(XXX, 95, 4))
![ssbc] = Mid(XXX, 100, 11)

.Update

Loop

Close #1

End With

prm_record.Close
prm_database.Close
End Sub
Private Sub READY_SPECIES()

Dim I, J, dbn, XKEY, fnm, ddd

ddd = "D" + Left(lblDOC.Caption, 6)

Dim SC(), SPL(), SPF(), SPP(), SPV()

ReDim SC(1 To NSP), SPL(1 To NSP), SPP(1 To NSP), SPV(1 To NSP), SPF(1 To NSP)

For I = 1 To NSP

SC(I) = SPEC(I)
SPL(I) = 0: SPF(I) = 0: SPP(I) = 0: SPV(I) = 0

Next I

dbn = APPROOT + "\ARTBAS\LANDINGS\WORKSP.MDB"

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(dbn)
Set prm_record = prm_database.OpenRecordset("STAB")

With prm_record

.MoveFirst

Do Until .EOF

.Edit
![slan] = 0: ![snof] = 0: ![spri] = 0: ![sval] = 0
.Update

.MoveNext

Loop

End With

prm_record.Close
prm_database.Close

dbn = APPROOT + "\ARTBAS\LANDINGS\WORK2.MDB"

Set prm_database = OpenDatabase(dbn)
Set prm_record = prm_database.OpenRecordset("STAB")

With prm_record

.Index = "primarykey"

For I = 1 To NSP

XKEY = ddd + "+S" + SC(I)

.Seek "=", XKEY

If .NoMatch = True Then
   GoTo CONT_LOOP
   End If

SPL(I) = ![slan]
SPF(I) = ![snof]
SPP(I) = ![spri]
SPV(I) = ![sval]

CONT_LOOP:

Next I

End With

prm_record.Close
prm_database.Close

dbn = APPROOT + "\ARTBAS\LANDINGS\WORKSP.MDB"

Set prm_database = OpenDatabase(dbn)
Set prm_record = prm_database.OpenRecordset("STAB")

With prm_record

.Index = "primarykey"

For I = 1 To NSP

XKEY = SC(I)

.Seek "=", XKEY

If .NoMatch = True Then
    GoTo CONT_LOOP2
    End If
    
   .Edit
   
   ![slan] = SPL(I)
   ![snof] = SPF(I)
   ![spri] = SPP(I)
   ![sval] = SPV(I)
   
   .Update
   
CONT_LOOP2:

Next I

End With

prm_record.Close
prm_database.Close

End Sub
Private Sub CONTROL_TOTALS()

Dim dbn, I

ERRTOT = 0

SUMERR = " "

If txtWSMP.Text = 0.001 Then Exit Sub

dbn = APPROOT + "\ARTBAS\LANDINGS\WORKSP.MDB"

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(dbn)
Set prm_record = prm_database.OpenRecordset("STAB")

With prm_record

.Index = "primarykey"

.MoveFirst

Do Until .EOF

ERRTOT = ERRTOT + ![slan]

.MoveNext

Loop

End With

prm_record.Close
prm_database.Close

If Abs(ERRTOT - CDbl(txtWSMP.Text)) > 0.5 Then
   SUMERR = "Y"
   txtWTOT.Text = 0
   End If
   
If Abs(ERRTOT - CDbl(txtWSMP.Text)) <= 0.5 Then
   If txtWSMP.Text = txtWTOT.Text Then
      txtWSMP.Text = ERRTOT: txtWTOT.Text = ERRTOT
      End If
   End If

dbn = APPROOT + "\ARTBAS\LANDINGS\WORK.MDB"

Set prm_database = OpenDatabase(dbn)
Set prm_record = prm_database.OpenRecordset("LTAB")

Dim ddd

With prm_record

.Index = "primarykey"

ddd = CDbl(lblDOC.Caption)

.Seek "=", ddd

If .NoMatch = True Then Exit Sub

.Edit

![LSMP] = CDbl(txtWSMP.Text)

If Len(txtWTOT.Text) = 0 Then ![ltot] = ![LSMP]

If Len(txtWTOT.Text) <> 0 Then ![ltot] = CDbl(txtWTOT.Text)

If ![ltot] < ![LSMP] Then ![ltot] = ![LSMP]

.Update

End With

prm_record.Close
prm_database.Close

End Sub
Private Sub DELETE_ESTIM()

Dim fnm

fnm = APPROOT + "\ARTBAS\RESULTS\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "*.*"

If Dir(fnm) <> "" Then Kill fnm

End Sub
Private Sub txtWTOT_Change()

If IsNumeric(txtWTOT.Text) = False Then txtWTOT.Text = ""


End Sub
