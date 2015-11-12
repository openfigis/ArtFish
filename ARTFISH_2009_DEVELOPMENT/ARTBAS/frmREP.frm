VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmREP 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ARTPLAN 1"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11520
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
   MousePointer    =   3  'I-Beam
   Moveable        =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   11520
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
      Left            =   5760
      TabIndex        =   51
      Top             =   6120
      Width           =   5535
   End
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
      Left            =   240
      TabIndex        =   50
      Top             =   6120
      Width           =   5535
   End
   Begin VB.CommandButton cmdEXCEL_UNITS 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6840
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmREP.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   6960
      Width           =   375
   End
   Begin VB.CommandButton cmdEXCEL_ESTIM 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8280
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmREP.frx":1326
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   6960
      Width           =   375
   End
   Begin VB.CommandButton cmdEXCEL_RAW 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7560
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmREP.frx":264C
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   6960
      Width           =   375
   End
   Begin VB.CommandButton cmdEXCEL_ACTIVE 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmREP.frx":3972
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   6960
      Width           =   375
   End
   Begin VB.CommandButton cmdEXCEL_SPECIES 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmREP.frx":4C98
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   6960
      Width           =   375
   End
   Begin VB.CommandButton cmdEXCEL_FRAME 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmREP.frx":5FBE
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   6960
      Width           =   375
   End
   Begin VB.CommandButton cmdEXCEL_BG 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmREP.frx":72E4
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   6960
      Width           =   375
   End
   Begin VB.CommandButton cmdEXCEL_SITMIN 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmREP.frx":860A
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   6960
      Width           =   375
   End
   Begin VB.CommandButton cmdEXCEL_SITES 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmREP.frx":9930
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   6960
      Width           =   375
   End
   Begin VB.CommandButton cmdEXCEL_MINMAJ 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmREP.frx":AC56
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   6960
      Width           =   375
   End
   Begin VB.CommandButton cmdEXCEL_MINOR 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmREP.frx":BF7C
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   6960
      Width           =   375
   End
   Begin VB.CommandButton cmdEXCEL_MAJOR 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmREP.frx":D2A2
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   6960
      Width           =   375
   End
   Begin VB.CommandButton cmdDOCS 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   7440
      Picture         =   "frmREP.frx":E5C8
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   6360
      Width           =   615
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
      Left            =   9720
      MousePointer    =   1  'Arrow
      Picture         =   "frmREP.frx":E84A
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   6360
      Width           =   735
   End
   Begin VB.ListBox lstLIST 
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
      ForeColor       =   &H00800000&
      Height          =   5100
      Left            =   120
      MultiSelect     =   2  'Extended
      TabIndex        =   32
      Top             =   360
      Width           =   7455
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5295
      Left            =   7680
      TabIndex        =   15
      Top             =   360
      Width           =   3735
      Begin VB.OptionButton optLOG 
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
         TabIndex        =   31
         Top             =   1560
         Width           =   3375
      End
      Begin VB.OptionButton optGT4 
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
         TabIndex        =   30
         Top             =   4440
         Width           =   3375
      End
      Begin VB.OptionButton optGT3 
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
         TabIndex        =   29
         Top             =   4200
         Width           =   3375
      End
      Begin VB.OptionButton optGT2 
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
         TabIndex        =   28
         Top             =   3960
         Width           =   3375
      End
      Begin VB.OptionButton optGT1 
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
         TabIndex        =   27
         Top             =   3720
         Width           =   3375
      End
      Begin VB.OptionButton optMJ4 
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
         TabIndex        =   25
         Top             =   3000
         Width           =   3375
      End
      Begin VB.OptionButton optMJ3 
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
         TabIndex        =   24
         Top             =   2760
         Width           =   3375
      End
      Begin VB.OptionButton optMJ2 
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
         TabIndex        =   23
         Top             =   2520
         Width           =   3375
      End
      Begin VB.OptionButton optMJ1 
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
         TabIndex        =   22
         Top             =   2280
         Width           =   3375
      End
      Begin VB.OptionButton optMN4 
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
         Top             =   1200
         Width           =   3375
      End
      Begin VB.OptionButton optMN3 
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
         Top             =   960
         Width           =   3375
      End
      Begin VB.OptionButton optMN2 
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
         Top             =   720
         Width           =   3375
      End
      Begin VB.OptionButton optMN1 
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
         TabIndex        =   17
         Top             =   480
         Width           =   3375
      End
      Begin VB.Label lblOPTIONS 
         Alignment       =   2  'Center
         BackColor       =   &H00808000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   36
         Top             =   4800
         Width           =   3375
      End
      Begin VB.Label lblGT 
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
         TabIndex        =   26
         Top             =   3360
         Width           =   3375
      End
      Begin VB.Label lblMJ 
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
         TabIndex        =   21
         Top             =   1920
         Width           =   3375
      End
      Begin VB.Label lblMN 
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
         TabIndex        =   16
         Top             =   120
         Width           =   3375
      End
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
      Left            =   8880
      MousePointer    =   1  'Arrow
      Picture         =   "frmREP.frx":10AAC
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6360
      Width           =   735
   End
   Begin VB.CommandButton cmdQUIT 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   10920
      MousePointer    =   1  'Arrow
      Picture         =   "frmREP.frx":10D2E
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6720
      Width           =   375
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
      Height          =   615
      Left            =   6000
      Picture         =   "frmREP.frx":10FB0
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6360
      Width           =   615
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
      Height          =   615
      Left            =   8160
      MousePointer    =   1  'Arrow
      Picture         =   "frmREP.frx":11232
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6360
      Width           =   615
   End
   Begin VB.CommandButton cmdRETURN 
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
      Left            =   10560
      MousePointer    =   1  'Arrow
      Picture         =   "frmREP.frx":114B4
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6360
      Width           =   735
   End
   Begin VB.CommandButton cmdMAJOR 
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
      Height          =   615
      Left            =   240
      Picture         =   "frmREP.frx":11736
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6360
      Width           =   615
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
      Height          =   585
      Left            =   3120
      Picture         =   "frmREP.frx":13658
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6390
      Width           =   615
   End
   Begin VB.CommandButton cmdSITES 
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
      Height          =   615
      Left            =   2400
      Picture         =   "frmREP.frx":138DA
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6360
      Width           =   615
   End
   Begin VB.CommandButton cmdSPECIES 
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
      Height          =   615
      Left            =   5280
      Picture         =   "frmREP.frx":13B5C
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6360
      Width           =   615
   End
   Begin VB.CommandButton cmdFRAME 
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
      Height          =   615
      Left            =   4560
      Picture         =   "frmREP.frx":1465E
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6360
      Width           =   615
   End
   Begin VB.CommandButton cmdUNITS 
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
      Height          =   615
      Left            =   6720
      Picture         =   "frmREP.frx":147C4
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6360
      Width           =   615
   End
   Begin VB.CommandButton cmdBG 
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
      Height          =   615
      Left            =   3840
      Picture         =   "frmREP.frx":14A46
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6360
      Width           =   615
   End
   Begin VB.CommandButton cmdASSOM 
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
      Height          =   615
      Left            =   1680
      Picture         =   "frmREP.frx":14CC8
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6360
      Width           =   615
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
      Height          =   615
      Left            =   960
      MousePointer    =   1  'Arrow
      Picture         =   "frmREP.frx":14F4A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6360
      Width           =   615
   End
   Begin RichTextLib.RichTextBox rtsDISP 
      Height          =   5775
      Left            =   120
      TabIndex        =   14
      Top             =   360
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   10186
      _Version        =   393217
      BackColor       =   12648447
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      MousePointer    =   1
      TextRTF         =   $"frmREP.frx":173A4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FFFF&
      Caption         =   " 11"
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
      TabIndex        =   37
      Top             =   7080
      Width           =   255
   End
   Begin VB.Label lblEXP 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   33
      Top             =   5880
      Width           =   8535
   End
End
Attribute VB_Name = "frmREP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SITEN(), MAJORN(), MINORN(), BGN(), SPEN()
Private NOMN, NOMJ, NOSI, NOBG, NOSP, NOMNBG, TMNBG(), TMNBGF(), KSP
Private NOMBS, TMBS()
Private TMJC(), TMJN(), TMJOK(), NTMJ
Private TMNC(), TMNN(), TMNOK(), NTMN
Private TSIC(), TSIN(), NTSI
Private TBGC(), TBGN(), NTBG
Private TSPC(), TSPN(), NTSP
Private ASSOMN(), ASSOSI(), NASSOMN, NASSOSI, ESTFLAG
Private ESTDISP
Private NOEXCL, EXN()
Private EXPFNM
Private ZF, ZC, ZE, ZU, ZP, ZV, ZW
Private DOCMNBG, DOCSP, SELMBC, SELSPC, SELMBN, SELSPN
Private LOG_OPTION, DOCTMNBG(1 To 200, 1 To 200), DOCS_FLAG, GRAND_TOTAL_SP
Dim FINAL_CPUE, FINAL_BAC, FINAL_ACT, FINAL_FRAME, FINAL_CATCH, FINAL_EFFORT, SAMPLE_TOT

Private Sub cmdACT_Click()

ESTDISP = "N"
Frame1.Visible = False
lstLIST.Visible = False
rtsDISP.Visible = True
cmdPRINT.Visible = True

Open APPROOT + "\ARTBAS\RESULTS\WORK.TXT" For Output As #2
Open APPROOT + "\ARTBAS\CURRENT_TABLES\ACTIVE.TXT" For Output As #4

Print #2, Tab(5); frmREP.Caption
Print #2, " "
Print #2, Tab(5); msgtab(35)
Print #2, " "
Print #2, Tab(5); msgtab(54); Tab(70); msgtab(69)
Print #2, Tab(5); String(80, "=")

Write #4, frmREP.Caption, msgtab(35), "  "
Write #4, " ", " ", "  "
Write #4, msgtab(42), msgtab(139), msgtab(35)
Write #4, " ", " ", "  "

Dim fnm, XXX, I, J, V, W

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_ACTIVE.TXT"

If Dir(fnm) = "" Then Exit Sub

Open fnm For Input As #1

Do Until EOF(1)

Line Input #1, XXX

I = Val(Mid(XXX, 2, 4)): J = Val(Mid(XXX, 8, 4)): V = CDbl(Mid(XXX, 26, 15))

W = Format(V, "#####0.00")

Print #2, Tab(5); MINORN(I) + " " + BGN(J) + Right(Space(9) + LTrim(W), 9)
Write #4, LTrim(RTrim(MINORN(I))), LTrim(RTrim(BGN(J))), CDbl(W)
Loop

Close #1
Close #2
Close #4

rtsDISP.FileName = APPROOT + "\ARTBAS\RESULTS\WORK.TXT"
rtsDISP.Visible = True

End Sub
Private Sub cmdASSO_Click()

Dim KK

ESTDISP = "N"
cmdPRINT.Visible = True
Frame1.Visible = False
lstLIST.Visible = False
rtsDISP.Visible = True

Open APPROOT + "\ARTBAS\RESULTS\WORK.TXT" For Output As #2
Open APPROOT + "\ARTBAS\CURRENT_TABLES\SITES_MINOR.TXT" For Output As #4

Print #2, Tab(5); frmREP.Caption
Print #2, " "
Print #2, Tab(5); msgtab(44)
Print #2, Tab(5); String(80, "=")
Print #2, " "

Write #4, frmREP.Caption, msgtab(44)
Write #4, "  ", "  "


Dim fnm, XXX, I, J, V, W, yyy

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_ASSOSI.TXT"

If Dir(fnm) = "" Then Exit Sub

Open fnm For Input As #1

Do Until EOF(1)

Line Input #1, XXX

KK = Val(Left(XXX, 4))

J = Val(Mid(XXX, 37, 4))

Print #2, Tab(5); Left(XXX, 35)
Print #2, " "

Write #4, Format(KK, "0000") + " " + MINORN(KK), "  "

For I = 1 To J

Line Input #1, yyy

KK = Val(Mid(yyy, 6, 4))

Print #2, Tab(5); yyy

Write #4, "  ", Format(KK, "0000") + " " + SITEN(KK)

Next I

Print #2, " "

Loop

Close #1
Close #2
Close #4

rtsDISP.FileName = APPROOT + "\ARTBAS\RESULTS\WORK.TXT"
rtsDISP.Visible = True

cmdPRINT.Visible = True

End Sub
Private Sub cmdASSOM_Click()

Dim KK

ESTDISP = "N"
Frame1.Visible = False
lstLIST.Visible = False
rtsDISP.Visible = True
cmdPRINT.Visible = True

'=======================================================
Open APPROOT + "\ARTBAS\RESULTS\WORK.TXT" For Output As #2
Open APPROOT + "\ARTBAS\CURRENT_TABLES\MIN_MAJ.TXT" For Output As #4

Print #2, Tab(5); frmREP.Caption
Print #2, " "
Print #2, Tab(5); msgtab(63)
Print #2, Tab(5); String(80, "=")
Print #2, " "

Write #4, frmREP.Caption, msgtab(63)
Write #4, " ", "  "

Dim fnm, XXX, I, J, V, W, yyy

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_ASSOMN.TXT"

If Dir(fnm) = "" Then Exit Sub

Open fnm For Input As #1

Dim V1, V2, V3, V4

Do Until EOF(1)

Line Input #1, XXX

KK = Val(Left(XXX, 4))

J = Val(Mid(XXX, 37, 4))

Print #2, Tab(5); Left(XXX, 35)
Print #2, " "

Write #4, Format(KK, "0000") + " " + MAJORN(KK), "  "

For I = 1 To J

Line Input #1, yyy

KK = Val(Mid(yyy, 6, 4))

Print #2, yyy

Write #4, "  ", Format(KK, "0000") + " " + MINORN(KK)

Next I

Print #2, " "
Write #4, " ", "  "

Loop

Close #1
Close #2
Close #4


rtsDISP.FileName = APPROOT + "\ARTBAS\RESULTS\WORK.TXT"
rtsDISP.Visible = True

cmdPRINT.Visible = True


End Sub

Private Sub cmdEXCEL_ACTIVE_Click()

On Error GoTo CLOSE_EXCEL

Call cmdACT_Click

rtsDISP.Visible = False
cmdPRINT.Visible = False

FileCopy APPROOT + "\ARTBAS\BLANK_WORKSHEETS\ACTIVE.XLS", APPROOT + "\ARTBAS\EXCEL_REPORTS\ARTFISH_WORK.XLS"

RUN_FILE = APPROOT + "\ARTBAS\EXCEL_REPORTS\GENERAL.BAT"

RUN_CODE = Shell(RUN_FILE, 4)

Exit Sub

CLOSE_EXCEL:

Call EXCEL_REQUEST_ERROR

Exit Sub

End Sub

Private Sub cmdEXCEL_BG_Click()

On Error GoTo CLOSE_EXCEL

Call cmdBG_Click

rtsDISP.Visible = False
cmdPRINT.Visible = False

FileCopy APPROOT + "\ARTBAS\BLANK_WORKSHEETS\BG.XLS", APPROOT + "\ARTBAS\EXCEL_REPORTS\ARTFISH_WORK.XLS"

RUN_FILE = APPROOT + "\ARTBAS\EXCEL_REPORTS\GENERAL.BAT"

RUN_CODE = Shell(RUN_FILE, 4)

Exit Sub

CLOSE_EXCEL:

Call EXCEL_REQUEST_ERROR

End Sub

Private Sub cmdEXCEL_ESTIM_Click()

On Error GoTo CLOSE_EXCEL

If LOG_OPTION = "YES" Then
   rtsDISP.Visible = False
   cmdPRINT.Visible = False
   cmdEXCEL_ESTIM.Visible = False
   FileCopy APPROOT + "\ARTBAS\BLANK_WORKSHEETS\LOG.XLS", APPROOT + "\ARTBAS\EXCEL_REPORTS\ARTFISH_WORK.XLS"
   RUN_FILE = APPROOT + "\ARTBAS\EXCEL_REPORTS\GENERAL.BAT"
   RUN_CODE = Shell(RUN_FILE, 4)
   Frame1.Visible = True
   Exit Sub
   End If

cmdEXCEL_ESTIM.Visible = False

rtsDISP.Visible = False
cmdPRINT.Visible = False

FileCopy APPROOT + "\ARTBAS\BLANK_WORKSHEETS\ESTIM.XLS", APPROOT + "\ARTBAS\EXCEL_REPORTS\ARTFISH_WORK.XLS"

RUN_FILE = APPROOT + "\ARTBAS\EXCEL_REPORTS\GENERAL.BAT"

RUN_CODE = Shell(RUN_FILE, 4)

Frame1.Visible = True

Exit Sub

CLOSE_EXCEL:

Call EXCEL_REQUEST_ERROR

End Sub

Private Sub cmdEXCEL_FRAME_Click()

On Error GoTo CLOSE_EXCEL

Call cmdFRAME_Click

rtsDISP.Visible = False
cmdPRINT.Visible = False

FileCopy APPROOT + "\ARTBAS\BLANK_WORKSHEETS\FRAME.XLS", APPROOT + "\ARTBAS\EXCEL_REPORTS\ARTFISH_WORK.XLS"

RUN_FILE = APPROOT + "\ARTBAS\EXCEL_REPORTS\GENERAL.BAT"

RUN_CODE = Shell(RUN_FILE, 4)

Exit Sub

CLOSE_EXCEL:

Call EXCEL_REQUEST_ERROR

End Sub

Private Sub cmdEXCEL_MAJOR_Click()

Dim resp

On Error GoTo CLOSE_EXCEL

Call cmdMAJOR_Click

Dim RUN_CODE, RUN_FILE

rtsDISP.Visible = False
cmdPRINT.Visible = False

FileCopy APPROOT + "\ARTBAS\BLANK_WORKSHEETS\MAJOR.XLS", APPROOT + "\ARTBAS\EXCEL_REPORTS\ARTFISH_WORK.XLS"

RUN_FILE = APPROOT + "\ARTBAS\EXCEL_REPORTS\GENERAL.BAT"

RUN_CODE = Shell(RUN_FILE, 4)

Exit Sub

CLOSE_EXCEL:

Call EXCEL_REQUEST_ERROR

End Sub

Private Sub cmdEXCEL_MINMAJ_Click()

On Error GoTo CLOSE_EXCEL

Call cmdASSOM_Click

rtsDISP.Visible = False
cmdPRINT.Visible = False

FileCopy APPROOT + "\ARTBAS\BLANK_WORKSHEETS\MIN_MAJ.XLS", APPROOT + "\ARTBAS\EXCEL_REPORTS\ARTFISH_WORK.XLS"

RUN_FILE = APPROOT + "\ARTBAS\EXCEL_REPORTS\GENERAL.BAT"

RUN_CODE = Shell(RUN_FILE, 4)

Exit Sub

CLOSE_EXCEL:

Call EXCEL_REQUEST_ERROR

End Sub

Private Sub cmdEXCEL_MINOR_Click()

On Error GoTo CLOSE_EXCEL

Call cmdMINOR_Click

Dim RUN_CODE, RUN_FILE

rtsDISP.Visible = False
cmdPRINT.Visible = False

FileCopy APPROOT + "\ARTBAS\BLANK_WORKSHEETS\MINOR.XLS", APPROOT + "\ARTBAS\EXCEL_REPORTS\ARTFISH_WORK.XLS"

RUN_FILE = APPROOT + "\ARTBAS\EXCEL_REPORTS\GENERAL.BAT"

RUN_CODE = Shell(RUN_FILE, 4)

Exit Sub

CLOSE_EXCEL:

Call EXCEL_REQUEST_ERROR

rtsDISP.Visible = False
cmdPRINT.Visible = False

End Sub
Private Sub cmdEXCEL_RAW_Click()

cmdESTIM.Enabled = True

On Error GoTo CLOSE_EXCEL

rtsDISP.Visible = False
cmdPRINT.Visible = False

If DOCS_FLAG = "SPECIES" Then
FileCopy APPROOT + "\ARTBAS\BLANK_WORKSHEETS\RAWDATA_SPECIES.XLS", APPROOT + "\ARTBAS\EXCEL_REPORTS\ARTFISH_WORK.XLS"
End If

If DOCS_FLAG = "TOTALS" Then
FileCopy APPROOT + "\ARTBAS\BLANK_WORKSHEETS\RAWDATA.XLS", APPROOT + "\ARTBAS\EXCEL_REPORTS\ARTFISH_WORK.XLS"
End If

RUN_FILE = APPROOT + "\ARTBAS\EXCEL_REPORTS\GENERAL.BAT"

RUN_CODE = Shell(RUN_FILE, 4)

cmdEXCEL_RAW.Visible = False

Exit Sub

CLOSE_EXCEL:

Call EXCEL_REQUEST_ERROR

End Sub

Private Sub cmdEXCEL_SITES_Click()

On Error GoTo CLOSE_EXCEL

Call cmdSITES_Click

rtsDISP.Visible = False
cmdPRINT.Visible = False

FileCopy APPROOT + "\ARTBAS\BLANK_WORKSHEETS\SITES.XLS", APPROOT + "\ARTBAS\EXCEL_REPORTS\ARTFISH_WORK.XLS"

RUN_FILE = APPROOT + "\ARTBAS\EXCEL_REPORTS\GENERAL.BAT"

RUN_CODE = Shell(RUN_FILE, 4)

Exit Sub

CLOSE_EXCEL:

Call EXCEL_REQUEST_ERROR

End Sub

Private Sub cmdEXCEL_SITMIN_Click()

On Error GoTo CLOSE_EXCEL

Call cmdASSO_Click

rtsDISP.Visible = False
cmdPRINT.Visible = False

FileCopy APPROOT + "\ARTBAS\BLANK_WORKSHEETS\MIN_SIT.XLS", APPROOT + "\ARTBAS\EXCEL_REPORTS\ARTFISH_WORK.XLS"

RUN_FILE = APPROOT + "\ARTBAS\EXCEL_REPORTS\GENERAL.BAT"

RUN_CODE = Shell(RUN_FILE, 4)

Exit Sub

CLOSE_EXCEL:

Call EXCEL_REQUEST_ERROR

End Sub

Private Sub cmdEXCEL_SPECIES_Click()

On Error GoTo CLOSE_EXCEL

Call cmdSPECIES_Click

rtsDISP.Visible = False
cmdPRINT.Visible = False

FileCopy APPROOT + "\ARTBAS\BLANK_WORKSHEETS\SPECIES.XLS", APPROOT + "\ARTBAS\EXCEL_REPORTS\ARTFISH_WORK.XLS"

RUN_FILE = APPROOT + "\ARTBAS\EXCEL_REPORTS\GENERAL.BAT"

RUN_CODE = Shell(RUN_FILE, 4)

Exit Sub

CLOSE_EXCEL:

Call EXCEL_REQUEST_ERROR

End Sub

Private Sub cmdEXCEL_UNITS_Click()

On Error GoTo CLOSE_EXCEL

Call cmdUNITS_Click

rtsDISP.Visible = False
cmdPRINT.Visible = False

FileCopy APPROOT + "\ARTBAS\BLANK_WORKSHEETS\UNITS.XLS", APPROOT + "\ARTBAS\EXCEL_REPORTS\ARTFISH_WORK.XLS"

RUN_FILE = APPROOT + "\ARTBAS\EXCEL_REPORTS\GENERAL.BAT"

RUN_CODE = Shell(RUN_FILE, 4)

Exit Sub

CLOSE_EXCEL:

Call EXCEL_REQUEST_ERROR

End Sub

Private Sub cmdSITES_Click()

ESTDISP = "N"
Frame1.Visible = False
lstLIST.Visible = False
rtsDISP.Visible = True
cmdPRINT.Visible = True

Open APPROOT + "\ARTBAS\RESULTS\WORK.TXT" For Output As #2
Open APPROOT + "\ARTBAS\current_tables\sites.TXT" For Output As #4

Print #2, Tab(5); frmREP.Caption
Print #2, " "
Print #2, Tab(5); msgtab(43)
Print #2, " "
Print #2, Tab(5); msgtab(53); Tab(10); msgtab(54); Tab(41); msgtab(55); _
          Tab(73); msgtab(56)
Print #2, Tab(5); String(80, "=")

Write #4, " ", frmREP.Caption, msgtab(43), "  "
Write #4, " ", "  ", "  ", "  "
Write #4, msgtab(53), msgtab(54), msgtab(55), msgtab(56)
Write #4, " ", "  ", "  ", "  "

Dim fnm, XXX, V1, V2, V3, V4

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_SITES.TXT"

If Dir(fnm) = "" Then Exit Sub

Open fnm For Input As #1

Do Until EOF(1)

Line Input #1, XXX

V1 = Left(XXX, 4): V2 = LTrim(RTrim(Mid(XXX, 6, 30)))
V3 = LTrim(RTrim(Mid(XXX, 37, 31))): V4 = LTrim(RTrim(Mid(XXX, 69, 6)))

Print #2, Tab(5); XXX
Write #4, V1, V2, V3, V4

Loop

Close #1
Close #2
Close #4

rtsDISP.FileName = APPROOT + "\ARTBAS\RESULTS\WORK.TXT"
rtsDISP.Visible = True

cmdPRINT.Visible = True

End Sub
Private Sub cmdBG_Click()

ESTDISP = "N"
rtsDISP.Visible = True
cmdPRINT.Visible = True
Frame1.Visible = False
lstLIST.Visible = False

Open APPROOT + "\ARTBAS\RESULTS\WORK.TXT" For Output As #2
Open APPROOT + "\ARTBAS\CURRENT_TABLES\BG.TXT" For Output As #4

Print #2, Tab(5); frmREP.Caption
Print #2, " "
Print #2, Tab(5); msgtab(45)
Print #2, " "
Print #2, Tab(5); msgtab(53); Tab(10); msgtab(54); Tab(41); msgtab(55); _
          Tab(73); msgtab(56)
Print #2, Tab(5); String(80, "=")

Write #4, "  ", frmREP.Caption, msgtab(45), "  "
Write #4, " ", "  ", "  ", "  "
Write #4, msgtab(53), msgtab(54), msgtab(55), msgtab(56)
Write #4, " ", "  ", "  ", "  "

Dim fnm, XXX, V1, V2, V3, V4

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_BG.TXT"

If Dir(fnm) = "" Then Exit Sub

Open fnm For Input As #1

Do Until EOF(1)

Line Input #1, XXX

V1 = LTrim(RTrim(Left(XXX, 4))): V2 = LTrim(RTrim(Mid(XXX, 6, 30)))
V3 = LTrim(RTrim(Mid(XXX, 37, 31))): V4 = LTrim(RTrim(Mid(XXX, 69, 6)))

Print #2, Tab(5); XXX
Write #4, V1, V2, V3, V4
Loop

Close #1
Close #2
Close #4

rtsDISP.FileName = APPROOT + "\ARTBAS\RESULTS\WORK.TXT"
rtsDISP.Visible = True

cmdPRINT.Visible = True

End Sub
Private Sub cmdDOCS_Click()

If Frame1.Visible = True Then
   cmdDOCS.Enabled = False
   Exit Sub
   End If

lblEXP.Visible = False

rtsDISP.Visible = False
DOCMNBG = "Y"
DOCSP = "N"

If lstLIST.Visible = True Then
   lstLIST.Visible = False
   cmdESTIM.Enabled = True
   DOCMNBG = "N"
   DOCSP = "N"
   Exit Sub
   End If

cmdESTIM.Enabled = False
lstLIST.Visible = False

cmdDOCS.MousePointer = 13
frmREP.MousePointer = 13

Frame1.Visible = False
lstLIST.Visible = False
   
Call CHECK_LSAMPLES

cmdDOCS.MousePointer = 1
frmREP.MousePointer = 1

Call DISPLAY_MNBG

cmdESTIM.FontUnderline = True

End Sub
Private Sub cmdESTIM_Click()

Dim fnm

fnm = APPROOT + "\ARTBAS\RESULTS\WORK.TXT"

Open fnm For Output As #1
Close #1

cmdEXCEL_ESTIM.Visible = False

'If Frame1.Visible = True Then
   
   'cmdDOCS.Enabled = True
   'Frame1.Visible = False
   'lstLIST.Visible = False
   'ESTDISP = "N"
   'cmdDOCS.Enabled = True
   
    optMN1.Value = False
    optMN2.Value = False
    optMN3.Value = False
    optMN4.Value = False
    
    optLOG.Value = False
    
    optMJ1.Value = False
    optMJ2.Value = False
    optMJ3.Value = False
    optMJ4.Value = False
    
    optGT1.Value = False
    optGT2.Value = False
    optGT3.Value = False
    optGT4.Value = False
   
   'Exit Sub
   'End If

cmdDOCS.Enabled = False

lstLIST.Visible = False
cmdPRINT.Visible = False
rtsDISP.Visible = False
cmdESTIM.MousePointer = 13

Call LOAD_MAJOR
Call LOAD_MINOR
Call LOAD_SITES
Call LOAD_BG
Call LOAD_SPECIES
Call LOAD_ASSOMN
Call LOAD_ASSOSI

cmdESTIM.MousePointer = 1

If NTMJ * NTMN * NTSI * NTBG * NTSP * NASSOMN * NASSOSI = 0 Then
   cmdESTIM.Visible = False
   Exit Sub
   End If

Frame1.Visible = True

End Sub
Private Sub cmdFRAME_Click()

ESTDISP = "N"
Frame1.Visible = False
lstLIST.Visible = False
rtsDISP.Visible = True
cmdPRINT.Visible = True

Open APPROOT + "\ARTBAS\RESULTS\WORK.TXT" For Output As #2
Open APPROOT + "\ARTBAS\CURRENT_TABLES\FRAME.TXT" For Output As #4

Print #2, Tab(5); frmREP.Caption
Print #2, " "
Print #2, Tab(5); msgtab(46)
Print #2, " "
Print #2, Tab(5); msgtab(54); Tab(70); msgtab(67)
Print #2, Tab(5); String(80, "=")

Write #4, frmREP.Caption, msgtab(46), "  "
Write #4, " ", "  ", "  "
Write #4, msgtab(43), msgtab(139), msgtab(67)
Write #4, " ", "  ", "  "

Dim fnm, XXX, I, J, V, W

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_FRAME.TXT"

If Dir(fnm) = "" Then Exit Sub

Open fnm For Input As #1

Do Until EOF(1)

Line Input #1, XXX

I = Val(Mid(XXX, 2, 4)): J = Val(Mid(XXX, 8, 4)): V = CDbl(Mid(XXX, 26, 15))

W = Format(V, "#####0.00")

Print #2, Tab(5); SITEN(I) + " " + BGN(J) + Right(Space(9) + LTrim(W), 9)
Write #4, LTrim(RTrim(SITEN(I))), LTrim(RTrim(BGN(J))), CDbl(W)
Loop

Write #4, " ", "  ", "  "

Close #1

Call FRAME_MINOR_TOTALS

Close #2
Close #4

rtsDISP.FileName = APPROOT + "\ARTBAS\RESULTS\WORK.TXT"
rtsDISP.Visible = True

cmdPRINT.Visible = True

End Sub
Private Sub FRAME_MINOR_TOTALS()

Dim FSIC(1 To 1000), FMNC(1 To 200), FMNN(1 To 200), FBGC(1 To 200), FBGN(1 To 200)
Dim FMNBG(1 To 200, 1 To 200)
Dim I, J, K, L, XXX, fnm, V, W

' Load boat/gear table

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_BG.TXT"

Open fnm For Input As #5

For I = 1 To 200
FBGN(I) = " "
Next I

Do While Not EOF(5)
Line Input #5, XXX
J = Val(Left(XXX, 4))
FBGN(J) = Mid(XXX, 6, 30)
Loop

Close #5

' Load MN table
'===============

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_ASSOSI.TXT"

If Dir(fnm) = "" Then Exit Sub

Open fnm For Input As #5

For I = 1 To 200
FMNN(I) = " "
Next I

For I = 1 To 200
For J = 1 To 200
FMNBG(I, J) = 0
Next J
Next I

Do While Not EOF(5)

Line Input #5, XXX

I = Val(Mid(XXX, 1, 4)): FMNN(I) = Mid(XXX, 6, 30): K = Val(Mid(XXX, 37, 4))

For J = 1 To K
Line Input #5, XXX
L = Val(Mid(XXX, 6, 4))
FSIC(L) = I
Next J

Loop

Close #5

'========================================
' READ FRAME

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_FRAME.TXT"

Open fnm For Input As #5

Do Until EOF(5)

Line Input #5, XXX

L = Val(Mid(XXX, 2, 4)): J = Val(Mid(XXX, 8, 4)): V = CDbl(Mid(XXX, 26, 15))

W = Format(V, "#####0.00")

I = FSIC(L)

FMNBG(I, J) = FMNBG(I, J) + V

DOCTMNBG(I, J) = FMNBG(I, J)

Loop

Close #5

'==================================================================
' Start supplementary report

Print #2, " "
Write #4, " ", " ", " "

Print #2, Tab(5); String(80, "=")
Print #2, Tab(5); msgtab(267)
Write #4, msgtab(267), " ", " "
Write #4, " ", " ", " "

Print #2, " "

For I = 1 To 200
For J = 1 To 200

If FMNN(I) = " " Or FBGN(J) = " " Then GoTo next_j

W = Format(FMNBG(I, J), "#####0.00")

Print #2, Tab(5); FMNN(I) + " " + FBGN(J) + Right(Space(9) + LTrim(W), 9)
Write #4, FMNN(I), FBGN(J), CDbl(W)

Print #2, " "
Write #4, " ", " ", " "

If FMNBG(I, J) < 10 Then GoTo next_j

Print #2, Tab(5); msgtab(262); Tab(40); "90%"; Tab(60); "95%"
Write #4, msgtab(262), " --- 90 % ---", "--- 95 % ---"
Print #2, " "
Write #4, " ", " ", " "

' Landings

Dim SMP90, SMP95

POPSIZE = 31 * FMNBG(I, J)
CONVEX_YN = "Y"
INACC = 0.9

Call SAMPLES_FOR_GIVEN_ACCURACY
SMP90 = OUTSMP

INACC = 0.95
Call SAMPLES_FOR_GIVEN_ACCURACY
SMP95 = OUTSMP

Print #2, Tab(5); msgtab(264); Tab(40); Int(SMP90); Tab(60); Int(SMP95)
Write #4, msgtab(264), Int(SMP90), Int(SMP95)

POPSIZE = 31 * FMNBG(I, J)
CONVEX_YN = "N"
INACC = 0.9

Call SAMPLES_FOR_GIVEN_ACCURACY
SMP90 = OUTSMP

INACC = 0.95
Call SAMPLES_FOR_GIVEN_ACCURACY
SMP95 = OUTSMP

Print #2, Tab(5); msgtab(265); Tab(40); Int(SMP90); Tab(60); Int(SMP95)
Write #4, msgtab(265), Int(SMP90), Int(SMP95)

POPSIZE = FMNBG(I, J)
CONVEX_YN = "Y"
INACC = 0.9

Call SAMPLES_FOR_GIVEN_ACCURACY
SMP90 = OUTSMP

INACC = 0.95
Call SAMPLES_FOR_GIVEN_ACCURACY
SMP95 = OUTSMP

Print #2, Tab(5); msgtab(266); Tab(40); Int(SMP90); Tab(60); Int(SMP95)
Write #4, msgtab(266), Int(SMP90), Int(SMP95)

Print #2, "  "
Write #4, " ", " ", " "

next_j:

Next J
Next I

End Sub
Private Sub cmdGUIDE_Click()

HTYPE = "D0"

HFNM = APPROOT + "\ARTBAS\HELP\" + current_language + "HELP" + HTYPE + ".rtf"

If Dir(HFNM) = "" Then Exit Sub

frmREP.Enabled = False
Load frmGUIDE
frmGUIDE.Show

End Sub
Private Sub cmdMAJOR_Click()

ESTDISP = "N"
Frame1.Visible = False
lstLIST.Visible = False

rtsDISP.Visible = True
cmdPRINT.Visible = True

'============================================================
Open APPROOT + "\ARTBAS\RESULTS\WORK.TXT" For Output As #2
Open APPROOT + "\ARTBAS\CURRENT_TABLES\MAJOR.TXT" For Output As #4

Print #2, Tab(5); frmREP.Caption
Print #2, " "
Print #2, Tab(5); msgtab(41)
Print #2, " "
Print #2, Tab(5); msgtab(53); Tab(10); msgtab(54); Tab(41); msgtab(55); _
          Tab(73); msgtab(56)
Print #2, Tab(5); String(80, "=")

Write #4, "  ", frmREP.Caption, msgtab(41), "  "
Write #4, " ", "  ", "  ", "  "
Write #4, msgtab(53), msgtab(54), msgtab(55), msgtab(56)
Write #4, " ", "  ", "  ", "  "

Dim fnm, XXX, FNMFRAME, V1, V2, V3, V4

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_MAJOR.TXT"

If Dir(fnm) = "" Then Exit Sub

Open fnm For Input As #1

Do Until EOF(1)

Line Input #1, XXX

V1 = LTrim(RTrim(Left(XXX, 4))): V2 = LTrim(RTrim(Mid(XXX, 6, 30)))
V3 = LTrim(RTrim(Mid(XXX, 37, 31))): V4 = LTrim(RTrim(Mid(XXX, 69, 6)))

Print #2, Tab(5); XXX
Write #4, V1, V2, V3, V4

Loop

Close #1
Close #2
Close #4

rtsDISP.FileName = APPROOT + "\ARTBAS\RESULTS\WORK.TXT"

End Sub
Private Sub cmdMINOR_Click()

ESTDISP = "N"

rtsDISP.Visible = True
cmdPRINT.Visible = True

Frame1.Visible = False
lstLIST.Visible = False

'===========================================================
Open APPROOT + "\ARTBAS\RESULTS\WORK.TXT" For Output As #2
Open APPROOT + "\ARTBAS\CURRENT_TABLES\MINOR.TXT" For Output As #4

Print #2, Tab(5); frmREP.Caption
Print #2, " "
Print #2, Tab(5); msgtab(42)
Print #2, " "
Print #2, Tab(5); msgtab(53); Tab(10); msgtab(54); Tab(41); msgtab(55); _
          Tab(73); msgtab(56)
Print #2, Tab(5); String(80, "=")

Write #4, "  ", frmREP.Caption, msgtab(42), "  "
Write #4, " ", "  ", "  ", "  "
Write #4, msgtab(53), msgtab(54), msgtab(55), msgtab(56)
Write #4, " ", "  ", "  ", "  "

Dim fnm, XXX, FNMFRAME, V1, V2, V3, V4

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_MINOR.TXT"

If Dir(fnm) = "" Then Exit Sub

Open fnm For Input As #1

Do Until EOF(1)

Line Input #1, XXX

V1 = LTrim(RTrim(Left(XXX, 4))): V2 = LTrim(RTrim(Mid(XXX, 6, 30)))
V3 = LTrim(RTrim(Mid(XXX, 37, 31))): V4 = LTrim(RTrim(Mid(XXX, 69, 6)))

Print #2, Tab(5); XXX
Write #4, V1, V2, V3, V4
Loop

Close #1
Close #2
Close #4

rtsDISP.FileName = APPROOT + "\ARTBAS\RESULTS\WORK.TXT"

End Sub
Private Sub cmdPRINT_Click()

Printer.FontBold = True
Printer.FontName = "Courier"
Printer.FontName = "Courier New"
Printer.FontSize = 9
Printer.FontItalic = False

Dim I, pageno, lineno

pageno = 0

GoSub CHANGE_PAGE

Dim XXX, fnm As String

fnm = APPROOT + "\ARTBAS\RESULTS\WORK.TXT"

If Dir(fnm) = "" Then Exit Sub

Open fnm For Input As #1

Do Until EOF(1)

Line Input #1, XXX

If Len(XXX) >= 200 And Right(XXX, 1) = "<" Then GoSub CHANGE_PAGE

Printer.Print XXX

lineno = lineno + 1

If lineno > 70 Then GoSub CHANGE_PAGE

Loop

Close #1

Printer.EndDoc

Exit Sub

'========================
CHANGE_PAGE:

lineno = 5
pageno = pageno + 1
If pageno > 1 Then Printer.NewPage

Printer.Print " "
Printer.Print " "

Printer.Print , Tab(50); pageno

Printer.Print " "
Printer.Print " "

Return
'====================================

End Sub
Private Sub cmdQUIT_Click()

Close #3

Dim fnm

Call CHECK_BACKUP_COMPLETE

fnm = APPROOT + "\ARTBAS\RESULTS\W*.*"

If Dir(fnm) <> "" Then Kill fnm

cmdRETURN.MousePointer = 13
cmdQUIT.MousePointer = 13

Call KILL_ARTBASIC_FOLDER

Call write_parms
Unload frmREP

End

End Sub
Private Sub cmdRETURN_Click()

Close #3

Dim fnm

fnm = APPROOT + "\ARTBAS\RESULTS\W*.*"

If Dir(fnm) <> "" Then Kill fnm

cmdRETURN.MousePointer = 13
frmREP.MousePointer = 13
Load frmARTB01
Unload frmREP
frmARTB01.Show

End Sub
Private Sub cmdSPECIES_Click()

ESTDISP = "N"
rtsDISP.Visible = True
cmdPRINT.Visible = True
Frame1.Visible = False
lstLIST.Visible = False

Open APPROOT + "\ARTBAS\RESULTS\WORK.TXT" For Output As #2
Open APPROOT + "\ARTBAS\CURRENT_TABLES\SPECIES.TXT" For Output As #4

Print #2, Tab(5); frmREP.Caption
Print #2, " "
Print #2, Tab(5); msgtab(47)
Print #2, " "
Print #2, Tab(5); msgtab(53); Tab(10); msgtab(54); Tab(41); msgtab(55); _
          Tab(73); msgtab(56)
Print #2, Tab(5); String(80, "=")

Write #4, " ", frmREP.Caption, msgtab(47), "  "
Write #4, " ", "  ", "  ", "  "

Write #4, msgtab(53), msgtab(54), msgtab(55), msgtab(56)
Write #4, " ", "  ", "  ", "  "

Dim fnm, XXX, V1, V2, V3, V4

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_SPECIES.TXT"

If Dir(fnm) = "" Then Exit Sub

Open fnm For Input As #1

Do Until EOF(1)

Line Input #1, XXX

V1 = Left(XXX, 4): V2 = LTrim(RTrim(Mid(XXX, 6, 30)))
V3 = LTrim(RTrim(Mid(XXX, 37, 31))): V4 = LTrim(RTrim(Mid(XXX, 69, 6)))

Print #2, Tab(5); XXX
Write #4, V1, V2, V3, V4

Loop

Close #1
Close #2
Close #4

rtsDISP.FileName = APPROOT + "\ARTBAS\RESULTS\WORK.TXT"
rtsDISP.Visible = True

cmdPRINT.Visible = True
End Sub
Private Sub cmdUNITS_Click()
ESTDISP = "N"
cmdPRINT.Visible = True
rtsDISP.Visible = True
Frame1.Visible = False
lstLIST.Visible = False

Open APPROOT + "\ARTBAS\RESULTS\WORK.TXT" For Output As #2
Open APPROOT + "\ARTBAS\current_tables\UNITS.TXT" For Output As #4

Print #2, Tab(5); frmREP.Caption
Write #4, frmREP.Caption

Print #2, " "
Write #4, "  "

Print #2, Tab(5); msgtab(48)
Write #4, msgtab(48)
Write #4, "  "

Print #2, Tab(5); String(80, "=")

Print #2, " "

Dim fnm, XXX, I, J, V, W, yyy

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_UNITS.txt"

If Dir(fnm) = "" Then Exit Sub

Open fnm For Input As #1

Dim UNW, UNM

Line Input #1, UNW
Line Input #1, UNM

Close #1

Print #2, Tab(5); msgtab(192) + " " + UNW
Print #2, Tab(5); msgtab(193) + " " + msgtab(198)
Print #2, Tab(5); msgtab(200) + " " + UNW + "/" + msgtab(198)
Print #2, Tab(5); msgtab(194) + " " + UNM + "/" + UNW
Print #2, Tab(5); msgtab(195) + " " + UNM
Print #2, Tab(5); msgtab(196) + " " + UNW + "/" + msgtab(197)

Write #4, msgtab(192) + " " + UNW
Write #4, msgtab(193) + " " + msgtab(198)
Write #4, msgtab(200) + " " + UNW + "/" + msgtab(198)
Write #4, msgtab(194) + " " + UNM + "/" + UNW
Write #4, msgtab(195) + " " + UNM
Write #4, msgtab(196) + " " + UNW + "/" + msgtab(197)

Close #2
Close #4

rtsDISP.FileName = APPROOT + "\ARTBAS\RESULTS\WORK.TXT"
rtsDISP.Visible = True

cmdPRINT.Visible = True
End Sub
Private Sub Form_Click()

lblEXP.Visible = False

If cmdDOCS.Visible = True Then
   cmdESTIM.Enabled = False
   Exit Sub
   End If

cmdEXCEL_RAW.Visible = False
cmdESTIM.Enabled = True

Frame1.Visible = False
lstLIST.Visible = False
'cmdESTIM.Enabled = True
'cmdDOCS.Enabled = True

optMN1.Value = False
optMN2.Value = False
optMN3.Value = False
optMN4.Value = False

optLOG.Value = False

optMJ1.Value = False
optMJ2.Value = False
optMJ3.Value = False
optMJ4.Value = False

optGT1.Value = False
optGT2.Value = False
optGT3.Value = False
optGT4.Value = False

rtsDISP.Visible = False
lblEXP.Visible = False
cmdPRINT.Visible = False

If ESTDISP = "Y" And cmdESTIM.Enabled = True Then
   Frame1.Visible = True
   lstLIST.Visible = False
   End If

End Sub
Private Sub Form_Load()

optLOC.Value = True

LOG_OPTION = "NO"

cmdEXCEL_RAW.Visible = False
cmdEXCEL_ESTIM.Visible = False

Set Picture = LoadPicture(APPROOT + "\ARTBAS\PICS_RUNTIME\SCREEN_11.JPG")

lblEXP.Visible = False
lblEXP.Caption = msgtab(245) + " " + msgtab(246)

Dim fnm, XXX, fnm2, FNM1

fnm = APPROOT + "\ARTBAS\TRANSFER\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "*.*"

If Dir(fnm) = "" Then GoTo NO_TRANSFER

fnm = Dir(fnm)
FNM1 = APPROOT + "\ARTBAS\TRANSFER\" + fnm
fnm2 = APPROOT + "\ARTBAS\RESULTS\" + fnm

FileCopy FNM1, fnm2

fnm = "?"

Do Until fnm = ""

fnm = Dir

FNM1 = APPROOT + "\ARTBAS\TRANSFER\" + fnm
fnm2 = APPROOT + "\ARTBAS\RESULTS\" + fnm

If fnm <> "" Then FileCopy FNM1, fnm2

Loop

NO_TRANSFER:

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_UNITS.TXT"

If Dir(fnm) <> "" Then
   Open fnm For Input As #1
   Line Input #1, XXX
   UNW = RTrim(Left(XXX, 6))
   Line Input #1, XXX
   UNM = RTrim(Left(XXX, 10))
   Close #1
   End If

If Dir(fnm) = "" Then
   UNW = "Kg": UNM = "ÜS$"
   End If

frmREP.Caption = monthtab(current_month) + " " + _
                    Format(current_year, "0000") + " - " + msgtab(38)


frmREP.MousePointer = 1

Frame1.MousePointer = 1
cmdPRINT.Visible = False
rtsDISP.Visible = False
lstLIST.Visible = False

Frame1.Visible = False

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
cmdDOCS.ToolTipText = msgtab(247)
cmdPRINT.ToolTipText = msgtab(52)
cmdESTIM.ToolTipText = msgtab(37)
cmdACT.ToolTipText = msgtab(115)
cmdGUIDE.ToolTipText = msgtab(243)

cmdEXCEL_MAJOR.ToolTipText = msgtab(251)
cmdEXCEL_MINOR.ToolTipText = msgtab(251)
cmdEXCEL_MINMAJ.ToolTipText = msgtab(251)
cmdEXCEL_SITES.ToolTipText = msgtab(251)
cmdEXCEL_SITMIN.ToolTipText = msgtab(251)
cmdEXCEL_BG.ToolTipText = msgtab(251)
cmdEXCEL_FRAME.ToolTipText = msgtab(251)
cmdEXCEL_SPECIES.ToolTipText = msgtab(251)
cmdEXCEL_RAW.ToolTipText = msgtab(251)
cmdEXCEL_ESTIM.ToolTipText = msgtab(251)
cmdEXCEL_ACTIVE.ToolTipText = msgtab(251)
cmdEXCEL_UNITS.ToolTipText = msgtab(251)

lblOPTIONS.Caption = msgtab(231)

lblMN.Caption = msgtab(141)
lblMJ.Caption = msgtab(142)
lblGT.Caption = msgtab(143)

optMN1.Caption = msgtab(144)
optMN2.Caption = msgtab(145)
optMN3.Caption = msgtab(146)
optMN4.Caption = msgtab(147)

optMJ1.Caption = msgtab(144)
optMJ2.Caption = msgtab(145)
optMJ3.Caption = msgtab(146)
optMJ4.Caption = msgtab(147)

optGT1.Caption = msgtab(144)
optGT2.Caption = msgtab(145)
optGT3.Caption = msgtab(146)
optGT4.Caption = msgtab(147)

optLOG.Caption = msgtab(121)

optSTD.Caption = msgtab(274)
optLOC.Caption = msgtab(275)

Call CHECK_STATUS
Call LOAD_TABLES

If NOSI * NOBG = 0 Then
   cmdFRAME.Visible = False
   cmdEXCEL_FRAME.Visible = False
   End If

If NOMJ * NOMN = 0 Then
   cmdASSOM.Visible = False
   cmdEXCEL_MINMAJ.Visible = False
   End If
   
If NOMN * NOSI = 0 Then
   cmdASSO.Visible = False
   cmdEXCEL_SITMIN.Visible = False
   End If
   
If NOMN * NOBG = 0 Then
   cmdACT.Visible = False
   cmdEXCEL_ACTIVE.Visible = False
   End If

If NOSI * NOBG * NOSP = 0 Then
   cmdESTIM.Visible = False
   cmdDOCS.Visible = False
   cmdEXCEL_RAW.Visible = False
   cmdEXCEL_ESTIM.Visible = False
   End If
   
End Sub
Private Sub CHECK_STATUS()

Dim fnm

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_MAJOR.TXT"
      
If Dir(fnm) = "" Then
   cmdMAJOR.Visible = False
   cmdEXCEL_MAJOR.Visible = False
   End If

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_MINOR.TXT"
      
If Dir(fnm) = "" Then
   cmdMINOR.Visible = False
   cmdEXCEL_MINOR.Visible = False
   End If
   
fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_ASSOMN.TXT"
      
If Dir(fnm) = "" Then
   cmdASSOM.Visible = False
   cmdEXCEL_MINMAJ.Visible = False
   End If
   
fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_SITES.TXT"
      
If Dir(fnm) = "" Then
   cmdSITES.Visible = False
   cmdEXCEL_SITES.Visible = False
   End If
   
fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_ASSOSI.TXT"
      
If Dir(fnm) = "" Then
   cmdASSO.Visible = False
   cmdEXCEL_SITMIN.Visible = False
   End If

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_BG.TXT"

If Dir(fnm) = "" Then
   cmdBG.Visible = False
   cmdEXCEL_BG.Visible = False
   End If
   
fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_FRAME.TXT"
      
If Dir(fnm) = "" Then
   cmdFRAME.Visible = False
   cmdEXCEL_FRAME.Visible = False
   End If
   
fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_SPECIES.TXT"
      
If Dir(fnm) = "" Then
   cmdSPECIES.Visible = False
   cmdEXCEL_SPECIES.Visible = False
   End If
   
fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_UNITS.TXT"
      
If Dir(fnm) = "" Then
   cmdUNITS.Visible = False
   cmdEXCEL_UNITS.Visible = False
   End If

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_ACTIVE.TXT"
      
If Dir(fnm) = "" Then
   cmdACT.Visible = False
   cmdEXCEL_ACTIVE.Visible = False
   cmdDOCS.Visible = False
   cmdEXCEL_RAW.Visible = False
   End If

fnm = APPROOT + "\ARTBAS\RESULTS\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_*.*"
      
If Dir(fnm) = "" Then
   cmdESTIM.Visible = False
   cmdEXCEL_ESTIM.Visible = False
   End If

fnm = APPROOT + "\ARTBAS\LANDINGS\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_LSAMPLES.TXT"
      
If Dir(fnm) = "" Then
   cmdDOCS.Visible = False
   cmdEXCEL_RAW.Visible = False
   End If

fnm = APPROOT + "\ARTBAS\EFFORT\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_ESAMPLES.TXT"
      
If Dir(fnm) = "" Then
   cmdDOCS.Visible = False
   cmdEXCEL_RAW.Visible = False
   End If

End Sub

Private Sub lblOPTIONS_Click()

cmdEXCEL_ESTIM.Visible = False

ESTDISP = "N"

optMN1.Value = False
optMN2.Value = False
optMN3.Value = False
optMN4.Value = False

optLOG.Value = False

optMJ1.Value = False
optMJ2.Value = False
optMJ3.Value = False
optMJ4.Value = False

optGT1.Value = False
optGT2.Value = False
optGT3.Value = False
optGT4.Value = False

cmdDOCS.Enabled = True
Frame1.Visible = False
lstLIST.Visible = False

End Sub

Private Sub lstLIST_Click()

cmdEXCEL_RAW.Visible = False
cmdEXCEL_ESTIM.Visible = False

If DOCMNBG = "Y" And DOCSP = "N" Then
   lstLIST.Visible = False
   Call SELECT_MNBG
   Call DOC_TOTALS
   Exit Sub
   End If

'If DOCMNBG = "N" And DOCSP = "Y" Then
'   Call SELECT_MBS
'   Exit Sub
'   End If

If optMN1.Value = True Or optMN2.Value = True Or optMN3.Value = True _
   Or optMN4.Value = True Or optLOG.Value = True Then
   
   Call MINOR_STRATA
   
   End If

If optMJ1.Value = True Or optMJ2.Value = True Or optMJ3.Value = True _
   Or optMJ4.Value = True Then
   
   Call MAJOR_STRATA
   
   End If

End Sub
Private Sub MAJOR_STRATA()

cmdEXCEL_ESTIM.Visible = True

Dim I, resp

I = lstLIST.ListIndex + 1

If TMJOK(I) <> "+" Then cmdEXCEL_ESTIM.Visible = False

If TMJOK(I) <> "+" Then
   resp = MsgBox(msgtab(188), vbCritical, " ")
   optMJ1.Value = False: optMJ2.Value = False
   optMJ3.Value = False: optMJ4.Value = False
   lstLIST.Visible = False
   Exit Sub
   End If

CURMJC = TMJC(I): ESTFLAG = "N"

Call CREATE_MJDB

If ESTFLAG <> "Y" Then
   resp = MsgBox(msgtab(189), vbCritical, " ")
   optMJ1.Value = False: optMJ2.Value = False
   optMJ3.Value = False: optMJ4.Value = False
   lstLIST.Visible = False
   Exit Sub
   End If

cmdPRINT.Visible = True

If LOG_OPTION <> "YES" Then cmdEXCEL_ESTIM.Visible = True

If optMJ1.Value = True Then
   optMJ1.Value = False
   Call BYMJ_BYBG_BYSP
   Exit Sub
   End If

If optMJ2.Value = True Then
   optMJ2.Value = False
   Call BYMJ_BYSP_BYBG
   Exit Sub
   End If

If optMJ3.Value = True Then
   optMJ3.Value = False
   Call BYMJ_BYBG
   Exit Sub
   End If

If optMJ4.Value = True Then
   optMJ4.Value = False
   Call BYMJ_BYSP
   Exit Sub
   End If

End Sub
Private Sub GRAND_TOTALS()

cmdEXCEL_ESTIM.Visible = True

Dim I, resp

Call CREATE_GTDB

If ESTFLAG <> "Y" Then
   resp = MsgBox(msgtab(189), vbCritical, " ")
   optGT1.Value = False: optGT2.Value = False
   optGT3.Value = False: optGT4.Value = False
   Exit Sub
   End If

cmdPRINT.Visible = True

If optGT1.Value = True Then
   optGT1.Value = False
   Call BYGT_BYBG_BYSP
   Exit Sub
   End If

If optGT2.Value = True Then
   optGT2.Value = False
   Call BYGT_BYSP_BYBG
   Exit Sub
   End If

If optGT3.Value = True Then
   optGT3.Value = False
   Call BYGT_BYBG
   Exit Sub
   End If

If optGT4.Value = True Then
   optGT4.Value = False
   Call BYGT_BYSP
   Exit Sub
   End If

End Sub
Private Sub MINOR_STRATA()

Dim I, resp

I = lstLIST.ListIndex + 1

If TMNOK(I) <> "+" Then cmdEXCEL_ESTIM.Visible = False

If TMNOK(I) <> "+" And optLOG.Value = False Then
   resp = MsgBox(msgtab(188), vbCritical, " ")
   optMN1.Value = False: optMN2.Value = False
   optMN3.Value = False: optMN4.Value = False: optLOG.Value = False
   lstLIST.Visible = False
   Exit Sub
   End If

CURMNC = TMNC(I): ESTFLAG = "N"

Call CREATE_MNDB

If ESTFLAG <> "Y" And optLOG.Value = False Then
   resp = MsgBox(msgtab(189), vbCritical, " ")
   optMN1.Value = False: optMN2.Value = False
   optMN3.Value = False: optMN4.Value = False
   lstLIST.Visible = False
   Exit Sub
   End If

cmdPRINT.Visible = True

If optLOG.Value = True Then
   optLOG.Value = False
   Call DISP_LOG
   Exit Sub
   End If

If LOG_OPTION <> "YES" Then cmdEXCEL_ESTIM.Visible = True

If optMN1.Value = True Then
   optMN1.Value = False
   Call BYMN_BYBG_BYSP
   Exit Sub
   End If

If optMN2.Value = True Then
   optMN2.Value = False
   Call BYMN_BYSP_BYBG
   Exit Sub
   End If

If optMN3.Value = True Then
   optMN3.Value = False
   Call BYMN_BYBG
   Exit Sub
   End If

If optMN4.Value = True Then
   optMN4.Value = False
   Call BYMN_BYSP
   Exit Sub
   End If

End Sub
Private Sub DISP_LOG()

Dim fnm, XXX, fnm2

fnm2 = APPROOT + "\ARTBAS\CURRENT_TABLES\LOG.TXT"

fnm = APPROOT + "\ARTBAS\RESULTS\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_MN" + Format(CURMNC, "0000") + "_LOG.TXT"

If Dir(fnm) = "" Then Exit Sub

FileCopy fnm, fnm2

rtsDISP.FileName = fnm

Frame1.Visible = False
lstLIST.Visible = False
rtsDISP.Visible = True
cmdPRINT.Visible = True
cmdEXCEL_ESTIM.Visible = True

Close #1

End Sub
Private Sub optGT1_Click()

LOG_OPTION = "NO"
ESTDISP = "Y"
Call GRAND_TOTALS

End Sub
Private Sub optGT2_Click()

LOG_OPTION = "NO"
ESTDISP = "Y"
Call GRAND_TOTALS

End Sub
Private Sub optGT3_Click()

LOG_OPTION = "NO"
ESTDISP = "Y"
Call GRAND_TOTALS

End Sub
Private Sub optGT4_Click()

LOG_OPTION = "NO"
ESTDISP = "Y"
Call GRAND_TOTALS

End Sub

Private Sub optLOC_Click()

If rtsDISP.Visible = True Then rtsDISP.Visible = False
If lstLIST.Visible = True Then lstLIST.Visible = False
If Frame1.Visible = True Then Frame1.Visible = False

Call LOAD_TABLES

End Sub

Private Sub optLOG_Click()

LOG_OPTION = "YES"
ESTDISP = "Y"
Call LIST_MINOR

End Sub
Private Sub optMJ1_Click()

LOG_OPTION = "NO"
ESTDISP = "Y"
Call LIST_MAJOR

End Sub
Private Sub optMJ2_Click()

LOG_OPTION = "NO"
ESTDISP = "Y"
Call LIST_MAJOR

End Sub
Private Sub optMJ3_Click()

LOG_OPTION = "NO"
ESTDISP = "Y"
Call LIST_MAJOR

End Sub
Private Sub optMJ4_Click()

LOG_OPTION = "NO"
ESTDISP = "Y"
Call LIST_MAJOR

End Sub
Private Sub optMN1_Click()

LOG_OPTION = "NO"
ESTDISP = "Y"
Call LIST_MINOR

End Sub
Private Sub optMN2_Click()

LOG_OPTION = "NO"
ESTDISP = "Y"
Call LIST_MINOR

End Sub
Private Sub optMN3_Click()

LOG_OPTION = "NO"
ESTDISP = "Y"
Call LIST_MINOR

End Sub
Private Sub optMN4_Click()

LOG_OPTION = "NO"
ESTDISP = "Y"
Call LIST_MINOR

End Sub

Private Sub optSTD_Click()

If rtsDISP.Visible = True Then rtsDISP.Visible = False
If lstLIST.Visible = True Then lstLIST.Visible = False
If Frame1.Visible = True Then Frame1.Visible = False

Call LOAD_TABLES

End Sub

Private Sub rtsDISP_Click()

cmdEXCEL_RAW.Visible = False
cmdEXCEL_ESTIM.Visible = False

rtsDISP.Visible = False
lblEXP.Visible = False
cmdPRINT.Visible = False
cmdESTIM.Enabled = True


If ESTDISP = "Y" Then
   Frame1.Visible = True
   lstLIST.Visible = False
   End If

End Sub
Private Sub LOAD_TABLES()

NOMJ = 0: NOMN = 0: NOSI = 0: NOBG = 0: NOSP = 0

Dim fnm, XXX, I

ReDim MAJORN(1 To 10000), SITEN(1 To 10000), MINORN(1 To 10000)
ReDim SPEN(1 To 10000), BGN(1 To 10000)

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_MAJOR.TXT"

If Dir(fnm) = "" Then GoTo NEXT_LOAD1

Open fnm For Input As #1

Do Until EOF(1)

Line Input #1, XXX

NOMJ = NOMJ + 1

I = Val(Left(XXX, 4))
MAJORN(I) = Mid(XXX, 6, 30)

If optSTD.Value = True Then MAJORN(I) = Mid(XXX, 37, 30)

Loop

Close #1

NEXT_LOAD1:

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_MINOR.TXT"

If Dir(fnm) = "" Then GoTo NEXT_LOAD2

Open fnm For Input As #1

Do Until EOF(1)

Line Input #1, XXX

NOMN = NOMN + 1

I = Val(Left(XXX, 4))
MINORN(I) = Mid(XXX, 6, 30)

If optSTD.Value = True Then MINORN(I) = Mid(XXX, 37, 30)


Loop

Close #1

NEXT_LOAD2:

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_SITES.TXT"

If Dir(fnm) = "" Then GoTo NEXT_LOAD3

Open fnm For Input As #1

Do Until EOF(1)

NOSI = NOSI + 1
Line Input #1, XXX

I = Val(Left(XXX, 4))
SITEN(I) = Mid(XXX, 6, 30)

If optSTD.Value = True Then SITEN(I) = Mid(XXX, 37, 30)

Loop

Close #1

NEXT_LOAD3:

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_BG.TXT"

If Dir(fnm) = "" Then GoTo NEXT_LOAD4

Open fnm For Input As #1

Do Until EOF(1)

Line Input #1, XXX

NOBG = NOBG + 1

I = Val(Left(XXX, 4))
BGN(I) = Mid(XXX, 6, 30)

If optSTD.Value = True Then BGN(I) = Mid(XXX, 37, 30)

Loop

Close #1

NEXT_LOAD4:

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_SPECIES.TXT"

If Dir(fnm) = "" Then Exit Sub

Open fnm For Input As #1

Do Until EOF(1)

Line Input #1, XXX

NOSP = NOSP + 1
I = Val(Left(XXX, 4))
SPEN(I) = Mid(XXX, 6, 30)

If optSTD.Value = True Then SPEN(I) = Mid(XXX, 37, 30)

Loop

Close #1

End Sub
Private Sub CREATE_MNDB()

ESTFLAG = "Y"

Dim estdes, popn, smpn, BAC_ACCUR, FRNO, actno, cal, eact, _
          esmp, esites, edays, bac, bac_cvs, bac_cvsp, _
          bac_cvt, bac_cvtp, bac_cv, bac_low, bac_upper, eff, _
          eff_low, eff_upper, nland, CPUE_ACCUR, LPOP, ltot, _
          lsmpv, lsmpf, leff, cpue, lsites, ldays, _
          cpue_cvs, cpue_cvsp, cpue_cvt, cpue_cvtp, cpue_cv, cpue_low, _
          cpue_upper, catch, catch_low, catch_upper, catch_cv, Value, _
          price, fish, kgfish


Dim dbn, CREC

FileCopy APPROOT + "\ARTBAS\STRUS\ESTOT.MDB", APPROOT + "\ARTBAS\RESULTS\WORK.MDB"

dbn = APPROOT + "\ARTBAS\RESULTS\WORK.MDB"

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(dbn)
Set prm_record = prm_database.OpenRecordset("ESTAB")

With prm_record

Dim I, fnm, IX, XKEY, XXX

Dim TYPEREC

fnm = APPROOT + "\ARTBAS\RESULTS\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_MN" + Format(CURMNC, "0000") + "_ESTIM.TXT"

If Dir(fnm) = "" Then Exit Sub

Open fnm For Input As #1

Do Until EOF(1)

Input #1, XKEY

.AddNew

![estkey] = XKEY

TYPEREC = 2

If Mid(XKEY, 14, 4) = "0000" And Mid(XKEY, 8, 4) <> "0000" Then TYPEREC = 1

If TYPEREC = 1 Then

Input #1, estdes, popn, smpn, BAC_ACCUR, FRNO, actno, cal, eact, _
          esmp, esites, edays, bac, bac_cvs, bac_cvsp, _
          bac_cvt, bac_cvtp, bac_cv, bac_low, bac_upper, eff, _
          eff_low, eff_upper, nland, CPUE_ACCUR, LPOP, ltot, _
          lsmpv, lsmpf, leff, cpue, lsites, ldays, _
          cpue_cvs, cpue_cvsp, cpue_cvt, cpue_cvtp, cpue_cv, cpue_low, _
          cpue_upper, catch, catch_low, catch_upper, catch_cv, Value, _
          price, fish, kgfish
          
          ![smpn] = smpn
          ![estdes] = estdes: ![popn] = popn: ![BAC_ACCUR] = BAC_ACCUR
          ![FRNO] = FRNO: ![actno] = actno: ![cal] = cal: ![eact] = eact
          ![esmp] = esmp: ![esites] = esites: ![edays] = edays: ![bac] = bac: ![bac_cvs] = bac_cvs
          ![bac_cvsp] = bac_cvsp: ![bac_cvt] = bac_cvt: ![bac_cvtp] = bac_cvtp: ![bac_cv] = bac_cv
          ![bac_low] = bac_low: ![bac_upper] = bac_upper: ![eff] = eff: ![eff_low] = eff_low
          ![eff_upper] = eff_upper: ![nland] = nland: ![CPUE_ACCUR] = CPUE_ACCUR
          ![LPOP] = LPOP: ![ltot] = ltot: ![lsmpv] = lsmpv: ![lsmpf] = lsmpf: ![leff] = leff
          ![cpue] = cpue: ![lsites] = lsites: ![ldays] = ldays: ![cpue_cvs] = cpue_cvs
          ![cpue_cvsp] = cpue_cvsp: ![cpue_cvt] = cpue_cvt: ![cpue_cvtp] = cpue_cvtp
          ![cpue_cv] = cpue_cv: ![cpue_low] = cpue_low: ![cpue_upper] = cpue_upper
          ![catch] = catch: ![catch_low] = catch_low: ![catch_upper] = catch_upper
          ![catch_cv] = catch_cv: ![Value] = Value: ![price] = price: ![fish] = fish: ![kgfish] = kgfish
          

          End If

If TYPEREC = 2 Then

Input #1, estdes, eff, cpue, catch, Value, price, kgfish, fish, FRNO

![estdes] = estdes: ![eff] = eff: ![cpue] = cpue: ![catch] = catch: ![Value] = Value
![price] = price: ![kgfish] = kgfish: ![fish] = fish: ![FRNO] = FRNO

         End If

.Update

If Mid(XKEY, 14, 4) = "0000" And Mid(XKEY, 8, 4) = "0000" Then
   If eff * catch = 0 Then ESTFLAG = "N"
   End If

Loop

Close #1

End With

prm_record.Close
prm_database.Close

End Sub
Private Sub CREATE_MJDB()

Dim I, J, K, fnm

Dim dbn, CREC

FileCopy APPROOT + "\ARTBAS\STRUS\ESTOT.MDB", APPROOT + "\ARTBAS\RESULTS\WORK.MDB"

dbn = APPROOT + "\ARTBAS\RESULTS\WORK.MDB"

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(dbn)
Set prm_record = prm_database.OpenRecordset("ESTAB")

With prm_record

.Index = "PRIMARYKEY"

NOEXCL = 0

For J = 1 To NTMN

K = Val(TMNC(J))

If ASSOMN(K) <> Val(CURMJC) Then GoTo next_j

fnm = APPROOT + "\ARTBAS\RESULTS\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_MN" + TMNC(J) + "_ESTIM.TXT"

If Dir(fnm) = "" Then

   NOEXCL = NOEXCL + 1
   
   ReDim Preserve EXN(NOEXCL)
   
   EXN(NOEXCL) = MINORN(K)

   GoTo next_j
   
   End If

Open fnm For Input As #1

ESTFLAG = "Y"

Dim estdes, popn, smpn, BAC_ACCUR, FRNO, actno, cal, eact, _
          esmp, esites, edays, bac, bac_cvs, bac_cvsp, _
          bac_cvt, bac_cvtp, bac_cv, bac_low, bac_upper, eff, _
          eff_low, eff_upper, nland, CPUE_ACCUR, LPOP, ltot, _
          lsmpv, lsmpf, leff, cpue, lsites, ldays, _
          cpue_cvs, cpue_cvsp, cpue_cvt, cpue_cvtp, cpue_cv, cpue_low, _
          cpue_upper, catch, catch_low, catch_upper, catch_cv, Value, _
          price, fish, kgfish

Dim IX, XKEY

Dim TYPEREC

Do Until EOF(1)

Input #1, XKEY

XKEY = RTrim(XKEY)
XKEY = "M" + CURMJC + Right(XKEY, 12)

.Seek "=", XKEY

If .NoMatch = True Then GoTo ADD_CASE
If .NoMatch = False Then GoTo REPL_CASE

ADD_CASE:

'===================
'NEW RECORD
'===================

.AddNew

![estkey] = XKEY

TYPEREC = 2

If Mid(XKEY, 14, 4) = "0000" And Mid(XKEY, 8, 4) <> "0000" Then TYPEREC = 1

If TYPEREC = 1 Then

Input #1, estdes, popn, smpn, BAC_ACCUR, FRNO, actno, cal, eact, _
          esmp, esites, edays, bac, bac_cvs, bac_cvsp, _
          bac_cvt, bac_cvtp, bac_cv, bac_low, bac_upper, eff, _
          eff_low, eff_upper, nland, CPUE_ACCUR, LPOP, ltot, _
          lsmpv, lsmpf, leff, cpue, lsites, ldays, _
          cpue_cvs, cpue_cvsp, cpue_cvt, cpue_cvtp, cpue_cv, cpue_low, _
          cpue_upper, catch, catch_low, catch_upper, catch_cv, Value, _
          price, fish, kgfish
          
          ![eff] = eff
          ![catch] = catch
          ![Value] = Value
          ![FRNO] = FRNO
          
          If eff <> 0 Then ![cpue] = catch / eff
          If catch <> 0 Then ![price] = Value / catch
          
          End If

If TYPEREC = 2 Then

Input #1, estdes, eff, cpue, catch, Value, price, kgfish, fish, FRNO

          ![eff] = eff
          ![catch] = catch
          ![Value] = Value
          ![fish] = fish
          ![FRNO] = FRNO
          If eff <> 0 Then ![cpue] = catch / eff
          If catch <> 0 Then ![price] = Value / catch
          If ![fish] <> 0 Then ![kgfish] = ![catch] / ![fish]
          
          End If

.Update

GoTo CONT_READ

REPL_CASE:

'===================
'EXISTING RECORD
'===================

.Edit

TYPEREC = 2

If Mid(XKEY, 14, 4) = "0000" And Mid(XKEY, 8, 4) <> "0000" Then TYPEREC = 1

If TYPEREC = 1 Then

Input #1, estdes, popn, smpn, BAC_ACCUR, FRNO, actno, cal, eact, _
          esmp, esites, edays, bac, bac_cvs, bac_cvsp, _
          bac_cvt, bac_cvtp, bac_cv, bac_low, bac_upper, eff, _
          eff_low, eff_upper, nland, CPUE_ACCUR, LPOP, ltot, _
          lsmpv, lsmpf, leff, cpue, lsites, ldays, _
          cpue_cvs, cpue_cvsp, cpue_cvt, cpue_cvtp, cpue_cv, cpue_low, _
          cpue_upper, catch, catch_low, catch_upper, catch_cv, Value, _
          price, fish, kgfish
          
          ![eff] = ![eff] + eff
          ![catch] = ![catch] + catch
          ![Value] = ![Value] + Value
          ![fish] = ![fish] + fish
          ![FRNO] = ![FRNO] + FRNO
          
          If ![eff] <> 0 Then ![cpue] = ![catch] / ![eff]
          If ![catch] <> 0 Then ![price] = ![Value] / ![catch]
          If ![fish] <> 0 Then ![kgfish] = ![catch] / ![fish]
                    
          End If

If TYPEREC = 2 Then

Input #1, estdes, eff, cpue, catch, Value, price, kgfish, fish, FRNO

          ![eff] = ![eff] + eff
          ![catch] = ![catch] + catch
          ![Value] = ![Value] + Value
          ![FRNO] = ![FRNO] + FRNO
          
          If ![eff] <> 0 Then ![cpue] = ![catch] / ![eff]
          If ![catch] <> 0 Then ![price] = ![Value] / ![catch]

         End If

.Update

CONT_READ:

Loop

Close #1

next_j:

Next J

XKEY = "M" + (CURMJC) + "+B0000+S0000"

.Seek "=", XKEY

If .NoMatch = True Then End

If ![eff] * ![catch] = 0 Then ESTFLAG = "N"

End With

prm_record.Close
prm_database.Close

End Sub
Private Sub CREATE_GTDB()

Dim I, J, K, fnm

Dim dbn, CREC

FileCopy APPROOT + "\ARTBAS\STRUS\ESTOT.MDB", APPROOT + "\ARTBAS\RESULTS\WORK.MDB"

dbn = APPROOT + "\ARTBAS\RESULTS\WORK.MDB"

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(dbn)
Set prm_record = prm_database.OpenRecordset("ESTAB")

With prm_record

.Index = "PRIMARYKEY"

NOEXCL = 0

For J = 1 To NTMN

K = Val(TMNC(J))

fnm = APPROOT + "\ARTBAS\RESULTS\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_MN" + TMNC(J) + "_ESTIM.TXT"

If Dir(fnm) = "" Then

   NOEXCL = NOEXCL + 1
   
   ReDim Preserve EXN(NOEXCL)
   
   EXN(NOEXCL) = MINORN(K)

   GoTo next_j
   
   End If


Open fnm For Input As #1

ESTFLAG = "Y"

Dim estdes, popn, smpn, BAC_ACCUR, FRNO, actno, cal, eact, _
          esmp, esites, edays, bac, bac_cvs, bac_cvsp, _
          bac_cvt, bac_cvtp, bac_cv, bac_low, bac_upper, eff, _
          eff_low, eff_upper, nland, CPUE_ACCUR, LPOP, ltot, _
          lsmpv, lsmpf, leff, cpue, lsites, ldays, _
          cpue_cvs, cpue_cvsp, cpue_cvt, cpue_cvtp, cpue_cv, cpue_low, _
          cpue_upper, catch, catch_low, catch_upper, catch_cv, Value, _
          price, fish, kgfish

Dim IX, XKEY

Dim TYPEREC

Do Until EOF(1)

Input #1, XKEY

XKEY = RTrim(XKEY)
XKEY = "M0000" + Right(XKEY, 12)

.Seek "=", XKEY

If .NoMatch = True Then GoTo ADD_CASE
If .NoMatch = False Then GoTo REPL_CASE

ADD_CASE:

'===================
'NEW RECORD
'===================

.AddNew

![estkey] = XKEY

TYPEREC = 2

If Mid(XKEY, 14, 4) = "0000" And Mid(XKEY, 8, 4) <> "0000" Then TYPEREC = 1

If TYPEREC = 1 Then

Input #1, estdes, popn, smpn, BAC_ACCUR, FRNO, actno, cal, eact, _
          esmp, esites, edays, bac, bac_cvs, bac_cvsp, _
          bac_cvt, bac_cvtp, bac_cv, bac_low, bac_upper, eff, _
          eff_low, eff_upper, nland, CPUE_ACCUR, LPOP, ltot, _
          lsmpv, lsmpf, leff, cpue, lsites, ldays, _
          cpue_cvs, cpue_cvsp, cpue_cvt, cpue_cvtp, cpue_cv, cpue_low, _
          cpue_upper, catch, catch_low, catch_upper, catch_cv, Value, _
          price, fish, kgfish
          
          ![eff] = eff
          ![catch] = catch
          ![Value] = Value
          ![FRNO] = FRNO
          
          If eff <> 0 Then ![cpue] = catch / eff
          If catch <> 0 Then ![price] = Value / catch
          
          End If

If TYPEREC = 2 Then

Input #1, estdes, eff, cpue, catch, Value, price, kgfish, fish, FRNO

          ![eff] = eff
          ![catch] = catch
          ![Value] = Value
          ![fish] = fish
          ![FRNO] = FRNO
          If eff <> 0 Then ![cpue] = catch / eff
          If catch <> 0 Then ![price] = Value / catch
          If ![fish] <> 0 Then ![kgfish] = ![catch] / ![fish]
          
          End If

.Update

GoTo CONT_READ

REPL_CASE:

'===================
'EXISTING RECORD
'===================

.Edit

TYPEREC = 2

If Mid(XKEY, 14, 4) = "0000" And Mid(XKEY, 8, 4) <> "0000" Then TYPEREC = 1

If TYPEREC = 1 Then

Input #1, estdes, popn, smpn, BAC_ACCUR, FRNO, actno, cal, eact, _
          esmp, esites, edays, bac, bac_cvs, bac_cvsp, _
          bac_cvt, bac_cvtp, bac_cv, bac_low, bac_upper, eff, _
          eff_low, eff_upper, nland, CPUE_ACCUR, LPOP, ltot, _
          lsmpv, lsmpf, leff, cpue, lsites, ldays, _
          cpue_cvs, cpue_cvsp, cpue_cvt, cpue_cvtp, cpue_cv, cpue_low, _
          cpue_upper, catch, catch_low, catch_upper, catch_cv, Value, _
          price, fish, kgfish
          
          ![eff] = ![eff] + eff
          ![catch] = ![catch] + catch
          ![Value] = ![Value] + Value
          ![fish] = ![fish] + fish
          ![FRNO] = ![FRNO] + FRNO
          
          If ![eff] <> 0 Then ![cpue] = ![catch] / ![eff]
          If ![catch] <> 0 Then ![price] = ![Value] / ![catch]
          If ![fish] <> 0 Then ![kgfish] = ![catch] / ![fish]
                    
          End If

If TYPEREC = 2 Then

Input #1, estdes, eff, cpue, catch, Value, price, kgfish, fish, FRNO

          ![eff] = ![eff] + eff
          ![catch] = ![catch] + catch
          ![Value] = ![Value] + Value
          ![FRNO] = ![FRNO] + FRNO
          
          If ![eff] <> 0 Then ![cpue] = ![catch] / ![eff]
          If ![catch] <> 0 Then ![price] = ![Value] / ![catch]

         End If

.Update

CONT_READ:

Loop

Close #1

next_j:

Next J

XKEY = "M0000+B0000+S0000"

.Seek "=", XKEY

If .NoMatch = True Then End

If ![eff] * ![catch] = 0 Then ESTFLAG = "N"

End With

prm_record.Close
prm_database.Close

End Sub
Private Sub LOAD_MAJOR()

NTMJ = 0

Dim fnm, XXX, I, J, V, W

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_MAJOR.TXT"

If Dir(fnm) = "" Then Exit Sub

Open fnm For Input As #1

Do Until EOF(1)

Line Input #1, XXX

NTMJ = NTMJ + 1

ReDim Preserve TMJC(1 To NTMJ), TMJN(1 To NTMJ), TMJOK(1 To NTMJ)

TMJC(NTMJ) = Left(XXX, 4)
TMJN(NTMJ) = Mid(XXX, 6, 30)

If optSTD.Value = True Then TMJN(NTMJ) = Mid(XXX, 37, 30)

Loop

Close #1

End Sub
Private Sub LOAD_MINOR()

NTMN = 0

Dim fnm, XXX, I, J, V, W

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_MINOR.TXT"

If Dir(fnm) = "" Then Exit Sub

Open fnm For Input As #1

Do Until EOF(1)

Line Input #1, XXX

NTMN = NTMN + 1

ReDim Preserve TMNC(1 To NTMN), TMNN(1 To NTMN), TMNOK(1 To NTMN)

TMNC(NTMN) = Left(XXX, 4)
TMNN(NTMN) = Mid(XXX, 6, 30)

If optSTD.Value = True Then TMNN(NTMN) = Mid(XXX, 37, 30)

Loop

Close #1

End Sub
Private Sub LOAD_SITES()

NTSI = 0

Dim fnm, XXX, I, J, V, W

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_SITES.TXT"

If Dir(fnm) = "" Then Exit Sub

Open fnm For Input As #1

Do Until EOF(1)

Line Input #1, XXX

NTSI = NTSI + 1

ReDim Preserve TSIC(1 To NTSI), TSIN(1 To NTSI)

TSIC(NTSI) = Left(XXX, 4)
TSIN(NTSI) = Mid(XXX, 6, 30)

If optSTD.Value = True Then TSIN(NTSI) = Mid(XXX, 37, 30)

Loop

Close #1

End Sub
Private Sub LOAD_SPECIES()

NTSP = 0

Dim fnm, XXX, I, J, V, W

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_SPECIES.TXT"

If Dir(fnm) = "" Then Exit Sub

Open fnm For Input As #1

Do Until EOF(1)

Line Input #1, XXX

NTSP = NTSP + 1

ReDim Preserve TSPC(1 To NTSP), TSPN(1 To NTSP)

TSPC(NTSP) = Left(XXX, 4)
TSPN(NTSP) = Mid(XXX, 6, 30)

If optSTD.Value = True Then TSPN(NTSP) = Mid(XXX, 37, 30)

Loop

Close #1

End Sub
Private Sub LOAD_BG()

NTBG = 0

Dim fnm, XXX, I, J, V, W

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_BG.TXT"

If Dir(fnm) = "" Then Exit Sub

Open fnm For Input As #1

Do Until EOF(1)

Line Input #1, XXX

NTBG = NTBG + 1

ReDim Preserve TBGC(1 To NTBG), TBGN(1 To NTBG)

TBGC(NTBG) = Left(XXX, 4)
TBGN(NTBG) = Mid(XXX, 6, 30)

If optSTD.Value = True Then TBGN(NTBG) = Mid(XXX, 37, 30)

Loop

Close #1

End Sub
Private Sub LOAD_ASSOMN()

NASSOMN = 0

ReDim ASSOMN(1 To 10000)

Dim fnm, XXX, I, J, V, W, yyy, K, m

For I = 1 To 10000
ASSOMN(I) = 0
Next I

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_ASSOMN.TXT"

If Dir(fnm) = "" Then Exit Sub

NASSOMN = 1

Open fnm For Input As #1

Do Until EOF(1)

Line Input #1, XXX

J = Val(Mid(XXX, 37, 4)): m = Val(Left(XXX, 4))

For I = 1 To J

Line Input #1, yyy

K = Val(Mid(yyy, 6, 4))

ASSOMN(K) = m

Next I

Loop

Close #1

End Sub
Private Sub LOAD_ASSOSI()

NASSOSI = 0

ReDim ASSOSI(1 To 10000)

Dim fnm, XXX, I, J, V, W, yyy, K, m

For I = 1 To 10000
ASSOSI(I) = 0
Next I

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_ASSOSI.TXT"

If Dir(fnm) = "" Then Exit Sub

NASSOSI = 1

Open fnm For Input As #1

Do Until EOF(1)

Line Input #1, XXX

J = Val(Mid(XXX, 37, 4)): m = Val(Left(XXX, 4))

For I = 1 To J

Line Input #1, yyy

K = Val(Mid(yyy, 6, 4))

ASSOSI(K) = m

Next I

Loop

Close #1

End Sub
Private Sub LIST_MINOR()

Dim fnm, I

For I = 1 To NTMN

TMNOK(I) = "-"

fnm = APPROOT + "\ARTBAS\RESULTS\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_MN" + TMNC(I) + "_ESTIM.TXT"

If Dir(fnm) <> "" Then TMNOK(I) = "+"

Next I

lstLIST.Clear

For I = 1 To NTMN

lstLIST.AddItem TMNN(I) + " " + TMNOK(I)

Next I

lstLIST.Visible = True

End Sub
Private Sub LIST_MAJOR()

Dim fnm, I, J, K

For I = 1 To NTMJ

TMJOK(I) = "-"

For J = 1 To NTMN

K = Val(TMNC(J))

If ASSOMN(K) <> Val(TMJC(I)) Then GoTo next_j

fnm = APPROOT + "\ARTBAS\RESULTS\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_MN" + TMNC(J) + "_ESTIM.TXT"

If Dir(fnm) <> "" Then TMJOK(I) = "+"

next_j:

Next J

Next I

lstLIST.Clear

For I = 1 To NTMJ

lstLIST.AddItem TMJN(I) + " " + TMJOK(I)

Next I

lstLIST.Visible = True

End Sub
Private Sub BYMN_BYBG_BYSP()

Dim ESTFNM

ESTFNM = APPROOT + "\ARTBAS\CURRENT_TABLES\ESTIM.TXT"

Open ESTFNM For Output As #3

Dim dbn, I, J, m, K, XKEY, LSTR

LSTR = 91

dbn = APPROOT + "\ARTBAS\RESULTS\WORK.MDB"

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(dbn)
Set prm_record = prm_database.OpenRecordset("ESTAB")

With prm_record

.Index = "primarykey"

Open APPROOT + "\ARTBAS\RESULTS\WORK.TXT" For Output As #1

Print #1, Tab(5); frmREP.Caption + " ( " + msgtab(37) + " )"

Print #1, " "

Print #1, Tab(5); msgtab(228) + " : " + RTrim(MINORN(Val(CURMNC)))
Print #1, Tab(5); msgtab(227) + " : " + optMN1.Caption

Print #1, Tab(5); String(LSTR, "=")

Print #1, Tab(5); msgtab(191)

Print #1, " "

Print #1, Tab(5); msgtab(192) + " " + UNW
Print #1, Tab(5); msgtab(193) + " " + msgtab(198)
Print #1, Tab(5); msgtab(200) + " " + UNW + "/" + msgtab(198)
Print #1, Tab(5); msgtab(194) + " " + UNM + "/" + UNW
Print #1, Tab(5); msgtab(195) + " " + UNM
Print #1, Tab(5); msgtab(196) + " " + UNW + "/" + msgtab(197)
Print #1, Tab(5); msgtab(260)

Print #1, " "

'EXPORTING START
'===============

Write #3, " ", " ", " ", " ", " ", " ", " ", " "
Write #3, frmREP.Caption + " ( " + msgtab(37) + " )", _
        " ", " ", " ", " ", " ", " ", " "
Write #3, " ", " ", " ", " ", " ", " ", " ", " "

Write #3, msgtab(191), " ", " ", " ", " ", " ", " ", " "
Write #3, "  ", " ", " ", " ", " ", " ", " ", " "
Write #3, msgtab(192) + " " + UNW, " ", " ", " ", " ", " ", " ", " "
Write #3, msgtab(193) + " " + msgtab(198), " ", " ", " ", " ", " ", " ", " "
Write #3, msgtab(200) + " " + UNW + "/" + msgtab(198), " ", " ", " ", " ", " ", " ", " "
Write #3, msgtab(194) + " " + UNM + "/" + UNW, " ", " ", " ", " ", " ", " ", " "
Write #3, msgtab(195) + " " + UNM, " ", " ", " ", " ", " ", " ", " "
Write #3, msgtab(196) + " " + UNW + "/" + msgtab(197), " ", " ", " ", " ", " ", " ", " "
Write #3, msgtab(260), " ", " ", " ", " ", " ", " ", " "
Write #3, "  ", " ", " ", " ", " ", " ", " ", " "

Write #3, "--- " + msgtab(228) + " : " + RTrim(MINORN(Val(CURMNC))) + " ---", _
          " ", " ", " ", " ", " ", " ", " "
Write #3, "--- " + msgtab(227) + " : " + RTrim(optMN1.Caption) + " ---" _
          , " ", " ", " ", " ", " ", " ", " "
Write #3, " ", " ", " ", " ", " ", " ", " ", " "
'Write #3, " ", RTrim(msgtab(67)), RTrim(msgtab(33)), RTrim(msgtab(209)), RTrim(msgtab(173)), _
          RTrim(msgtab(94)), RTrim(msgtab(95))
Write #3, " ", " ", " ", " ", " ", " ", " ", " "

'=========================

K = Val(CURMNC)

For I = 1 To NTBG

XKEY = "M" + CURMNC + "+B" + TBGC(I) + "+S0000"

.Seek "=", XKEY

If .NoMatch = True Then GoTo NEXT_I

If ![catch] * ![eff] = 0 Then GoTo NEXT_I

Print #1, Tab(5); RTrim(MINORN(K)) + " : " + LTrim(RTrim(TBGN(I)))
Write #3, "--- " + LTrim(RTrim(TBGN(I))) + " ---" _
        , " ", " ", " ", " ", " ", " ", " "
         
Write #3, "  ", " ", " ", " ", " ", " ", " ", " "

Print #1, Tab(5); String(LSTR, "=")

Print #1, Tab(5); msgtab(152)
Write #3, msgtab(152), " ", " ", " ", " ", " ", " ", " "

Print #1, " "
Write #3, "  ", " ", " ", " ", " ", " ", " ", " "

Dim PPP, DECF1, DECF2, GTC, GTE, GTV

DECF1 = "### ### ### ##0"
DECF2 = "### ##0.000"

GTC = ![catch]: GTE = ![eff]: GTV = ![Value]

Dim ss, cc, NN, MM, PRST

ss = 0.5: NN = ![popn]

cc = (0.1 / (1.96 * ss)) ^ 2 * NN

MM = Int(NN / (1 + cc))

PPP = Left(msgtab(163) + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(![FRNO], DECF1))
Write #3, RTrim(msgtab(163)), ![FRNO], " ", " ", " ", " ", " ", " "
          

PPP = Left(msgtab(164) + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(![actno], "###0.0"))
Write #3, RTrim(msgtab(164)), ![actno], " ", " ", " ", " ", " ", " "

PPP = Left(msgtab(155) + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(100 * ![bac], "##0.00")) + " %"
Write #3, RTrim(msgtab(155)) + " (%)", 100 * ![bac], _
          " ", " ", " ", " ", " ", " "
           
          
'========================================================
' Calculate accuracy

CONVEX_YN = "N"
POPSIZE = ![FRNO] * ![actno]
POPSIZE = Int(POPSIZE)
INSMP = Int(![smpn])
      
Call ACCURACY_FOR_GIVEN_SAMPLES
    
If SPST_IND <> " " Then SPST_IND = "(***)"
'=========================================================

PPP = Left(RTrim(msgtab(156)) + " " + SPST_IND + String(40, "."), 40) + " " + _
      LTrim(Format(100 * OUTACC, "##0.000")) + " % "
      
      
Print #1, Tab(5); PPP
Write #3, RTrim(msgtab(156)) + " (%) " + SPST_IND, 100 * OUTACC, _
          " ", " ", " ", " ", " ", " "

PPP = Left(RTrim(msgtab(259)) + String(40, "."), 40) + " " + _
      LTrim(Format(POPSIZE, "#########0"))
Print #1, Tab(5); PPP
Write #3, RTrim(msgtab(259)), Int(POPSIZE), " ", " ", " ", " ", " ", " "

PPP = Left(msgtab(153) + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(![smpn], DECF1))
Write #3, RTrim(msgtab(153)), ![smpn], _
          " ", " ", " ", " ", " ", " "
 
PPP = Left(msgtab(154) + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(![eact], DECF1))
Write #3, RTrim(msgtab(154)), Int(![eact]), _
          " ", " ", " ", " ", " ", " "

Print #1, " "
Write #3, "  ", " ", " ", " ", " ", " ", " ", " "

PPP = Left(msgtab(204) + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(![esites], DECF1))
Write #3, RTrim(msgtab(204)), ![esites], _
 " ", " ", " ", " ", " ", " "

PPP = Left(msgtab(205) + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(![edays], DECF1))
Write #3, RTrim(msgtab(205)), ![edays], _
" ", " ", " ", " ", " ", " "

PPP = Left(msgtab(158) + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(100 * ![bac_cv], DECF2)) + " %"
Write #3, RTrim(msgtab(158)) + " (%)", 100 * ![bac_cv], _
" ", " ", " ", " ", " ", " "

PRST = ![bac_cvsp] * ![bac_cv]

If ![esites] >= 2 And ![edays] >= 2 Then
PPP = Left(msgtab(159) + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(PRST, DECF2)) + " %"
Write #3, RTrim(msgtab(159)) + " (%)", PRST, _
" ", " ", " ", " ", " ", " "
End If

If ![esites] < 2 Or ![edays] < 2 Then
PPP = Left(msgtab(159) + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + msgtab(168)
Write #3, RTrim(msgtab(159)), RTrim(msgtab(168)) _
; " ", " ", " ", " ", " ", " "
End If

PRST = ![bac_cvtp] * ![bac_cv]

If ![edays] >= 2 And ![esites] >= 2 Then
PPP = Left(msgtab(160) + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(PRST, DECF2)) + " %"
Write #3, RTrim(msgtab(160)) + " (%)", PRST, _
" ", " ", " ", " ", " ", " "
End If

If ![edays] < 2 Or ![esites] < 2 Then
PPP = Left(msgtab(160) + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + msgtab(168)
Write #3, RTrim(msgtab(160)), RTrim(msgtab(168)), _
" ", " ", " ", " ", " ", " "
End If

PPP = Left(msgtab(161) + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(100 * ![bac_low], DECF2)) + " %"
Write #3, RTrim(msgtab(161)), 100 * ![bac_low], _
" ", " ", " ", " ", " ", " "

PPP = Left(msgtab(162) + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(100 * ![bac_upper], DECF2)) + " %"
Write #3, RTrim(msgtab(162)), 100 * ![bac_upper], _
" ", " ", " ", " ", " ", " "

Print #1, " "
Write #3, "  ", " ", " ", " ", " ", " ", " ", " "

PPP = Left(msgtab(165) + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(![eff], DECF1))
Write #3, RTrim(msgtab(165)), ![eff], _
" ", " ", " ", " ", " ", " "

PPP = Left(msgtab(166) + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(![eff_low], DECF1))
Write #3, RTrim(msgtab(166)), ![eff_low], _
" ", " ", " ", " ", " ", " "

PPP = Left(msgtab(167) + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(![eff_upper], DECF1))
Write #3, RTrim(msgtab(167)), ![eff_upper], _
" ", " ", " ", " ", " ", " "

Print #1, " "
Write #3, "  ", " ", " ", " ", " ", " ", " ", " "

Print #1, Tab(5); msgtab(170)
Write #3, RTrim(msgtab(170)), " ", " ", " ", " ", " ", " ", " "

Print #1, " "
Write #3, "  ", " ", " ", " ", " ", " ", " ", " "

NN = ![LPOP]

ss = (2 * NN - 1) / (6 * NN - 6) - 0.25

ss = ss ^ 0.5

cc = (0.1 / (1.96 * ss)) ^ 2 * NN

MM = Int(NN / (1 + cc))

PPP = Left(msgtab(173) + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(![cpue], DECF2))
Write #3, RTrim(msgtab(173)), ![cpue], _
" ", " ", " ", " ", " ", " "

'========================================================
' Calculate accuracy

CONVEX_YN = "Y"
POPSIZE = ![FRNO] * ![actno] * ![bac]
POPSIZE = Int(POPSIZE)
INSMP = Int(![nland])
      
Call ACCURACY_FOR_GIVEN_SAMPLES
    
If SPST_IND <> " " Then SPST_IND = "(***)"
'=========================================================

PPP = Left(msgtab(156) + " " + SPST_IND + String(40, "."), 40) + " " + _
      LTrim(Format(100 * OUTACC, "###0.000")) + " % "

Print #1, Tab(5); PPP
Write #3, RTrim(msgtab(156)) + " (%)" + " " + SPST_IND, _
          100 * OUTACC, " ", " ", " ", " ", " ", " "

PPP = Left(RTrim(msgtab(258)) + String(40, "."), 40) + " " + _
      LTrim(Format(POPSIZE, "#########0"))
Print #1, Tab(5); PPP
Write #3, RTrim(msgtab(258)), POPSIZE, " ", " ", " ", " ", " ", " "

PPP = Left(msgtab(206) + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(![nland], DECF1))
Write #3, RTrim(msgtab(206)), ![nland], " ", " ", " ", " ", " ", " "

PPP = Left(msgtab(171) + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(![ltot], DECF1))
Write #3, RTrim(msgtab(171)), ![ltot], " ", " ", " ", " ", " ", " "

PPP = Left(msgtab(172) + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(![leff], DECF2))
Write #3, RTrim(msgtab(172)), ![leff], " ", " ", " ", " ", " ", " "

Print #1, " "
Write #3, "  ", " ", " ", " ", " ", " ", " ", " "

PPP = Left(msgtab(204) + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(![lsites], DECF1))
Write #3, RTrim(msgtab(204)), ![lsites], " ", " ", " ", " ", " ", " "

PPP = Left(msgtab(205) + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(![ldays], DECF1))
Write #3, RTrim(msgtab(205)), ![ldays], " ", " ", " ", " ", " ", " "

PPP = Left(msgtab(176) + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(100 * ![cpue_cv], DECF2)) + " %"
Write #3, RTrim(msgtab(176)) + " (%)", 100 * ![cpue_cv], _
" ", " ", " ", " ", " ", " "

PRST = ![cpue_cvsp] * ![cpue_cv]

If ![lsites] >= 2 And ![ldays] >= 2 Then
PPP = Left(msgtab(177) + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(PRST, DECF2)) + " %"
Write #3, RTrim(msgtab(177)) + " (%)", PRST, _
" ", " ", " ", " ", " ", " "
End If

If ![lsites] < 2 Or ![ldays] < 2 Then
PPP = Left(msgtab(177) + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + msgtab(168)
Write #3, RTrim(msgtab(177)), RTrim(msgtab(168)), _
 " ", " ", " ", " ", " ", " "
End If

PRST = ![cpue_cvtp] * ![cpue_cv]

If ![ldays] >= 2 And ![lsites] >= 2 Then
PPP = Left(msgtab(178) + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(PRST, DECF2)) + " %"
Write #3, RTrim(msgtab(178)) + " (%)", PRST, _
" ", " ", " ", " ", " ", " "
End If

If ![ldays] < 2 Or ![lsites] < 2 Then
PPP = Left(msgtab(178) + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + msgtab(168)
Write #3, RTrim(msgtab(178)), RTrim(msgtab(168)), " ", " ", " ", " ", " ", " "
End If

PPP = Left(msgtab(179) + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(![cpue_low], DECF2))
Write #3, RTrim(msgtab(179)), ![cpue_low], _
" ", " ", " ", " ", " ", " "

PPP = Left(msgtab(180) + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(![cpue_upper], DECF2))
Write #3, RTrim(msgtab(180)), ![cpue_upper], _
" ", " ", " ", " ", " ", " "

Print #1, " "
Write #3, "  ", " ", " ", " ", " ", " ", " ", " "

PPP = Left(msgtab(181) + " (" + UNW + ") " + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(![catch], DECF1))
Write #3, RTrim(msgtab(181)), ![catch], _
" ", " ", " ", " ", " ", " "

PPP = Left(msgtab(190) + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(100 * ![catch_cv], DECF2)) + " %"
Write #3, RTrim(msgtab(190)) + " (%)", 100 * ![catch_cv], _
" ", " ", " ", " ", " ", " "

PPP = Left(msgtab(182) + " (" + UNW + ") " + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(![catch_low], DECF1))
Write #3, RTrim(msgtab(182)), ![catch_low], _
" ", " ", " ", " ", " ", " "

PPP = Left(msgtab(183) + " (" + UNW + ") " + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(![catch_upper], DECF1))
Write #3, RTrim(msgtab(183)), ![catch_upper], _
" ", " ", " ", " ", " ", " "

Print #1, " "
Write #3, "  ", " ", " ", " ", " ", " ", " ", " "

PPP = Left(msgtab(184) + " (" + UNM + ") " + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(![Value], DECF1))
Write #3, RTrim(msgtab(184)), ![Value], _
" ", " ", " ", " ", " ", " "

PPP = Left(msgtab(185) + " (" + UNM + "/" + UNW + ") " + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(![price], DECF2))
Write #3, RTrim(msgtab(185)), ![price], _
" ", " ", " ", " ", " ", " "

Print #1, " "
Write #3, "  ", " ", " ", " ", " ", " ", " ", " "

Print #1, Tab(200); "<"

PPP = Left(msgtab(186) + Space(30), 30) + " "
PPP = PPP + Right(Space(15) + RTrim(msgtab(97)), 15) + Space(10)
PPP = PPP + Right(Space(11) + RTrim(msgtab(173)), 11) + " "
PPP = PPP + Right(Space(15) + RTrim(msgtab(95)), 15)

Print #1, Tab(5); PPP
Write #3, RTrim(msgtab(186)), RTrim(msgtab(97)); RTrim(msgtab(33)), _
             RTrim(msgtab(173)), RTrim(msgtab(199)), _
             RTrim(msgtab(94)), RTrim(msgtab(95)), " "

PPP = Space(31)
PPP = PPP + Right(Space(15) + RTrim(msgtab(33)), 15) + Space(10)
PPP = PPP + Right(Space(11) + RTrim(msgtab(199)), 11) + " "
PPP = PPP + Right(Space(15) + RTrim(msgtab(94)), 15)

Print #1, Tab(5); PPP

Print #1, " "
Write #3, "  ", " ", " ", " ", " ", " ", " ", " "

For J = 1 To NTSP

XKEY = "M" + CURMNC + "+B" + TBGC(I) + "+S" + TSPC(J)

.Seek "=", XKEY

If .NoMatch = True Then GoTo next_j

Dim PGTC, PGTE, PGTV

PGTC = 0
If GTC <> 0 Then PGTC = 100 * ![catch] / GTC

PGTE = 0
If GTE <> 0 Then PGTE = 100 * ![eff] / GTE

PGTV = 0
If GTV <> 0 Then PGTV = 100 * ![Value] / GTV

PPP = TSPN(J) + " "
PPP = PPP + Right(Space(15) + LTrim(Format(![catch], DECF1)), 15) + _
      " (" + Right(Space(5) + Format(PGTC, "##0.0"), 5) + "%) "
PPP = PPP + Right(Space(11) + LTrim(Format(![cpue], DECF2)), 11) + " "
PPP = PPP + Right(Space(15) + LTrim(Format(![Value], DECF1)), 15) + _
      " (" + Right(Space(5) + Format(PGTV, "##0.0"), 5) + "%) "

Print #1, Tab(5); PPP

ZC = ![catch]
ZE = ![eff]
ZU = ![cpue]
ZW = ![kgfish]
ZP = ![price]
ZV = ![Value]

If ZW = 0 Then ZW = "..."
If ZP = 0 Then ZP = "..."
If ZV = 0 Then ZV = "..."

Write #3, LTrim(RTrim(TSPN(J))), ZC, ZE, ZU, ZW, ZP, ZV, " "

PPP = Space(31) + Right(Space(15) + LTrim(Format(![eff], DECF1)), 15) + _
      Space(10)
PPP = PPP + Right(Space(11) + LTrim(Format(![kgfish], DECF2)), 11) + " "
PPP = PPP + Right(Space(15) + LTrim(Format(![price], DECF2)), 15) + " "

Print #1, Tab(5); PPP

Print #1, " "

next_j:

Next J

Write #3, " ", " ", " ", " ", " ", " ", " ", " "

Print #1, " "
Print #1, Tab(200); "<"

NEXT_I:

Next I

Close #1

End With

Close #3

prm_record.Close
prm_database.Close

rtsDISP.FileName = APPROOT + "\ARTBAS\RESULTS\WORK.TXT"
rtsDISP.Visible = True
lblEXP.Visible = False
lstLIST.Visible = False
Frame1.Visible = False

cmdEXCEL_ESTIM.Visible = True

End Sub
Private Sub BYMJ_BYBG_BYSP()

Dim ESTFNM

ESTFNM = APPROOT + "\ARTBAS\CURRENT_TABLES\ESTIM.TXT"

Open ESTFNM For Output As #3

Dim dbn, I, J, m, K, XKEY, LSTR

LSTR = 91

dbn = APPROOT + "\ARTBAS\RESULTS\WORK.MDB"

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(dbn)
Set prm_record = prm_database.OpenRecordset("ESTAB")

With prm_record

.Index = "primarykey"

Open APPROOT + "\ARTBAS\RESULTS\WORK.TXT" For Output As #1

Print #1, Tab(5); frmREP.Caption + " ( " + msgtab(37) + " )"

Print #1, " "

'EXPORTING START
'===============

Write #3, " ", " ", " ", " ", " ", " ", " ", " "
Write #3, frmREP.Caption + " ( " + msgtab(37) + " )", " ", " ", " ", " ", " ", " ", " "
Write #3, " ", " ", " ", " ", " ", " ", " ", " "

Write #3, msgtab(191), " ", " ", " ", " ", " ", " ", " "
Write #3, "  ", " ", " ", " ", " ", " ", " ", " "
Write #3, msgtab(192) + " " + UNW, " ", " ", " ", " ", " ", " ", " "
Write #3, msgtab(193) + " " + msgtab(198), " ", " ", " ", " ", " ", " ", " "
Write #3, msgtab(200) + " " + UNW + "/" + msgtab(198), " ", " ", " ", " ", " ", " ", " "
Write #3, msgtab(194) + " " + UNM + "/" + UNW, " ", " ", " ", " ", " ", " ", " "
Write #3, msgtab(195) + " " + UNM, " ", " ", " ", " ", " ", " ", " "
Write #3, msgtab(196) + " " + UNW + "/" + msgtab(197), " ", " ", " ", " ", " ", " ", " "
Write #3, "  ", " ", " ", " ", " ", " ", " ", " "

'=========================

Print #1, Tab(5); msgtab(229) + " : " + RTrim(MAJORN(Val(CURMJC)))
Print #1, Tab(5); msgtab(227) + " : " + optMJ1.Caption

Call NOT_INCLUDED

Write #3, "  ", " ", " ", " ", " ", " ", " ", " "
Write #3, "--- " + RTrim(MAJORN(Val(CURMJC))) + " ---", " ", " ", " ", " ", " ", " ", " "
Write #3, "--- " + msgtab(227) + " : " + optMJ1.Caption + " ---", " ", " ", " ", " ", " ", " ", " "
Write #3, " ", " ", " ", " ", " ", " ", " ", " "
Write #3, " ", RTrim(msgtab(67)), RTrim(msgtab(33)), RTrim(msgtab(209)), RTrim(msgtab(173)), _
          RTrim(msgtab(94)), RTrim(msgtab(95)), " "
Write #3, " ", " ", " ", " ", " ", " ", " ", " "

Print #1, Tab(5); String(LSTR, "=")

Print #1, Tab(5); msgtab(191)

Print #1, " "

Print #1, Tab(5); msgtab(192) + " " + UNW
Print #1, Tab(5); msgtab(193) + " " + msgtab(198)
Print #1, Tab(5); msgtab(200) + " " + UNW + "/" + msgtab(198)
Print #1, Tab(5); msgtab(194) + " " + UNM + "/" + UNW
Print #1, Tab(5); msgtab(195) + " " + UNM
Print #1, Tab(5); msgtab(196) + " " + UNW + "/" + msgtab(197)

Print #1, " "

K = Val(CURMJC)

For I = 1 To NTBG

XKEY = "M" + CURMJC + "+B" + TBGC(I) + "+S0000"

.Seek "=", XKEY

If .NoMatch = True Then GoTo NEXT_I

Print #1, Tab(5); RTrim(MAJORN(K)) + " : " + TBGN(I)

Print #1, Tab(5); String(LSTR, "=")

Dim PPP, DECF1, DECF2, GTC, GTE, GTV

DECF1 = "### ### ### ##0"
DECF2 = "### ##0.000"

GTC = ![catch]: GTE = ![eff]: GTV = ![Value]

PPP = Left(msgtab(163) + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(![FRNO], DECF1))

PPP = Left(msgtab(165) + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(![eff], DECF1))

PPP = Left(msgtab(173) + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(![cpue], DECF2))

PPP = Left(msgtab(181) + " (" + UNW + ") " + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(![catch], DECF1))

PPP = Left(msgtab(184) + " (" + UNM + ") " + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(![Value], DECF1))

PPP = Left(msgtab(185) + " (" + UNM + "/" + UNW + ") " + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(![price], DECF2))

Print #1, " "

PPP = Left(msgtab(186) + Space(30), 30) + " "
PPP = PPP + Right(Space(15) + RTrim(msgtab(97)), 15) + Space(10)
PPP = PPP + Right(Space(11) + RTrim(msgtab(173)), 11) + " "
PPP = PPP + Right(Space(15) + RTrim(msgtab(95)), 15)

Print #1, Tab(5); PPP

PPP = Space(31)
PPP = PPP + Right(Space(15) + RTrim(msgtab(33)), 15) + Space(10)
PPP = PPP + Right(Space(11) + RTrim(msgtab(199)), 11) + " "
PPP = PPP + Right(Space(15) + RTrim(msgtab(94)), 15)

Print #1, Tab(5); PPP

Print #1, " "

'EXPORTING
'=========

ZC = ![catch]: ZE = ![eff]: ZU = ![cpue]: ZP = ![price]: ZV = ![Value]
ZF = (![FRNO])

If ZP = 0 Then
   ZP = "...": ZV = "..."
   End If

Write #3, " ", " ", " ", " ", " ", " ", " ", " "
Write #3, RTrim(TBGN(I)), ZF, ZE, ZC, ZU, ZP, ZV, " "
Write #3, " ", " ", " ", " ", " ", " ", " ", " "

'================

For J = 1 To NTSP

XKEY = "M" + CURMJC + "+B" + TBGC(I) + "+S" + TSPC(J)

.Seek "=", XKEY

If .NoMatch = True Then GoTo next_j

Dim PGTC, PGTE, PGTV

PGTC = 0
If GTC <> 0 Then PGTC = 100 * ![catch] / GTC

PGTE = 0
If GTE <> 0 Then PGTE = 100 * ![eff] / GTE

PGTV = 0
If GTV <> 0 Then PGTV = 100 * ![Value] / GTV

PPP = TSPN(J) + " "
PPP = PPP + Right(Space(15) + LTrim(Format(![catch], DECF1)), 15) + _
      " (" + Right(Space(5) + Format(PGTC, "##0.0"), 5) + "%) "
PPP = PPP + Right(Space(11) + LTrim(Format(![cpue], DECF2)), 11) + " "
PPP = PPP + Right(Space(15) + LTrim(Format(![Value], DECF1)), 15) + _
      " (" + Right(Space(5) + Format(PGTV, "##0.0"), 5) + "%) "
Print #1, Tab(5); PPP

PPP = Space(31) + Right(Space(15) + LTrim(Format(![eff], DECF1)), 15) + _
      Space(10)
PPP = PPP + Right(Space(11) + LTrim(Format(![kgfish], DECF2)), 11) + " "
PPP = PPP + Right(Space(15) + LTrim(Format(![price], DECF2)), 15) + " "

Print #1, Tab(5); PPP

Print #1, " "

'EXPORTING
'=========

ZC = ![catch]: ZE = ![eff]: ZU = ![cpue]: ZP = ![price]: ZV = ![Value]

If ZP = 0 Then
   ZP = "...": ZV = "..."
   End If

Write #3, RTrim(TSPN(J)), " ", ZE, ZC, ZU, ZP, ZV, " "

next_j:

Next J

Print #1, " "

NEXT_I:

Next I

Close #1

End With

prm_record.Close
prm_database.Close

rtsDISP.FileName = APPROOT + "\ARTBAS\RESULTS\WORK.TXT"
rtsDISP.Visible = True
lblEXP.Visible = False
lstLIST.Visible = False
Frame1.Visible = False

Close #3

End Sub
Private Sub BYGT_BYBG_BYSP()

Dim ESTFNM

ESTFNM = APPROOT + "\ARTBAS\CURRENT_TABLES\ESTIM.TXT"

Open ESTFNM For Output As #3

Dim dbn, I, J, m, K, XKEY, LSTR

LSTR = 91

dbn = APPROOT + "\ARTBAS\RESULTS\WORK.MDB"

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(dbn)
Set prm_record = prm_database.OpenRecordset("ESTAB")

With prm_record

.Index = "primarykey"

Open APPROOT + "\ARTBAS\RESULTS\WORK.TXT" For Output As #1

Print #1, Tab(5); frmREP.Caption + " ( " + msgtab(37) + " : " + msgtab(143) + " )"
Write #3, frmREP.Caption + " ( " + msgtab(37) + " : " + msgtab(143) + " )", _
" ", " ", " ", " ", " ", " ", " "

Write #3, " ", " ", " ", " ", " ", " ", " ", " "
Write #3, msgtab(227) + " : " + optGT1.Caption, " ", " ", " ", " ", " ", " ", " "

Write #3, " ", " ", " ", " ", " ", " ", " ", " "
Write #3, msgtab(191), " ", " ", " ", " ", " ", " ", " "
Write #3, "  ", " ", " ", " ", " ", " ", " ", " "
Write #3, msgtab(192) + " " + UNW, " ", " ", " ", " ", " ", " ", " "
Write #3, msgtab(193) + " " + msgtab(198), " ", " ", " ", " ", " ", " ", " "
Write #3, msgtab(200) + " " + UNW + "/" + msgtab(198), " ", " ", " ", " ", " ", " ", " "
Write #3, msgtab(194) + " " + UNM + "/" + UNW, " ", " ", " ", " ", " ", " ", " "
Write #3, msgtab(195) + " " + UNM, " ", " ", " ", " ", " ", " ", " "
Write #3, msgtab(196) + " " + UNW + "/" + msgtab(197), " ", " ", " ", " ", " ", " ", " "
Write #3, " ", " ", " ", " ", " ", " ", " ", " "

Print #1, " "

Print #1, Tab(5); msgtab(227) + " : " + optGT1.Caption

Call NOT_INCLUDED

Print #1, Tab(5); String(LSTR, "=")
Write #3, " ", " ", " ", " ", " ", " ", " ", " "

Print #1, Tab(5); msgtab(191)

Print #1, " "

'EXPORTING START
'===============

Write #3, " ", RTrim(msgtab(67)), RTrim(msgtab(33)), RTrim(msgtab(209)), RTrim(msgtab(173)), _
          RTrim(msgtab(94)), RTrim(msgtab(95)), " "
Write #3, " ", " ", " ", " ", " ", " ", " ", " "

'=========================

Print #1, Tab(5); msgtab(192) + " " + UNW
Print #1, Tab(5); msgtab(193) + " " + msgtab(198)
Print #1, Tab(5); msgtab(200) + " " + UNW + "/" + msgtab(198)
Print #1, Tab(5); msgtab(194) + " " + UNM + "/" + UNW
Print #1, Tab(5); msgtab(195) + " " + UNM
Print #1, Tab(5); msgtab(196) + " " + UNW + "/" + msgtab(197)

Print #1, " "

For I = 1 To NTBG

XKEY = "M0000" + "+B" + TBGC(I) + "+S0000"

.Seek "=", XKEY

If .NoMatch = True Then GoTo NEXT_I

Print #1, Tab(5); RTrim(msgtab(143)) + " : " + TBGN(I)


ZC = ![catch]: ZE = ![eff]: ZU = ![cpue]: ZP = ![price]: ZV = ![Value]
ZF = (![FRNO])

If ZP = 0 Then
   ZP = "...": ZV = "..."
   End If

Write #3, RTrim(TBGN(I)), ZF, ZE, ZC, ZU, ZP, ZV, " "
Write #3, " ", " ", " ", " ", " ", " ", " ", " "

Print #1, Tab(5); String(LSTR, "=")

Dim PPP, DECF1, DECF2, GTC, GTE, GTV

DECF1 = "### ### ### ##0"
DECF2 = "### ##0.000"

GTC = ![catch]: GTE = ![eff]: GTV = ![Value]

PPP = Left(msgtab(163) + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(![FRNO], DECF1))

PPP = Left(msgtab(165) + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(![eff], DECF1))

PPP = Left(msgtab(173) + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(![cpue], DECF2))

PPP = Left(msgtab(181) + " (" + UNW + ") " + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(![catch], DECF1))

PPP = Left(msgtab(184) + " (" + UNM + ") " + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(![Value], DECF1))

PPP = Left(msgtab(185) + " (" + UNM + "/" + UNW + ") " + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(![price], DECF2))

Print #1, " "

PPP = Left(msgtab(186) + Space(30), 30) + " "
PPP = PPP + Right(Space(15) + RTrim(msgtab(97)), 15) + Space(10)
PPP = PPP + Right(Space(11) + RTrim(msgtab(173)), 11) + " "
PPP = PPP + Right(Space(15) + RTrim(msgtab(95)), 15)

Print #1, Tab(5); PPP

PPP = Space(31)
PPP = PPP + Right(Space(15) + RTrim(msgtab(33)), 15) + Space(10)
PPP = PPP + Right(Space(11) + RTrim(msgtab(199)), 11) + " "
PPP = PPP + Right(Space(15) + RTrim(msgtab(94)), 15)

Print #1, Tab(5); PPP

Print #1, " "

'EXPORTING
'=========


For J = 1 To NTSP

XKEY = "M0000" + "+B" + TBGC(I) + "+S" + TSPC(J)

.Seek "=", XKEY

If .NoMatch = True Then GoTo next_j

Dim PGTC, PGTE, PGTV

PGTC = 0
If GTC <> 0 Then PGTC = 100 * ![catch] / GTC

PGTE = 0
If GTE <> 0 Then PGTE = 100 * ![eff] / GTE

PGTV = 0
If GTV <> 0 Then PGTV = 100 * ![Value] / GTV

PPP = TSPN(J) + " "
PPP = PPP + Right(Space(15) + LTrim(Format(![catch], DECF1)), 15) + _
      " (" + Right(Space(5) + Format(PGTC, "##0.0"), 5) + "%) "
PPP = PPP + Right(Space(11) + LTrim(Format(![cpue], DECF2)), 11) + " "
PPP = PPP + Right(Space(15) + LTrim(Format(![Value], DECF1)), 15) + _
      " (" + Right(Space(5) + Format(PGTV, "##0.0"), 5) + "%) "
Print #1, Tab(5); PPP

PPP = Space(31) + Right(Space(15) + LTrim(Format(![eff], DECF1)), 15) + _
      Space(10)
PPP = PPP + Right(Space(11) + LTrim(Format(![kgfish], DECF2)), 11) + " "
PPP = PPP + Right(Space(15) + LTrim(Format(![price], DECF2)), 15) + " "

Print #1, Tab(5); PPP

Print #1, " "

'EXPORTING
'=========

ZC = ![catch]: ZE = ![eff]: ZU = ![cpue]: ZP = ![price]: ZV = ![Value]
ZF = (![FRNO])

If ZP = 0 Then
   ZP = "...": ZV = "..."
   End If

Write #3, RTrim(TSPN(J)), " ", ZE, ZC, ZU, ZP, ZV, " "

next_j:

Next J

Print #1, " "
Write #3, " ", " ", " ", " ", " ", " ", " ", " "

NEXT_I:

Next I

Close #1

End With

prm_record.Close
prm_database.Close

rtsDISP.FileName = APPROOT + "\ARTBAS\RESULTS\WORK.TXT"
rtsDISP.Visible = True
lblEXP.Visible = False

lstLIST.Visible = False
Frame1.Visible = False

Close #3

End Sub
Private Sub BYMN_BYSP_BYBG()

Dim ESTFNM

ESTFNM = APPROOT + "\ARTBAS\CURRENT_TABLES\ESTIM.TXT"

Open ESTFNM For Output As #3

Dim dbn, I, J, m, K, XKEY, LSTR

LSTR = 91

dbn = APPROOT + "\ARTBAS\RESULTS\WORK.MDB"

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(dbn)
Set prm_record = prm_database.OpenRecordset("ESTAB")

With prm_record

.Index = "primarykey"

Open APPROOT + "\ARTBAS\RESULTS\WORK.TXT" For Output As #1

Print #1, Tab(5); frmREP.Caption + " ( " + msgtab(37) + " )"

Print #1, " "

Print #1, Tab(5); msgtab(228) + " : " + RTrim(MINORN(Val(CURMNC)))
Print #1, Tab(5); msgtab(227) + " : " + optMN2.Caption

Print #1, Tab(5); String(LSTR, "=")

Print #1, Tab(5); msgtab(191)

Print #1, " "

'EXPORTING START
'===============

Write #3, " ", " ", " ", " ", " ", " ", " ", " "
Write #3, frmREP.Caption + " ( " + msgtab(37) + " )", " ", " ", " ", " ", " ", " ", " "
Write #3, " ", " ", " ", " ", " ", " ", " ", " "

Write #3, msgtab(191), " ", " ", " ", " ", " ", " ", " "
Write #3, " ", " ", " ", " ", " ", " ", " ", " "
Write #3, msgtab(192) + " " + UNW, " ", " ", " ", " ", " ", " ", " "
Write #3, msgtab(193) + " " + msgtab(198), " ", " ", " ", " ", " ", " ", " "
Write #3, msgtab(200) + " " + UNW + "/" + msgtab(198), " ", " ", " ", " ", " ", " ", " "
Write #3, msgtab(194) + " " + UNM + "/" + UNW, " ", " ", " ", " ", " ", " ", " "
Write #3, msgtab(195) + " " + UNM, " ", " ", " ", " ", " ", " ", " "
Write #3, msgtab(196) + " " + UNW + "/" + msgtab(197), " ", " ", " ", " ", " ", " ", " "
Write #3, "  ", " ", " ", " ", " ", " ", " ", " "

Write #3, "--- " + msgtab(228) + " : " + RTrim(MINORN(Val(CURMNC))) + " ---", _
" ", " ", " ", " ", " ", " ", " "
Write #3, "--- " + msgtab(227) + " : " + optMN2.Caption + " ---", _
" ", " ", " ", " ", " ", " ", " "
Write #3, " ", " ", " ", " ", " ", " ", " ", " "
Write #3, RTrim(msgtab(186)), RTrim(msgtab(33)); RTrim(msgtab(97)), _
             RTrim(msgtab(173)), RTrim(msgtab(199)), _
             RTrim(msgtab(94)), RTrim(msgtab(95)), " "
Write #3, " ", " ", " ", " ", " ", " ", " ", " "

'=========================

Print #1, Tab(5); msgtab(192) + " " + UNW
Print #1, Tab(5); msgtab(193) + " " + msgtab(198)
Print #1, Tab(5); msgtab(200) + " " + UNW + "/" + msgtab(198)
Print #1, Tab(5); msgtab(194) + " " + UNM + "/" + UNW
Print #1, Tab(5); msgtab(195) + " " + UNM
Print #1, Tab(5); msgtab(196) + " " + UNW + "/" + msgtab(197)

Print #1, " "
Write #3, "  ", " ", " ", " ", " ", " ", " ", " "

K = Val(CURMNC)

NTSP = NTSP + 1

ReDim Preserve TSPC(1 To NTSP), TSPN(1 To NTSP)

TSPC(NTSP) = "0000": TSPN(NTSP) = msgtab(201)

For I = 1 To NTSP

XKEY = "M" + CURMNC + "+B0000" + "+S" + TSPC(I)

.Seek "=", XKEY

If .NoMatch = True Then GoTo NEXT_I

Print #1, Tab(5); RTrim(MINORN(K)) + " : " + TSPN(I)

Print #1, Tab(5); String(LSTR, "=")

Dim PPP, DECF1, DECF2, GTC, GTE, GTV

DECF1 = "### ### ### ##0"
DECF2 = "### ##0.000"

GTC = ![catch]: GTE = ![eff]: GTV = ![Value]

PPP = Left(msgtab(165) + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(![eff], DECF1))

PPP = Left(msgtab(181) + " (" + UNW + ") " + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(![catch], DECF1))

PPP = Left(msgtab(173) + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(![cpue], DECF2))

PPP = Left(msgtab(184) + " (" + UNM + ") " + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(![Value], DECF1))

PPP = Left(msgtab(185) + " (" + UNM + "/" + UNW + ") " + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(![price], DECF2))

Print #1, " "

PPP = Left(msgtab(203) + Space(30), 30) + " "
PPP = PPP + Right(Space(15) + RTrim(msgtab(97)), 15) + Space(10)
PPP = PPP + Right(Space(11) + RTrim(msgtab(173)), 11) + " "
PPP = PPP + Right(Space(15) + RTrim(msgtab(95)), 15)

Print #1, Tab(5); PPP

PPP = Space(31)
PPP = PPP + Right(Space(15) + RTrim(msgtab(33)), 15) + Space(10)
PPP = PPP + Right(Space(11) + RTrim(msgtab(199)), 11) + " "
PPP = PPP + Right(Space(15) + RTrim(msgtab(94)), 15)

Print #1, Tab(5); PPP

Print #1, " "

'EXPORTING
'=========

ZC = ![catch]: ZE = ![eff]: ZU = ![cpue]: ZP = ![price]: ZV = ![Value]
ZF = Int(![FRNO]): ZW = ![kgfish]

If ZP = 0 Then ZP = "..."
If ZV = 0 Then ZV = "..."
If ZW = 0 Then ZW = "..."

Write #3, "--- " + RTrim(TSPN(I)) + " ---", ZE, ZC, ZU, ZW, ZP, ZV, " "
Write #3, " ", " ", " ", " ", " ", " ", " ", " "

'================

For J = 1 To NTBG

XKEY = "M" + CURMNC + "+B" + TBGC(J) + "+S" + TSPC(I)

.Seek "=", XKEY

If .NoMatch = True Then GoTo next_j
If ![catch] + ![eff] = 0 Then GoTo next_j:

Dim PGTC, PGTE, PGTV

PGTC = 0
If GTC <> 0 Then PGTC = 100 * ![catch] / GTC

PGTE = 0
If GTE <> 0 Then PGTE = 100 * ![eff] / GTE

PGTV = 0
If GTV <> 0 Then PGTV = 100 * ![Value] / GTV

PPP = TBGN(J) + " "
PPP = PPP + Right(Space(15) + LTrim(Format(![catch], DECF1)), 15) + _
      " (" + Right(Space(5) + Format(PGTC, "##0.0"), 5) + "%) "
PPP = PPP + Right(Space(11) + LTrim(Format(![cpue], DECF2)), 11) + " "
PPP = PPP + Right(Space(15) + LTrim(Format(![Value], DECF1)), 15) + _
      " (" + Right(Space(5) + Format(PGTV, "##0.0"), 5) + "%) "
Print #1, Tab(5); PPP

PPP = Space(31) + Right(Space(15) + LTrim(Format(![eff], DECF1)), 15) + _
      " (" + Right(Space(5) + Format(PGTE, "##0.0"), 5) + "%) "
PPP = PPP + Right(Space(11) + LTrim(Format(![kgfish], DECF2)), 11) + " "
PPP = PPP + Right(Space(15) + LTrim(Format(![price], DECF2)), 15) + " "

Print #1, Tab(5); PPP

Print #1, " "

'EXPORTING
'=========
ZC = ![catch]: ZE = ![eff]: ZU = ![cpue]: ZP = ![price]: ZV = ![Value]
ZF = Int(![FRNO]): ZW = ![kgfish]

If ZP = 0 Then ZP = "..."
If ZV = 0 Then ZV = "..."
If ZW = 0 Then ZW = "..."

Write #3, RTrim(TBGN(J)), ZE, ZC, ZU, ZW, ZP, ZV, " "

'======================

next_j:

Next J

Print #1, " "
Write #3, " ", " ", " ", " ", " ", " ", " ", " "

NEXT_I:

Next I

Close #1

End With

NTSP = NTSP - 1

prm_record.Close
prm_database.Close

rtsDISP.FileName = APPROOT + "\ARTBAS\RESULTS\WORK.TXT"
rtsDISP.Visible = True
lblEXP.Visible = False

lstLIST.Visible = False
Frame1.Visible = False

Close #3

End Sub
Private Sub BYMJ_BYSP_BYBG()

Dim ESTFNM

ESTFNM = APPROOT + "\ARTBAS\CURRENT_TABLES\ESTIM.TXT"

Open ESTFNM For Output As #3

Dim dbn, I, J, m, K, XKEY, LSTR

LSTR = 91

dbn = APPROOT + "\ARTBAS\RESULTS\WORK.MDB"

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(dbn)
Set prm_record = prm_database.OpenRecordset("ESTAB")

With prm_record

.Index = "primarykey"

Open APPROOT + "\ARTBAS\RESULTS\WORK.TXT" For Output As #1

Print #1, Tab(5); frmREP.Caption + " ( " + msgtab(37) + " )"

Print #1, " "

'EXPORTING START
'===============

Write #3, " ", " ", " ", " ", " ", " ", " ", " "
Write #3, frmREP.Caption + " ( " + msgtab(37) + " )", " ", " ", " ", " ", " ", " ", " "
Write #3, " ", " ", " ", " ", " ", " ", " ", " "

Write #3, "--- " + RTrim(MAJORN(Val(CURMJC))) + " ---", " ", " ", " ", " ", " ", " ", " "
Write #3, "--- " + msgtab(227) + " : " + optMJ2.Caption + " ---", " ", " ", " ", " ", " ", " ", " "

'=========================

Print #1, Tab(5); msgtab(229) + " : " + RTrim(MAJORN(Val(CURMJC)))
Print #1, Tab(5); msgtab(227) + " : " + optMJ2.Caption

Call NOT_INCLUDED

Write #3, " ", " ", " ", " ", " ", " ", " ", " "
Write #3, msgtab(191), " ", " ", " ", " ", " ", " ", " "
Write #3, " ", " ", " ", " ", " ", " ", " ", " "

Write #3, msgtab(192) + " " + UNW, " ", " ", " ", " ", " ", " ", " "
Write #3, msgtab(193) + " " + msgtab(198), " ", " ", " ", " ", " ", " ", " "
Write #3, msgtab(200) + " " + UNW + "/" + msgtab(198), " ", " ", " ", " ", " ", " ", " "
Write #3, msgtab(194) + " " + UNM + "/" + UNW, " ", " ", " ", " ", " ", " ", " "
Write #3, msgtab(195) + " " + UNM, " ", " ", " ", " ", " ", " ", " "
Write #3, msgtab(196) + " " + UNW + "/" + msgtab(197), " ", " ", " ", " ", " ", " ", " "
Write #3, " ", " ", " ", " ", " ", " ", " ", " "

Write #3, " ", RTrim(msgtab(33)), RTrim(msgtab(209)), RTrim(msgtab(173)), _
          RTrim(msgtab(94)), RTrim(msgtab(95)), " ", " "
Write #3, " ", " ", " ", " ", " ", " ", " ", " "

Print #1, Tab(5); String(LSTR, "=")

Print #1, Tab(5); msgtab(191)

Print #1, " "

Print #1, Tab(5); msgtab(192) + " " + UNW
Print #1, Tab(5); msgtab(193) + " " + msgtab(198)
Print #1, Tab(5); msgtab(200) + " " + UNW + "/" + msgtab(198)
Print #1, Tab(5); msgtab(194) + " " + UNM + "/" + UNW
Print #1, Tab(5); msgtab(195) + " " + UNM
Print #1, Tab(5); msgtab(196) + " " + UNW + "/" + msgtab(197)

Print #1, " "

K = Val(CURMJC)

NTSP = NTSP + 1

ReDim Preserve TSPC(1 To NTSP), TSPN(1 To NTSP)

TSPC(NTSP) = "0000": TSPN(NTSP) = msgtab(201)

For I = 1 To NTSP

XKEY = "M" + CURMJC + "+B0000" + "+S" + TSPC(I)

.Seek "=", XKEY

If .NoMatch = True Then GoTo NEXT_I

Print #1, Tab(5); RTrim(MAJORN(K)) + " : " + TSPN(I)

Print #1, Tab(5); String(LSTR, "=")

Dim PPP, DECF1, DECF2, GTC, GTE, GTV

DECF1 = "### ### ### ##0"
DECF2 = "### ##0.000"

GTC = ![catch]: GTE = ![eff]: GTV = ![Value]

PPP = Left(msgtab(165) + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(![eff], DECF1))

PPP = Left(msgtab(181) + " (" + UNW + ") " + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(![catch], DECF1))

PPP = Left(msgtab(173) + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(![cpue], DECF2))

PPP = Left(msgtab(184) + " (" + UNM + ") " + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(![Value], DECF1))

PPP = Left(msgtab(185) + " (" + UNM + "/" + UNW + ") " + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(![price], DECF2))

Print #1, " "

PPP = Left(msgtab(203) + Space(30), 30) + " "
PPP = PPP + Right(Space(15) + RTrim(msgtab(97)), 15) + Space(10)
PPP = PPP + Right(Space(11) + RTrim(msgtab(173)), 11) + " "
PPP = PPP + Right(Space(15) + RTrim(msgtab(95)), 15)

Print #1, Tab(5); PPP

PPP = Space(31)
PPP = PPP + Right(Space(15) + RTrim(msgtab(33)), 15) + Space(10)
PPP = PPP + Right(Space(11) + RTrim(msgtab(199)), 11) + " "
PPP = PPP + Right(Space(15) + RTrim(msgtab(94)), 15)

Print #1, Tab(5); PPP

Print #1, " "

'EXPORTING
'=========

ZC = ![catch]: ZE = ![eff]: ZU = ![cpue]: ZP = ![price]: ZV = ![Value]

If ZP = 0 Then
   ZP = "...": ZV = "..."
   End If

Write #3, "--- " + RTrim(TSPN(I)) + " ---", ZE, ZC, ZU, ZP, ZV, " ", " "
Write #3, " ", " ", " ", " ", " ", " ", " ", " "

'================

For J = 1 To NTBG

XKEY = "M" + CURMJC + "+B" + TBGC(J) + "+S" + TSPC(I)

.Seek "=", XKEY

If .NoMatch = True Then GoTo next_j
If ![catch] + ![eff] = 0 Then GoTo next_j:

Dim PGTC, PGTE, PGTV

PGTC = 0
If GTC <> 0 Then PGTC = 100 * ![catch] / GTC

PGTE = 0
If GTE <> 0 Then PGTE = 100 * ![eff] / GTE

PGTV = 0
If GTV <> 0 Then PGTV = 100 * ![Value] / GTV

PPP = TBGN(J) + " "
PPP = PPP + Right(Space(15) + LTrim(Format(![catch], DECF1)), 15) + _
      " (" + Right(Space(5) + Format(PGTC, "##0.0"), 5) + "%) "
PPP = PPP + Right(Space(11) + LTrim(Format(![cpue], DECF2)), 11) + " "
PPP = PPP + Right(Space(15) + LTrim(Format(![Value], DECF1)), 15) + _
      " (" + Right(Space(5) + Format(PGTV, "##0.0"), 5) + "%) "
Print #1, Tab(5); PPP

PPP = Space(31) + Right(Space(15) + LTrim(Format(![eff], DECF1)), 15) + _
      " (" + Right(Space(5) + Format(PGTE, "##0.0"), 5) + "%) "
PPP = PPP + Right(Space(11) + LTrim(Format(![kgfish], DECF2)), 11) + " "
PPP = PPP + Right(Space(15) + LTrim(Format(![price], DECF2)), 15) + " "

Print #1, Tab(5); PPP

Print #1, " "

'EXPORTING
'=========

ZC = ![catch]: ZE = ![eff]: ZU = ![cpue]: ZP = ![price]: ZV = ![Value]

If ZP = 0 Then
   ZP = "...": ZV = "..."
   End If

Write #3, RTrim(TBGN(J)), ZE, ZC, ZU, ZP, ZV, " ", " "

'======================

next_j:

Next J

Print #1, " "
Write #3, " ", " ", " ", " ", " ", " ", " ", " "

NEXT_I:

Next I

Close #1

End With

NTSP = NTSP - 1

prm_record.Close
prm_database.Close

rtsDISP.FileName = APPROOT + "\ARTBAS\RESULTS\WORK.TXT"
rtsDISP.Visible = True
lblEXP.Visible = False

lstLIST.Visible = False
Frame1.Visible = False

Close #3

End Sub
Private Sub BYGT_BYSP_BYBG()

Dim ESTFNM

ESTFNM = APPROOT + "\ARTBAS\CURRENT_TABLES\ESTIM.TXT"

Open ESTFNM For Output As #3

Dim dbn, I, J, m, K, XKEY, LSTR

LSTR = 91

dbn = APPROOT + "\ARTBAS\RESULTS\WORK.MDB"

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(dbn)
Set prm_record = prm_database.OpenRecordset("ESTAB")

With prm_record

.Index = "primarykey"

Open APPROOT + "\ARTBAS\RESULTS\WORK.TXT" For Output As #1

Print #1, Tab(5); frmREP.Caption + " ( " + msgtab(37) + " : " + msgtab(143) + " )"
Write #3, frmREP.Caption + " ( " + msgtab(37) + " : " + msgtab(143) + " )", _
" ", " ", " ", " ", " ", " ", " "

Print #1, " "
Write #3, " ", " ", " ", " ", " ", " ", " ", " "

Print #1, Tab(5); msgtab(227) + " : " + optGT2.Caption
Write #3, msgtab(227) + " : " + optGT2.Caption, " ", " ", " ", " ", " ", " ", " "

Write #3, " ", " ", " ", " ", " ", " ", " ", " "
Write #3, msgtab(191), " ", " ", " ", " ", " ", " ", " "
Write #3, "  ", " ", " ", " ", " ", " ", " ", " "
Write #3, msgtab(192) + " " + UNW, " ", " ", " ", " ", " ", " ", " "
Write #3, msgtab(193) + " " + msgtab(198), " ", " ", " ", " ", " ", " ", " "
Write #3, msgtab(200) + " " + UNW + "/" + msgtab(198), " ", " ", " ", " ", " ", " ", " "
Write #3, msgtab(194) + " " + UNM + "/" + UNW, " ", " ", " ", " ", " ", " ", " "
Write #3, msgtab(195) + " " + UNM, " ", " ", " ", " ", " ", " ", " "
Write #3, msgtab(196) + " " + UNW + "/" + msgtab(197), " ", " ", " ", " ", " ", " ", " "
Write #3, " ", " ", " ", " ", " ", " ", " ", " "

Call NOT_INCLUDED

Print #1, Tab(5); String(LSTR, "=")
Write #3, " ", " ", " ", " ", " ", " ", " ", " "

Print #1, Tab(5); msgtab(191)
Print #1, " "

'EXPORTING START
'===============

Write #3, " ", RTrim(msgtab(33)), RTrim(msgtab(209)), RTrim(msgtab(173)), _
          RTrim(msgtab(94)), RTrim(msgtab(95)), " ", " "
Write #3, " ", " ", " ", " ", " ", " ", " ", " "


'=========================

Print #1, Tab(5); msgtab(192) + " " + UNW
Print #1, Tab(5); msgtab(193) + " " + msgtab(198)
Print #1, Tab(5); msgtab(200) + " " + UNW + "/" + msgtab(198)
Print #1, Tab(5); msgtab(194) + " " + UNM + "/" + UNW
Print #1, Tab(5); msgtab(195) + " " + UNM
Print #1, Tab(5); msgtab(196) + " " + UNW + "/" + msgtab(197)

Print #1, " "

NTSP = NTSP + 1

ReDim Preserve TSPC(1 To NTSP), TSPN(1 To NTSP)

TSPC(NTSP) = "0000": TSPN(NTSP) = msgtab(201)

For I = 1 To NTSP

XKEY = "M0000" + "+B0000" + "+S" + TSPC(I)

.Seek "=", XKEY

If .NoMatch = True Then GoTo NEXT_I

Print #1, Tab(5); RTrim(msgtab(143)) + " : " + TSPN(I)

Print #1, Tab(5); String(LSTR, "=")

Dim PPP, DECF1, DECF2, GTC, GTE, GTV

DECF1 = "### ### ### ##0"
DECF2 = "### ##0.000"

GTC = ![catch]: GTE = ![eff]: GTV = ![Value]

PPP = Left(msgtab(165) + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(![eff], DECF1))

PPP = Left(msgtab(181) + " (" + UNW + ") " + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(![catch], DECF1))

PPP = Left(msgtab(173) + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(![cpue], DECF2))

PPP = Left(msgtab(184) + " (" + UNM + ") " + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(![Value], DECF1))

PPP = Left(msgtab(185) + " (" + UNM + "/" + UNW + ") " + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(![price], DECF2))

Print #1, " "

PPP = Left(msgtab(203) + Space(30), 30) + " "
PPP = PPP + Right(Space(15) + RTrim(msgtab(97)), 15) + Space(10)
PPP = PPP + Right(Space(11) + RTrim(msgtab(173)), 11) + " "
PPP = PPP + Right(Space(15) + RTrim(msgtab(95)), 15)

Print #1, Tab(5); PPP

PPP = Space(31)
PPP = PPP + Right(Space(15) + RTrim(msgtab(33)), 15) + Space(10)
PPP = PPP + Right(Space(11) + RTrim(msgtab(199)), 11) + " "
PPP = PPP + Right(Space(15) + RTrim(msgtab(94)), 15)

Print #1, Tab(5); PPP

Print #1, " "

'EXPORTING
'=========

ZC = ![catch]: ZE = ![eff]: ZU = ![cpue]: ZP = ![price]: ZV = ![Value]
ZF = (![FRNO])

If ZP = 0 Then
   ZP = "...": ZV = "..."
   End If

Write #3, "--- " + RTrim(TSPN(I)) + " ---", ZE, ZC, ZU, ZP, ZV, " ", " "
Write #3, " ", " ", " ", " ", " ", " ", " ", " "
'================

For J = 1 To NTBG

XKEY = "M0000" + "+B" + TBGC(J) + "+S" + TSPC(I)

.Seek "=", XKEY

If .NoMatch = True Then GoTo next_j
If ![catch] + ![eff] = 0 Then GoTo next_j:

Dim PGTC, PGTE, PGTV

PGTC = 0
If GTC <> 0 Then PGTC = 100 * ![catch] / GTC

PGTE = 0
If GTE <> 0 Then PGTE = 100 * ![eff] / GTE

PGTV = 0
If GTV <> 0 Then PGTV = 100 * ![Value] / GTV

PPP = TBGN(J) + " "
PPP = PPP + Right(Space(15) + LTrim(Format(![catch], DECF1)), 15) + _
      " (" + Right(Space(5) + Format(PGTC, "##0.0"), 5) + "%) "
PPP = PPP + Right(Space(11) + LTrim(Format(![cpue], DECF2)), 11) + " "
PPP = PPP + Right(Space(15) + LTrim(Format(![Value], DECF1)), 15) + _
      " (" + Right(Space(5) + Format(PGTV, "##0.0"), 5) + "%) "
Print #1, Tab(5); PPP

PPP = Space(31) + Right(Space(15) + LTrim(Format(![eff], DECF1)), 15) + _
      " (" + Right(Space(5) + Format(PGTE, "##0.0"), 5) + "%) "
PPP = PPP + Right(Space(11) + LTrim(Format(![kgfish], DECF2)), 11) + " "
PPP = PPP + Right(Space(15) + LTrim(Format(![price], DECF2)), 15) + " "

Print #1, Tab(5); PPP

Print #1, " "

'EXPORTING
'=========

ZC = ![catch]: ZE = ![eff]: ZU = ![cpue]: ZP = ![price]: ZV = ![Value]
ZF = (![FRNO])

If ZP = 0 Then
   ZP = "...": ZV = "..."
   End If

Write #3, RTrim(TBGN(J)), ZE, ZC, ZU, ZP, ZV, " ", " "

'======================

next_j:

Next J

Print #1, " "
Write #3, " ", " ", " ", " ", " ", " ", " ", " "

NEXT_I:

Next I

Close #1

End With

NTSP = NTSP - 1

prm_record.Close
prm_database.Close

rtsDISP.FileName = APPROOT + "\ARTBAS\RESULTS\WORK.TXT"
rtsDISP.Visible = True
lblEXP.Visible = False

lstLIST.Visible = False
Frame1.Visible = False

Close #3

End Sub
Private Sub BYMN_BYBG()

Dim ESTFNM

ESTFNM = APPROOT + "\ARTBAS\CURRENT_TABLES\ESTIM.TXT"

Open ESTFNM For Output As #3

Dim dbn, I, J, m, K, XKEY, LSTR

LSTR = 91

dbn = APPROOT + "\ARTBAS\RESULTS\WORK.MDB"

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(dbn)
Set prm_record = prm_database.OpenRecordset("ESTAB")

With prm_record

.Index = "primarykey"

Open APPROOT + "\ARTBAS\RESULTS\WORK.TXT" For Output As #1

Print #1, Tab(5); frmREP.Caption + " ( " + msgtab(37) + " )"

Print #1, " "

Print #1, Tab(5); msgtab(228) + " : " + RTrim(MINORN(Val(CURMNC))) + _
                  " ( " + optMN3.Caption + " )"

Print #1, Tab(5); String(LSTR, "=")

Print #1, Tab(5); msgtab(191)

Print #1, " "

'EXPORTING START
'===============

Write #3, " ", " ", " ", " ", " ", " ", " ", " "
Write #3, frmREP.Caption + " ( " + msgtab(37) + " )", " ", " ", " ", " ", " ", " ", " "
Write #3, " ", " ", " ", " ", " ", " ", " ", " "

Write #3, msgtab(191), " ", " ", " ", " ", " ", " ", " "
Write #3, "  ", " ", " ", " ", " ", " ", " ", " "
Write #3, msgtab(192) + " " + UNW, " ", " ", " ", " ", " ", " ", " "
Write #3, msgtab(193) + " " + msgtab(198), " ", " ", " ", " ", " ", " ", " "
Write #3, msgtab(200) + " " + UNW + "/" + msgtab(198), " ", " ", " ", " ", " ", " ", " "
Write #3, msgtab(194) + " " + UNM + "/" + UNW, " ", " ", " ", " ", " ", " ", " "
Write #3, msgtab(195) + " " + UNM, " ", " ", " ", " ", " ", " ", " "
Write #3, msgtab(196) + " " + UNW + "/" + msgtab(197), " ", " ", " ", " ", " ", " ", " "
Write #3, "  "" ", " ", " ", " ", " ", " ", " "

Write #3, "--- " + RTrim(MINORN(Val(CURMNC))) + _
                  " ( " + optMN3.Caption + " ) ---", " ", " ", " ", " ", " ", " ", " "
                  
Write #3, " ", " ", " ", " ", " ", " ", " ", " "
Write #3, " ", RTrim(msgtab(67)), RTrim(msgtab(33)), RTrim(msgtab(209)), RTrim(msgtab(173)), _
          RTrim(msgtab(94)), RTrim(msgtab(95)), " "
Write #3, " ", " ", " ", " ", " ", " ", " ", " "

'=========================

Print #1, Tab(5); msgtab(192) + " " + UNW
Print #1, Tab(5); msgtab(193) + " " + msgtab(198)
Print #1, Tab(5); msgtab(200) + " " + UNW + "/" + msgtab(198)
Print #1, Tab(5); msgtab(194) + " " + UNM + "/" + UNW
Print #1, Tab(5); msgtab(195) + " " + UNM
Print #1, Tab(5); msgtab(196) + " " + UNW + "/" + msgtab(197)

Print #1, " "

K = Val(CURMNC)

XKEY = "M" + CURMNC + "+B0000" + "+S0000"

.Seek "=", XKEY

If .NoMatch = True Then End

Print #1, Tab(5); RTrim(MINORN(K)) + " : " + msgtab(202)

Print #1, Tab(5); String(LSTR, "=")

Dim PPP, DECF1, DECF2, GTC, GTE, GTV

DECF1 = "### ### ### ##0"
DECF2 = "### ##0.000"

GTC = ![catch]: GTE = ![eff]: GTV = ![Value]

PPP = Left(msgtab(163) + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(![FRNO], DECF1))

PPP = Left(msgtab(165) + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(![eff], DECF1))

PPP = Left(msgtab(181) + " (" + UNW + ") " + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(![catch], DECF1))

PPP = Left(msgtab(173) + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(![cpue], DECF2))

PPP = Left(msgtab(184) + " (" + UNM + ") " + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(![Value], DECF1))

PPP = Left(msgtab(185) + " (" + UNM + "/" + UNW + ") " + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(![price], DECF2))

Print #1, " "

PPP = Left(msgtab(203) + Space(30), 30) + " "
PPP = PPP + Right(Space(15) + RTrim(msgtab(97)), 15) + Space(10)
PPP = PPP + Right(Space(11) + RTrim(msgtab(173)), 11) + " "
PPP = PPP + Right(Space(15) + RTrim(msgtab(95)), 15)

Print #1, Tab(5); PPP

PPP = Space(31)
PPP = PPP + Right(Space(15) + RTrim(msgtab(33)), 15) + Space(10)
PPP = PPP + Space(12)
PPP = PPP + Right(Space(15) + RTrim(msgtab(94)), 15)

Print #1, Tab(5); PPP

Print #1, " "

'EXPORTING
'=========

ZC = ![catch]: ZE = ![eff]: ZU = ![cpue]: ZP = ![price]: ZV = ![Value]
ZF = Int(![FRNO])

If ZP = 0 Then
   ZP = "...": ZV = "..."
   End If


Write #3, RTrim(msgtab(202)), ZF, ZE, ZC, ZU, ZP, ZV, " "
Write #3, " ", " ", " ", " ", " ", " ", " ", " "

'================

For J = 1 To NTBG

XKEY = "M" + CURMNC + "+B" + TBGC(J) + "+S0000"

.Seek "=", XKEY

If .NoMatch = True Then GoTo next_j

Dim PGTC, PGTE, PGTV

PGTC = 0
If GTC <> 0 Then PGTC = 100 * ![catch] / GTC

PGTE = 0
If GTE <> 0 Then PGTE = 100 * ![eff] / GTE

PGTV = 0

If GTV <> 0 Then PGTV = 100 * ![Value] / GTV

PPP = TBGN(J) + " "
PPP = PPP + Right(Space(15) + LTrim(Format(![catch], DECF1)), 15) + _
      " (" + Right(Space(5) + Format(PGTC, "##0.0"), 5) + "%) "
PPP = PPP + Right(Space(11) + LTrim(Format(![cpue], DECF2)), 11) + " "
PPP = PPP + Right(Space(15) + LTrim(Format(![Value], DECF1)), 15) + _
      " (" + Right(Space(5) + Format(PGTV, "##0.0"), 5) + "%) "
Print #1, Tab(5); PPP

PPP = Space(31) + Right(Space(15) + LTrim(Format(![eff], DECF1)), 15) + _
      " (" + Right(Space(5) + Format(PGTE, "##0.0"), 5) + "%) "
PPP = PPP + Space(12)
PPP = PPP + Right(Space(15) + LTrim(Format(![price], DECF2)), 15) + " "

Print #1, Tab(5); PPP

Print #1, " "

'EXPORTING
'=========
ZC = ![catch]: ZE = ![eff]: ZU = ![cpue]: ZP = ![price]: ZV = ![Value]
ZF = Int(![FRNO])

If ZP = 0 Then
   ZP = "...": ZV = "..."
   End If

Write #3, RTrim(TBGN(J)), ZF, ZE, ZC, ZU, ZP, ZV, " "

'================

next_j:

Next J

Print #1, " "
Write #3, " ", " ", " ", " ", " ", " ", " ", " "

Close #1

End With

prm_record.Close
prm_database.Close

rtsDISP.FileName = APPROOT + "\ARTBAS\RESULTS\WORK.TXT"
rtsDISP.Visible = True
lblEXP.Visible = False

lstLIST.Visible = False
Frame1.Visible = False

Close #3

End Sub
Private Sub BYMJ_BYBG()

Dim ESTFNM

ESTFNM = APPROOT + "\ARTBAS\CURRENT_TABLES\ESTIM.TXT"

Open ESTFNM For Output As #3

Dim dbn, I, J, m, K, XKEY, LSTR

LSTR = 91

dbn = APPROOT + "\ARTBAS\RESULTS\WORK.MDB"

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(dbn)
Set prm_record = prm_database.OpenRecordset("ESTAB")

With prm_record

.Index = "primarykey"

Open APPROOT + "\ARTBAS\RESULTS\WORK.TXT" For Output As #1

Print #1, Tab(5); frmREP.Caption + " ( " + msgtab(37) + " )"

Print #1, " "

Print #1, Tab(5); msgtab(229) + " : " + RTrim(MAJORN(Val(CURMJC))) + _
                  " ( " + optMJ3.Caption + " )"

Call NOT_INCLUDED

Print #1, Tab(5); String(LSTR, "=")

Print #1, Tab(5); msgtab(191)

Print #1, " "

'EXPORTING START
'===============

Write #3, " ", " ", " ", " ", " ", " ", " ", " "
Write #3, frmREP.Caption + " ( " + msgtab(37) + " )", " ", " ", " ", " ", " ", " ", " "
Write #3, " ", " ", " ", " ", " ", " ", " ", " "

Write #3, msgtab(191), " ", " ", " ", " ", " ", " ", " "
Write #3, "  ", " ", " ", " ", " ", " ", " ", " "
Write #3, msgtab(192) + " " + UNW, " ", " ", " ", " ", " ", " ", " "
Write #3, msgtab(193) + " " + msgtab(198), " ", " ", " ", " ", " ", " ", " "
Write #3, msgtab(200) + " " + UNW + "/" + msgtab(198), " ", " ", " ", " ", " ", " ", " "
Write #3, msgtab(194) + " " + UNM + "/" + UNW, " ", " ", " ", " ", " ", " ", " "
Write #3, msgtab(195) + " " + UNM, " ", " ", " ", " ", " ", " ", " "
Write #3, msgtab(196) + " " + UNW + "/" + msgtab(197), " ", " ", " ", " ", " ", " ", " "
Write #3, "  ", " ", " ", " ", " ", " ", " ", " "

Write #3, "--- " + RTrim(MAJORN(Val(CURMJC))) + _
                  " ( " + optMJ3.Caption + " ) ---", " ", " ", " ", " ", " ", " ", " "
Write #3, " ", " ", " ", " ", " ", " ", " ", " "
Write #3, " ", RTrim(msgtab(67)), RTrim(msgtab(33)), RTrim(msgtab(209)), RTrim(msgtab(173)), _
          RTrim(msgtab(94)), RTrim(msgtab(95)), " "
Write #3, " ", " ", " ", " ", " ", " ", " ", " "

'=========================

Print #1, Tab(5); msgtab(192) + " " + UNW
Print #1, Tab(5); msgtab(193) + " " + msgtab(198)
Print #1, Tab(5); msgtab(200) + " " + UNW + "/" + msgtab(198)
Print #1, Tab(5); msgtab(194) + " " + UNM + "/" + UNW
Print #1, Tab(5); msgtab(195) + " " + UNM
Print #1, Tab(5); msgtab(196) + " " + UNW + "/" + msgtab(197)

Print #1, " "

K = Val(CURMJC)

XKEY = "M" + CURMJC + "+B0000" + "+S0000"

.Seek "=", XKEY

If .NoMatch = True Then End

Print #1, Tab(5); RTrim(MAJORN(K)) + " : " + msgtab(202)

Print #1, Tab(5); String(LSTR, "=")

Dim PPP, DECF1, DECF2, GTC, GTE, GTV

DECF1 = "### ### ### ##0"
DECF2 = "### ##0.000"

GTC = ![catch]: GTE = ![eff]: GTV = ![Value]

PPP = Left(msgtab(163) + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(![FRNO], DECF1))

PPP = Left(msgtab(165) + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(![eff], DECF1))

PPP = Left(msgtab(181) + " (" + UNW + ") " + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(![catch], DECF1))

PPP = Left(msgtab(173) + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(![cpue], DECF2))

PPP = Left(msgtab(184) + " (" + UNM + ") " + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(![Value], DECF1))

PPP = Left(msgtab(185) + " (" + UNM + "/" + UNW + ") " + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(![price], DECF2))

Print #1, " "

PPP = Left(msgtab(203) + Space(30), 30) + " "
PPP = PPP + Right(Space(15) + RTrim(msgtab(97)), 15) + Space(10)
PPP = PPP + Right(Space(11) + RTrim(msgtab(173)), 11) + " "
PPP = PPP + Right(Space(15) + RTrim(msgtab(95)), 15)

Print #1, Tab(5); PPP

PPP = Space(31)
PPP = PPP + Right(Space(15) + RTrim(msgtab(33)), 15) + Space(10)
PPP = PPP + Space(12)
PPP = PPP + Right(Space(15) + RTrim(msgtab(94)), 15)

Print #1, Tab(5); PPP

Print #1, " "

'EXPORTING
'=========

ZC = ![catch]: ZE = ![eff]: ZU = ![cpue]: ZP = ![price]: ZV = ![Value]
ZF = Int(![FRNO])

If ZP = 0 Then
   ZP = "...": ZV = "..."
   End If

Write #3, RTrim(msgtab(202)), ZF, ZE, ZC, ZU, ZP, ZV, " "
Write #3, " ", " ", " ", " ", " ", " ", " ", " "

'================

For J = 1 To NTBG

XKEY = "M" + CURMJC + "+B" + TBGC(J) + "+S0000"

.Seek "=", XKEY

If .NoMatch = True Then GoTo next_j

Dim PGTC, PGTE, PGTV

PGTC = 0
If GTC <> 0 Then PGTC = 100 * ![catch] / GTC

PGTE = 0
If GTE <> 0 Then PGTE = 100 * ![eff] / GTE

PGTV = 0

If GTV <> 0 Then PGTV = 100 * ![Value] / GTV

PPP = TBGN(J) + " "
PPP = PPP + Right(Space(15) + LTrim(Format(![catch], DECF1)), 15) + _
      " (" + Right(Space(5) + Format(PGTC, "##0.0"), 5) + "%) "
PPP = PPP + Right(Space(11) + LTrim(Format(![cpue], DECF2)), 11) + " "
PPP = PPP + Right(Space(15) + LTrim(Format(![Value], DECF1)), 15) + _
      " (" + Right(Space(5) + Format(PGTV, "##0.0"), 5) + "%) "
Print #1, Tab(5); PPP

PPP = Space(31) + Right(Space(15) + LTrim(Format(![eff], DECF1)), 15) + _
      " (" + Right(Space(5) + Format(PGTE, "##0.0"), 5) + "%) "
PPP = PPP + Space(12)
PPP = PPP + Right(Space(15) + LTrim(Format(![price], DECF2)), 15) + " "

Print #1, Tab(5); PPP

Print #1, " "

'EXPORTING
'=========

ZC = ![catch]: ZE = ![eff]: ZU = ![cpue]: ZP = ![price]: ZV = ![Value]
ZF = Int(![FRNO])

If ZP = 0 Then
   ZP = "...": ZV = "..."
   End If

Write #3, RTrim(TBGN(J)), ZF, ZE, ZC, ZU, ZP, ZV, " "

'================

next_j:

Next J

Print #1, " "
Write #3, " ", " ", " ", " ", " ", " ", " ", " "

Close #1

End With

prm_record.Close
prm_database.Close

rtsDISP.FileName = APPROOT + "\ARTBAS\RESULTS\WORK.TXT"
rtsDISP.Visible = True
lblEXP.Visible = False

lstLIST.Visible = False
Frame1.Visible = False

Close #3

End Sub
Private Sub BYGT_BYBG()

Dim ESTFNM

ESTFNM = APPROOT + "\ARTBAS\CURRENT_TABLES\ESTIM.TXT"

Open ESTFNM For Output As #3

Dim dbn, I, J, m, K, XKEY, LSTR

LSTR = 91

dbn = APPROOT + "\ARTBAS\RESULTS\WORK.MDB"

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(dbn)
Set prm_record = prm_database.OpenRecordset("ESTAB")

With prm_record

.Index = "primarykey"

Open APPROOT + "\ARTBAS\RESULTS\WORK.TXT" For Output As #1

Print #1, Tab(5); frmREP.Caption + " " + "( " + msgtab(37) + " : " + msgtab(143) + " )"
Write #3, frmREP.Caption + " " + "( " + msgtab(37) + " : " + msgtab(143) + " )", _
" ", " ", " ", " ", " ", " ", " "

Write #3, " ", " ", " ", " ", " ", " ", " ", " "
Write #3, optGT3.Caption, " ", " ", " ", " ", " ", " ", " "
Write #3, " ", " ", " ", " ", " ", " ", " ", " "

Write #3, msgtab(191), " ", " ", " ", " ", " ", " ", " "
Write #3, "  ", " ", " ", " ", " ", " ", " ", " "
Write #3, msgtab(192) + " " + UNW, " ", " ", " ", " ", " ", " ", " "
Write #3, msgtab(193) + " " + msgtab(198), " ", " ", " ", " ", " ", " ", " "
Write #3, msgtab(200) + " " + UNW + "/" + msgtab(198), " ", " ", " ", " ", " ", " ", " "
Write #3, msgtab(194) + " " + UNM + "/" + UNW, " ", " ", " ", " ", " ", " ", " "
Write #3, msgtab(195) + " " + UNM, " ", " ", " ", " ", " ", " ", " "
Write #3, msgtab(196) + " " + UNW + "/" + msgtab(197), " ", " ", " ", " ", " ", " ", " "
Write #3, "  ", " ", " ", " ", " ", " ", " ", " "

Print #1, " "
Print #1, Tab(5); optGT3.Caption

Call NOT_INCLUDED

Print #1, Tab(5); String(LSTR, "=")
Write #3, " ", " ", " ", " ", " ", " ", " ", " "

Print #1, Tab(5); msgtab(191)

Print #1, " "

'EXPORTING START
'===============

Write #3, " ", RTrim(msgtab(67)), RTrim(msgtab(33)), RTrim(msgtab(209)), RTrim(msgtab(173)), _
          RTrim(msgtab(94)), RTrim(msgtab(95)), " "
Write #3, " ", " ", " ", " ", " ", " ", " ", " "

'=========================

Print #1, Tab(5); msgtab(192) + " " + UNW
Print #1, Tab(5); msgtab(193) + " " + msgtab(198)
Print #1, Tab(5); msgtab(200) + " " + UNW + "/" + msgtab(198)
Print #1, Tab(5); msgtab(194) + " " + UNM + "/" + UNW
Print #1, Tab(5); msgtab(195) + " " + UNM
Print #1, Tab(5); msgtab(196) + " " + UNW + "/" + msgtab(197)

Print #1, " "

XKEY = "M0000" + "+B0000" + "+S0000"

.Seek "=", XKEY

If .NoMatch = True Then End

Print #1, Tab(5); RTrim(msgtab(143)) + " : " + msgtab(202)

Print #1, Tab(5); String(LSTR, "=")

Dim PPP, DECF1, DECF2, GTC, GTE, GTV

DECF1 = "### ### ### ##0"
DECF2 = "### ##0.000"

GTC = ![catch]: GTE = ![eff]: GTV = ![Value]

PPP = Left(msgtab(163) + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(![FRNO], DECF1))

PPP = Left(msgtab(165) + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(![eff], DECF1))

PPP = Left(msgtab(181) + " (" + UNW + ") " + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(![catch], DECF1))

PPP = Left(msgtab(173) + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(![cpue], DECF2))

PPP = Left(msgtab(184) + " (" + UNM + ") " + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(![Value], DECF1))

PPP = Left(msgtab(185) + " (" + UNM + "/" + UNW + ") " + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(![price], DECF2))

Print #1, " "

PPP = Left(msgtab(203) + Space(30), 30) + " "
PPP = PPP + Right(Space(15) + RTrim(msgtab(97)), 15) + Space(10)
PPP = PPP + Right(Space(11) + RTrim(msgtab(173)), 11) + " "
PPP = PPP + Right(Space(15) + RTrim(msgtab(95)), 15)

Print #1, Tab(5); PPP

PPP = Space(31)
PPP = PPP + Right(Space(15) + RTrim(msgtab(33)), 15) + Space(10)
PPP = PPP + Space(12)
PPP = PPP + Right(Space(15) + RTrim(msgtab(94)), 15)

Print #1, Tab(5); PPP

Print #1, " "

'EXPORTING
'=========

ZC = ![catch]: ZE = ![eff]: ZU = ![cpue]: ZP = ![price]: ZV = ![Value]
ZF = Int(![FRNO])

If ZP = 0 Then
   ZP = "...": ZV = "..."
   End If

Write #3, RTrim(msgtab(202)), ZF, ZE, ZC, ZU, ZP, ZV, " "
Write #3, " ", " ", " ", " ", " ", " ", " ", " "

'================

For J = 1 To NTBG

XKEY = "M0000" + "+B" + TBGC(J) + "+S0000"

.Seek "=", XKEY

If .NoMatch = True Then GoTo next_j

Dim PGTC, PGTE, PGTV

PGTC = 0
If GTC <> 0 Then PGTC = 100 * ![catch] / GTC

PGTE = 0
If GTE <> 0 Then PGTE = 100 * ![eff] / GTE

PGTV = 0

If GTV <> 0 Then PGTV = 100 * ![Value] / GTV

PPP = TBGN(J) + " "
PPP = PPP + Right(Space(15) + LTrim(Format(![catch], DECF1)), 15) + _
      " (" + Right(Space(5) + Format(PGTC, "##0.0"), 5) + "%) "
PPP = PPP + Right(Space(11) + LTrim(Format(![cpue], DECF2)), 11) + " "
PPP = PPP + Right(Space(15) + LTrim(Format(![Value], DECF1)), 15) + _
      " (" + Right(Space(5) + Format(PGTV, "##0.0"), 5) + "%) "
Print #1, Tab(5); PPP

PPP = Space(31) + Right(Space(15) + LTrim(Format(![eff], DECF1)), 15) + _
      " (" + Right(Space(5) + Format(PGTE, "##0.0"), 5) + "%) "
PPP = PPP + Space(12)
PPP = PPP + Right(Space(15) + LTrim(Format(![price], DECF2)), 15) + " "

Print #1, Tab(5); PPP

Print #1, " "

'EXPORTING
'=========

ZC = ![catch]: ZE = ![eff]: ZU = ![cpue]: ZP = ![price]: ZV = ![Value]
ZF = Int(![FRNO])

If ZP = 0 Then
   ZP = "...": ZV = "..."
   End If

Write #3, RTrim(TBGN(J)), ZF, ZE, ZC, ZU, ZP, ZV, " "

'================

next_j:

Next J

Print #1, " "
Write #3, " ", " ", " ", " ", " ", " ", " ", " "
Close #1

End With

prm_record.Close
prm_database.Close

rtsDISP.FileName = APPROOT + "\ARTBAS\RESULTS\WORK.TXT"
rtsDISP.Visible = True
lblEXP.Visible = False

lstLIST.Visible = False
Frame1.Visible = False

Close #3

End Sub
Private Sub BYMN_BYSP()

Dim ESTFNM

ESTFNM = APPROOT + "\ARTBAS\CURRENT_TABLES\ESTIM.TXT"

Open ESTFNM For Output As #3

Dim dbn, I, J, m, K, XKEY, LSTR

LSTR = 91

dbn = APPROOT + "\ARTBAS\RESULTS\WORK.MDB"

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(dbn)
Set prm_record = prm_database.OpenRecordset("ESTAB")

With prm_record

.Index = "primarykey"

Open APPROOT + "\ARTBAS\RESULTS\WORK.TXT" For Output As #1

Print #1, Tab(5); frmREP.Caption + " ( " + msgtab(37) + " )"

Print #1, " "

Print #1, Tab(5); msgtab(228) + " : " + RTrim(MINORN(Val(CURMNC))) + _
                  " ( " + optMN4.Caption + " )"

Print #1, Tab(5); String(LSTR, "=")

Print #1, Tab(5); msgtab(191)

Print #1, " "

'EXPORTING START
'===============

Write #3, " ", " ", " ", " ", " ", " ", " ", " "
Write #3, frmREP.Caption + " ( " + msgtab(37) + " )", " ", " ", " ", " ", " ", " ", " "
Write #3, " ", " ", " ", " ", " ", " ", " ", " "

Write #3, msgtab(191), " ", " ", " ", " ", " ", " ", " "
Write #3, "  ", " ", " ", " ", " ", " ", " ", " "
Write #3, msgtab(192) + " " + UNW, " ", " ", " ", " ", " ", " ", " "
Write #3, msgtab(193) + " " + msgtab(198), " ", " ", " ", " ", " ", " ", " "
Write #3, msgtab(200) + " " + UNW + "/" + msgtab(198), " ", " ", " ", " ", " ", " ", " "
Write #3, msgtab(194) + " " + UNM + "/" + UNW, " ", " ", " ", " ", " ", " ", " "
Write #3, msgtab(195) + " " + UNM, " ", " ", " ", " ", " ", " ", " "
Write #3, msgtab(196) + " " + UNW + "/" + msgtab(197), " ", " ", " ", " ", " ", " ", " "
Write #3, "  ", " ", " ", " ", " ", " ", " ", " "

Write #3, "--- " + RTrim(MINORN(Val(CURMNC))) + _
                  " ( " + optMN4.Caption + " ) ---", " ", " ", " ", " ", " ", " ", " "
Write #3, " ", " ", " ", " ", " ", " ", " ", " "
Write #3, " ", RTrim(msgtab(33)), RTrim(msgtab(209)), RTrim(msgtab(173)), _
          RTrim(msgtab(94)), RTrim(msgtab(95)), " ", " "
Write #3, " ", " ", " ", " ", " ", " ", " ", " "

'=========================

Print #1, Tab(5); msgtab(192) + " " + UNW
Print #1, Tab(5); msgtab(193) + " " + msgtab(198)
Print #1, Tab(5); msgtab(200) + " " + UNW + "/" + msgtab(198)
Print #1, Tab(5); msgtab(194) + " " + UNM + "/" + UNW
Print #1, Tab(5); msgtab(195) + " " + UNM
Print #1, Tab(5); msgtab(196) + " " + UNW + "/" + msgtab(197)

Print #1, " "

K = Val(CURMNC)

XKEY = "M" + CURMNC + "+B0000" + "+S0000"

.Seek "=", XKEY

If .NoMatch = True Then End

Print #1, Tab(5); RTrim(MINORN(K)) + " : " + msgtab(201)

Print #1, Tab(5); String(LSTR, "=")

Dim PPP, DECF1, DECF2, GTC, GTE, GTV

DECF1 = "### ### ### ##0"
DECF2 = "### ##0.000"

GTC = ![catch]: GTE = ![eff]: GTV = ![Value]

PPP = Left(msgtab(165) + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(![eff], DECF1))

PPP = Left(msgtab(181) + " (" + UNW + ") " + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(![catch], DECF1))

PPP = Left(msgtab(173) + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(![cpue], DECF2))

PPP = Left(msgtab(184) + " (" + UNM + ") " + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(![Value], DECF1))

PPP = Left(msgtab(185) + " (" + UNM + "/" + UNW + ") " + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(![price], DECF2))

Print #1, " "

PPP = Left(msgtab(186) + Space(30), 30) + " "
PPP = PPP + Right(Space(15) + RTrim(msgtab(97)), 15) + Space(10)
PPP = PPP + Right(Space(15) + RTrim(msgtab(95)), 15)

Print #1, Tab(5); PPP

PPP = Space(31)
PPP = PPP + Space(25)
PPP = PPP + Right(Space(15) + RTrim(msgtab(94)), 15)

Print #1, Tab(5); PPP

Print #1, " "

'EXPORTING
'=========

ZE = ![eff]: ZC = ![catch]: ZU = ![cpue]: ZP = ![price]: ZV = ![Value]

If ZP = 0 Then
   ZP = "...": ZV = "..."
   End If

Write #3, RTrim(msgtab(201)), ZE, ZC, ZU, ZP, ZV, " ", " "
Write #3, " ", " ", " ", " ", " ", " ", " ", " "

'================

For J = 1 To NTSP

XKEY = "M" + CURMNC + "+B0000" + "+S" + TSPC(J)
.Seek "=", XKEY

If .NoMatch = True Then GoTo next_j
If ![catch] + ![eff] = 0 Then GoTo next_j

Dim PGTC, PGTE, PGTV

PGTC = 0
If GTC <> 0 Then PGTC = 100 * ![catch] / GTC

PGTE = 0
If GTE <> 0 Then PGTE = 100 * ![eff] / GTE

PGTV = 0

If GTV <> 0 Then PGTV = 100 * ![Value] / GTV

PPP = TSPN(J) + " "
PPP = PPP + Right(Space(15) + LTrim(Format(![catch], DECF1)), 15) + _
      " (" + Right(Space(5) + Format(PGTC, "##0.0"), 5) + "%) "
PPP = PPP + Right(Space(15) + LTrim(Format(![Value], DECF1)), 15) + _
      " (" + Right(Space(5) + Format(PGTV, "##0.0"), 5) + "%) "

Print #1, Tab(5); PPP

PPP = Space(31) + Space(25)
PPP = PPP + Right(Space(15) + LTrim(Format(![price], DECF2)), 15) + " "

Print #1, Tab(5); PPP

Print #1, " "

'EXPORTING
'=========

ZE = ![eff]: ZC = ![catch]: ZU = ![cpue]: ZP = ![price]: ZV = ![Value]

If ZP = 0 Then
   ZP = "...": ZV = "..."
   End If

Write #3, RTrim(TSPN(J)), ZE, ZC, ZU, ZP, ZV, " ", " "

'================

next_j:

Next J

Print #1, " "
Write #3, " ", " ", " ", " ", " ", " ", " ", " "

Close #1

End With

prm_record.Close
prm_database.Close

rtsDISP.FileName = APPROOT + "\ARTBAS\RESULTS\WORK.TXT"
rtsDISP.Visible = True
lblEXP.Visible = False

lstLIST.Visible = False
Frame1.Visible = False

Close #3

End Sub
Private Sub BYMJ_BYSP()

Dim ESTFNM

ESTFNM = APPROOT + "\ARTBAS\CURRENT_TABLES\ESTIM.TXT"

Open ESTFNM For Output As #3

Dim dbn, I, J, m, K, XKEY, LSTR

LSTR = 91

dbn = APPROOT + "\ARTBAS\RESULTS\WORK.MDB"

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(dbn)
Set prm_record = prm_database.OpenRecordset("ESTAB")

With prm_record

.Index = "primarykey"

Open APPROOT + "\ARTBAS\RESULTS\WORK.TXT" For Output As #1

Print #1, Tab(5); frmREP.Caption + " ( " + msgtab(37) + " )"

Print #1, " "

Print #1, Tab(5); msgtab(229) + " : " + RTrim(MAJORN(Val(CURMJC))) + _
                 " ( " + optMJ4.Caption + " )"

Call NOT_INCLUDED

Print #1, Tab(5); String(LSTR, "=")

Print #1, Tab(5); msgtab(191)

Print #1, " "

'EXPORTING START
'===============

Write #3, " ", " ", " ", " ", " ", " ", " ", " "
Write #3, frmREP.Caption + " ( " + msgtab(37) + " )", " ", " ", " ", " ", " ", " ", " "
Write #3, " ", " ", " ", " ", " ", " ", " ", " "

Write #3, msgtab(191), " ", " ", " ", " ", " ", " ", " "
Write #3, "  ", " ", " ", " ", " ", " ", " ", " "
Write #3, msgtab(192) + " " + UNW, " ", " ", " ", " ", " ", " ", " "
Write #3, msgtab(193) + " " + msgtab(198), " ", " ", " ", " ", " ", " ", " "
Write #3, msgtab(200) + " " + UNW + "/" + msgtab(198), " ", " ", " ", " ", " ", " ", " "
Write #3, msgtab(194) + " " + UNM + "/" + UNW, " ", " ", " ", " ", " ", " ", " "
Write #3, msgtab(195) + " " + UNM, " ", " ", " ", " ", " ", " ", " "
Write #3, msgtab(196) + " " + UNW + "/" + msgtab(197), " ", " ", " ", " ", " ", " ", " "
Write #3, "  ", " ", " ", " ", " ", " ", " ", " "

Write #3, "--- " + RTrim(MAJORN(Val(CURMJC))) + _
                  " ( " + optMJ4.Caption + " ) ---", " ", " ", " ", " ", " ", " ", " "
                  
Write #3, " ", " ", " ", " ", " ", " ", " ", " "
Write #3, " ", RTrim(msgtab(33)), RTrim(msgtab(209)), RTrim(msgtab(173)), _
          RTrim(msgtab(94)), RTrim(msgtab(95)), " ", " "
Write #3, " ", " ", " ", " ", " ", " ", " ", " "

'=========================

Print #1, Tab(5); msgtab(192) + " " + UNW
Print #1, Tab(5); msgtab(193) + " " + msgtab(198)
Print #1, Tab(5); msgtab(200) + " " + UNW + "/" + msgtab(198)
Print #1, Tab(5); msgtab(194) + " " + UNM + "/" + UNW
Print #1, Tab(5); msgtab(195) + " " + UNM
Print #1, Tab(5); msgtab(196) + " " + UNW + "/" + msgtab(197)

Print #1, " "

K = Val(CURMJC)

XKEY = "M" + CURMJC + "+B0000" + "+S0000"

.Seek "=", XKEY

If .NoMatch = True Then End

Print #1, Tab(5); RTrim(MAJORN(K)) + " : " + msgtab(201)

Print #1, Tab(5); String(LSTR, "=")

Dim PPP, DECF1, DECF2, GTC, GTE, GTV

DECF1 = "### ### ### ##0"
DECF2 = "### ##0.000"

GTC = ![catch]: GTE = ![eff]: GTV = ![Value]

PPP = Left(msgtab(165) + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(![eff], DECF1))

PPP = Left(msgtab(181) + " (" + UNW + ") " + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(![catch], DECF1))

PPP = Left(msgtab(173) + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(![cpue], DECF2))

PPP = Left(msgtab(184) + " (" + UNM + ") " + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(![Value], DECF1))

PPP = Left(msgtab(185) + " (" + UNM + "/" + UNW + ") " + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(![price], DECF2))

Print #1, " "

PPP = Left(msgtab(186) + Space(30), 30) + " "
PPP = PPP + Right(Space(15) + RTrim(msgtab(97)), 15) + Space(10)
PPP = PPP + Right(Space(15) + RTrim(msgtab(95)), 15)

Print #1, Tab(5); PPP

PPP = Space(31)
PPP = PPP + Space(25)
PPP = PPP + Right(Space(15) + RTrim(msgtab(94)), 15)

Print #1, Tab(5); PPP

Print #1, " "

'EXPORTING
'=========

ZC = ![catch]: ZE = ![eff]: ZU = ![cpue]: ZP = ![price]: ZV = ![Value]

If ZP = 0 Then
   ZP = "...": ZV = "..."
   End If

Write #3, RTrim(msgtab(201)), ZE, ZC, ZU, ZP, ZV, " ", " "
Write #3, " ", " ", " ", " ", " ", " ", " ", " "

'================

For J = 1 To NTSP

XKEY = "M" + CURMJC + "+B0000" + "+S" + TSPC(J)
.Seek "=", XKEY

If .NoMatch = True Then GoTo next_j
If ![catch] + ![eff] = 0 Then GoTo next_j

Dim PGTC, PGTE, PGTV

PGTC = 0
If GTC <> 0 Then PGTC = 100 * ![catch] / GTC

PGTE = 0
If GTE <> 0 Then PGTE = 100 * ![eff] / GTE

PGTV = 0

If GTV <> 0 Then PGTV = 100 * ![Value] / GTV

PPP = TSPN(J) + " "
PPP = PPP + Right(Space(15) + LTrim(Format(![catch], DECF1)), 15) + _
      " (" + Right(Space(5) + Format(PGTC, "##0.0"), 5) + "%) "
PPP = PPP + Right(Space(15) + LTrim(Format(![Value], DECF1)), 15) + _
      " (" + Right(Space(5) + Format(PGTV, "##0.0"), 5) + "%) "

Print #1, Tab(5); PPP

PPP = Space(31) + Space(25)
PPP = PPP + Right(Space(15) + LTrim(Format(![price], DECF2)), 15) + " "

Print #1, Tab(5); PPP

Print #1, " "

'EXPORTING
'=========

ZC = ![catch]: ZE = ![eff]: ZU = ![cpue]: ZP = ![price]: ZV = ![Value]

If ZP = 0 Then
   ZP = "...": ZV = "..."
   End If

Write #3, RTrim(TSPN(J)), ZE, ZC, ZU, ZP, ZV, " ", " "

'================

next_j:

Next J

Print #1, " "
Write #3, " ", " ", " ", " ", " ", " ", " ", " "

Close #1

End With

prm_record.Close
prm_database.Close

rtsDISP.FileName = APPROOT + "\ARTBAS\RESULTS\WORK.TXT"
rtsDISP.Visible = True
lblEXP.Visible = False

lstLIST.Visible = False
Frame1.Visible = False

Close #3

End Sub
Private Sub BYGT_BYSP()

Dim ESTFNM

ESTFNM = APPROOT + "\ARTBAS\CURRENT_TABLES\ESTIM.TXT"

Open ESTFNM For Output As #3

Dim dbn, I, J, m, K, XKEY, LSTR

LSTR = 91

dbn = APPROOT + "\ARTBAS\RESULTS\WORK.MDB"

Dim prm_database As Database, prm_record As Recordset

Set prm_database = OpenDatabase(dbn)
Set prm_record = prm_database.OpenRecordset("ESTAB")

With prm_record

.Index = "primarykey"

Open APPROOT + "\ARTBAS\RESULTS\WORK.TXT" For Output As #1

Print #1, Tab(5); frmREP.Caption + " " + "( " + msgtab(37) + " : " + msgtab(143) + " )"
Write #3, frmREP.Caption + " " + "( " + msgtab(37) + " : " + msgtab(143) + " )", _
" ", " ", " ", " ", " ", " ", " "

Print #1, " "

Write #3, " ", " ", " ", " ", " ", " ", " ", " "
Write #3, optGT4.Caption, " ", " ", " ", " ", " ", " ", " "
Write #3, " ", " ", " ", " ", " ", " ", " ", " "

Write #3, msgtab(191), " ", " ", " ", " ", " ", " ", " "
Write #3, "  ", " ", " ", " ", " ", " ", " ", " "
Write #3, msgtab(192) + " " + UNW, " ", " ", " ", " ", " ", " ", " "
Write #3, msgtab(193) + " " + msgtab(198), " ", " ", " ", " ", " ", " ", " "
Write #3, msgtab(200) + " " + UNW + "/" + msgtab(198), " ", " ", " ", " ", " ", " ", " "
Write #3, msgtab(194) + " " + UNM + "/" + UNW, " ", " ", " ", " ", " ", " ", " "
Write #3, msgtab(195) + " " + UNM, " ", " ", " ", " ", " ", " ", " "
Write #3, msgtab(196) + " " + UNW + "/" + msgtab(197), " ", " ", " ", " ", " ", " ", " "
Write #3, "  ", " ", " ", " ", " ", " ", " ", " "


Call NOT_INCLUDED

Print #1, Tab(5); String(LSTR, "=")
Write #3, " ", " ", " ", " ", " ", " ", " ", " "

Print #1, Tab(5); msgtab(191)

Print #1, " "

'EXPORTING START
'===============


Write #3, " ", RTrim(msgtab(33)), RTrim(msgtab(209)), RTrim(msgtab(173)), _
          RTrim(msgtab(94)), RTrim(msgtab(95)), " ", " "
Write #3, " ", " ", " ", " ", " ", " ", " ", " "

'=========================

Print #1, Tab(5); msgtab(192) + " " + UNW
Print #1, Tab(5); msgtab(193) + " " + msgtab(198)
Print #1, Tab(5); msgtab(200) + " " + UNW + "/" + msgtab(198)
Print #1, Tab(5); msgtab(194) + " " + UNM + "/" + UNW
Print #1, Tab(5); msgtab(195) + " " + UNM
Print #1, Tab(5); msgtab(196) + " " + UNW + "/" + msgtab(197)

Print #1, " "

XKEY = "M0000" + "+B0000" + "+S0000"

.Seek "=", XKEY

If .NoMatch = True Then End

Print #1, Tab(5); RTrim(msgtab(143)) + " : " + msgtab(201)

Print #1, Tab(5); String(LSTR, "=")

Dim PPP, DECF1, DECF2, GTC, GTE, GTV

DECF1 = "### ### ### ##0"
DECF2 = "### ##0.000"

GTC = ![catch]: GTE = ![eff]: GTV = ![Value]

PPP = Left(msgtab(165) + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(![eff], DECF1))

PPP = Left(msgtab(181) + " (" + UNW + ") " + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(![catch], DECF1))

PPP = Left(msgtab(173) + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(![cpue], DECF2))

PPP = Left(msgtab(184) + " (" + UNM + ") " + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(![Value], DECF1))

PPP = Left(msgtab(185) + " (" + UNM + "/" + UNW + ") " + String(40, "."), 40)
Print #1, Tab(5); PPP + " " + LTrim(Format(![price], DECF2))

Print #1, " "

PPP = Left(msgtab(186) + Space(30), 30) + " "
PPP = PPP + Right(Space(15) + RTrim(msgtab(97)), 15) + Space(10)
PPP = PPP + Right(Space(15) + RTrim(msgtab(95)), 15)

Print #1, Tab(5); PPP

PPP = Space(31)
PPP = PPP + Space(25)
PPP = PPP + Right(Space(15) + RTrim(msgtab(94)), 15)

Print #1, Tab(5); PPP

Print #1, " "

'EXPORTING
'=========

ZC = ![catch]: ZE = ![eff]
ZU = ![cpue]: ZP = ![price]: ZV = ![Value]

If ZP = 0 Then
   ZP = "...": ZV = "..."
   End If

Write #3, RTrim(msgtab(201)), ZE, ZC, ZU, ZP, ZV, " ", " "
Write #3, " ", " ", " ", " ", " ", " ", " ", " "

'================

For J = 1 To NTSP

XKEY = "M0000" + "+B0000" + "+S" + TSPC(J)
.Seek "=", XKEY

If .NoMatch = True Then GoTo next_j
If ![catch] + ![eff] = 0 Then GoTo next_j

Dim PGTC, PGTE, PGTV

PGTC = 0
If GTC <> 0 Then PGTC = 100 * ![catch] / GTC

PGTE = 0
If GTE <> 0 Then PGTE = 100 * ![eff] / GTE

PGTV = 0

If GTV <> 0 Then PGTV = 100 * ![Value] / GTV

PPP = TSPN(J) + " "
PPP = PPP + Right(Space(15) + LTrim(Format(![catch], DECF1)), 15) + _
      " (" + Right(Space(5) + Format(PGTC, "##0.0"), 5) + "%) "
PPP = PPP + Right(Space(15) + LTrim(Format(![Value], DECF1)), 15) + _
      " (" + Right(Space(5) + Format(PGTV, "##0.0"), 5) + "%) "

Print #1, Tab(5); PPP

PPP = Space(31) + Space(25)
PPP = PPP + Right(Space(15) + LTrim(Format(![price], DECF2)), 15) + " "

Print #1, Tab(5); PPP

Print #1, " "

'EXPORTING
'=========

ZC = ![catch]: ZE = ![eff]
ZU = ![cpue]: ZP = ![price]: ZV = ![Value]

If ZP = 0 Then
   ZP = "...": ZV = "..."
   End If

Write #3, RTrim(TSPN(J)), ZE, ZC, ZU, ZP, ZV, " ", " "

'================

next_j:

Next J

Print #1, " "
Write #3, " ", " ", " ", " ", " ", " ", " ", " "

Close #1

End With

prm_record.Close
prm_database.Close

rtsDISP.FileName = APPROOT + "\ARTBAS\RESULTS\WORK.TXT"
rtsDISP.Visible = True
lblEXP.Visible = False

lstLIST.Visible = False
Frame1.Visible = False

Close #3

End Sub
Private Sub NOT_INCLUDED()

If NOEXCL = 0 Then Exit Sub

Print #1, " "
Write #3, " ", " ", " ", " ", " ", " ", " ", " "

Print #1, Tab(5); msgtab(225)
Write #3, msgtab(225), " ", " ", " ", " ", " ", " ", " "

Print #1, " "
Write #3, " ", " ", " ", " ", " ", " ", " ", " "

Dim I

For I = 1 To NOEXCL

Print #1, Tab(5); EXN(I)
Write #3, EXN(I), " ", " ", " ", " ", " ", " ", " "

Next I

End Sub
Private Sub CHECK_LSAMPLES()

Dim fnm, XXX, I, J, V, W, K, yyy

NOMNBG = 0

fnm = APPROOT + "\ARTBAS\LANDINGS\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_LSAMPLES.TXT"

If Dir(fnm) = "" Then Exit Sub

Open fnm For Input As #1

Do Until EOF(1)

Line Input #1, XXX

XXX = Right(XXX, 16)

yyy = Left(XXX, 4) + Right(XXX, 4)

For K = 1 To NOMNBG

If TMNBG(K) = yyy Then GoTo DO_NOT_ADD

Next K

NOMNBG = NOMNBG + 1

ReDim Preserve TMNBG(1 To NOMNBG)

TMNBG(NOMNBG) = yyy

DO_NOT_ADD:

Loop

'SORT TABLE
'==========

If NOMNBG <= 1 Then GoTo NOT_SORT

For I = 1 To NOMNBG - 1
For J = I + 1 To NOMNBG

If TMNBG(I) <= TMNBG(J) Then GoTo next_j

XXX = TMNBG(I)
TMNBG(I) = TMNBG(J)
TMNBG(J) = XXX

next_j:

Next J
Next I

NOT_SORT:

Close #1

End Sub
Private Sub DISPLAY_MNBG()

Dim I, J, K

lstLIST.Visible = True
lstLIST.Clear

For K = 1 To NOMNBG

I = Val(Left(TMNBG(K), 4)): J = Val(Right(TMNBG(K), 4))

lstLIST.AddItem MINORN(I) + " " + BGN(J)

Next K

End Sub
Private Sub SELECT_MNBG()

lstLIST.MousePointer = 13

Dim I

I = lstLIST.ListIndex + 1

SELMBC = TMNBG(I)

Call CHECK_LSPECIES
Call DISPLAY_MBS

lstLIST.MousePointer = 1

DOCMNBG = "N": DOCSP = "Y"

End Sub
Private Sub CHECK_LSPECIES()

Dim fnm, XXX, I, J, V, W, K, yyy, ZZZ

NOMBS = 0

fnm = APPROOT + "\ARTBAS\LANDINGS\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_LSPECIES.TXT"

If Dir(fnm) = "" Then Exit Sub

Open fnm For Input As #1

Do Until EOF(1)

Line Input #1, ZZZ

XXX = Right(ZZZ, 16)

yyy = Left(XXX, 4) + Right(XXX, 4)

If yyy <> SELMBC Then GoTo DO_NOT_ADD

yyy = Mid(ZZZ, 10, 4)

For K = 1 To NOMBS

If TMBS(K) = yyy Then GoTo DO_NOT_ADD

Next K

NOMBS = NOMBS + 1

ReDim Preserve TMBS(1 To NOMBS)

TMBS(NOMBS) = yyy

DO_NOT_ADD:

Loop

Close #1

'SORT TABLE
'==========

If NOMBS <= 1 Then GoTo NOT_SORT

For I = 1 To NOMBS - 1
For J = I + 1 To NOMBS

XXX = SPEN(Val(TMBS(I))): yyy = SPEN(Val(TMBS(J)))

If XXX <= yyy Then GoTo next_j

XXX = TMBS(I)
TMBS(I) = TMBS(J)
TMBS(J) = XXX

next_j:

Next J
Next I

NOT_SORT:

End Sub
Private Sub DISPLAY_MBS()

Dim I, J, K

lstLIST.Visible = True
lstLIST.Clear

lstLIST.AddItem "[" + RTrim(msgtab(201)) + "]"

For K = 1 To NOMBS

I = Val(TMBS(K))

lstLIST.AddItem SPEN(I)

Next K

End Sub
Private Sub SELECT_MBS()

lstLIST.Visible = False

Dim I

I = lstLIST.ListIndex + 1

If I = 1 Then
   Call DOC_TOTALS
   lstLIST.MousePointer = 1
   DOCMNBG = "N"
   DOCSP = "N"
   Exit Sub
   End If

If I > 1 Then
   SELSPC = TMBS(I - 1)
   Call DOC_SPECIES
   lstLIST.MousePointer = 1
   DOCMNBG = "N"
   DOCSP = "N"
   Exit Sub
   End If

End Sub
Private Sub DOC_TOTALS()

lstLIST.Visible = False

GRAND_TOTAL_SP = 0

DOCS_FLAG = "TOTALS"

On Error GoTo EXIT_SUB_FINAL2

Call FRAME_MINOR_TOTALS2

EXPFNM = APPROOT + "\ARTBAS\CURRENT_TABLES\RAWDATA.TXT"

Open EXPFNM For Output As #3

Dim I, J, XXX, yyy, ZZZ, fnm

Open APPROOT + "\ARTBAS\RESULTS\WORK.TXT" For Output As #2

fnm = APPROOT + "\ARTBAS\LANDINGS\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_LSAMPLES.TXT"

If Dir(fnm) = "" Then GoTo EFFORT_ROUTINE

I = Val(Left(SELMBC, 4))

SELMBN = RTrim(MINORN(I))

I = Val(Mid(SELMBC, 5, 4))

SELMBN = SELMBN + " " + RTrim(BGN(I))

Print #2, Tab(5); frmREP.Caption
Print #2, " "
Print #2, Tab(5); msgtab(247)
Print #2, ""
Print #2, Tab(5); SELMBN + ":" + msgtab(171)
Print #2, " "

Write #3, frmREP.Caption; " "; " "; " "; " "
Write #3, " "; " "; " "; " "; " "
Write #3, msgtab(247); " "; " "; " "; " "
Write #3, " "; " "; " "; " "; " "
Write #3, SELMBN + ":" + msgtab(171); " "; " "; " "; " "
Write #3, " "; " "; " "; " "; " "

ZZZ = ""
ZZZ = ZZZ + Right(Space(10) + RTrim(msgtab(88)), 10) + " "
ZZZ = ZZZ + Right(Space(10) + RTrim(msgtab(89)), 10) + " "
ZZZ = ZZZ + Right(Space(10) + RTrim(msgtab(90)), 10) + " "
ZZZ = ZZZ + Right(Space(10) + RTrim(msgtab(91)), 10) + " "

Print #2, Tab(5); msgtab(248); Tab(54); ZZZ
Print #2, Tab(5); String(92, "=")

Write #3, msgtab(248); msgtab(88); msgtab(89); msgtab(90); msgtab(91)
Write #3, " "; " "; " "; " "; " "

Open fnm For Input As #1

Dim SE, SC

SE = 0: SC = 0

Do Until EOF(1)

Line Input #1, XXX

ZZZ = Right(XXX, 16)

J = Val(Mid(ZZZ, 7, 4))

yyy = Left(ZZZ, 4) + Right(ZZZ, 4)

If yyy <> SELMBC Then GoTo CONT_READ

ZZZ = Space(4)

ZZZ = ZZZ + Left(XXX, 6) + " "
ZZZ = ZZZ + Mid(XXX, 8, 2) + "/" + Format(current_month, "00") + "/" + Format(current_year, "0000") + " "
ZZZ = ZZZ + SITEN(J) + " "

Dim V1, V2, V3, V4, ZZZ_ZZZ

ZZZ_ZZZ = ZZZ

V1 = CDbl(Mid(XXX, 11, 8))
V2 = CDbl(Mid(XXX, 20, 8))
V3 = CDbl(Mid(XXX, 29, 19))
V4 = CDbl(Mid(XXX, 49, 19))

SE = SE + V1 * V2: SC = SC + V4

ZZZ = ZZZ + Right(Space(10) + (Format(CDbl(Mid(XXX, 11, 8)), "#####0.0##")), 10) + " "
ZZZ = ZZZ + Right(Space(10) + (Format(CDbl(Mid(XXX, 20, 8)), "#####0.0##")), 10) + " "
ZZZ = ZZZ + Right(Space(10) + (Format(CDbl(Mid(XXX, 29, 19)), "#####0.0##")), 10) + " "
ZZZ = ZZZ + Right(Space(10) + (Format(CDbl(Mid(XXX, 49, 19)), "#####0.0##")), 10) + " "

Print #2, ZZZ

ZZZ = Left(XXX, 6) + " "
ZZZ = ZZZ + Mid(XXX, 8, 2) + "/" + Format(current_month, "00") + "/" + Format(current_year, "0000") + " "
ZZZ = ZZZ + SITEN(J) + " "

Write #3, ZZZ; V1, V2, V3, V4

CONT_READ:

Loop

SAMPLE_TOT = SC

Print #2, " "
Write #3, " "; " "; " "; " "; " "

Print #2, Tab(5); msgtab(300); Tab(75); Format(SAMPLE_TOT, "#####0.0##")
Print #2, " "

Write #3, msgtab(300); SAMPLE_TOT; " "; " "; " "
Write #3, " "; " "; " "; " "; " "

If SE <> 0 Then

   FINAL_CPUE = SC / SE
   Print #2, Tab(5); msgtab(296); Tab(75); Format(SC / SE, "#####0.0##") + "  "
   Write #3, msgtab(296); SC / SE; " "; " "; " "
   End If

Print #2, " "
Write #3, " "; " "; " "; " "; " "

Print #2, Tab(5); String(92, "=")

Close #1

EFFORT_ROUTINE:

'--------- CHECK EFFORT FILE ---------------------

fnm = APPROOT + "\ARTBAS\EFFORT\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_ESAMPLES.TXT"

If Dir(fnm) = "" Then GoTo ACTIVE_ROUTINE

'=====================  EFFORT ROUTINE ==============

I = Val(Left(SELMBC, 4))

SELMBN = RTrim(MINORN(I))

I = Val(Mid(SELMBC, 5, 4))

SELMBN = SELMBN + " " + RTrim(BGN(I))

Print #2, Tab(5); SELMBN + ":" + msgtab(172)
Print #2, " "

Write #3, SELMBN + ":" + msgtab(172); " "; " "; " "; " "
Write #3, " "; " "; " "; " "; " "

ZZZ = ""

Dim WWW

ZZZ = ZZZ + Right(Space(10) + WWW, 10) + " "
ZZZ = ZZZ + Right(Space(10) + WWW, 10) + " "
ZZZ = ZZZ + Right(Space(10) + WWW, 10) + " "
ZZZ = ZZZ + Right(Space(10) + WWW, 10) + " "

Print #2, Tab(5); msgtab(43); Tab(57); msgtab(98); Tab(68); msgtab(85); _
          Tab(78); msgtab(154); Tab(90); msgtab(86)
          
Print #2, Tab(5); String(92, "=")

Write #3, " "; " "; " "; " "; " "
Write #3, msgtab(43); msgtab(98); msgtab(85); msgtab(154); msgtab(86)
Write #3, " "; " "; " "; " "; " "

Open fnm For Input As #1

Dim SA, ss, SF

SA = 0: ss = 0

Do Until EOF(1)

Line Input #1, XXX

yyy = Mid(XXX, 17, 4) + Mid(XXX, 8, 4)

If yyy <> SELMBC Then GoTo CONT_READ2

J = Val(Mid(XXX, 2, 4))

V1 = CDbl(Mid(XXX, 14, 2))
V2 = CDbl(Mid(XXX, 33, 10))
V3 = CDbl(Mid(XXX, 22, 10))
V4 = CDbl(Mid(XXX, 44, 10))

If V2 = 0 Then V2 = V4

ss = ss + V2: SA = SA + V3

If V3 = 0 Then ss = ss + V4
         
ZZZ = Space(4) + Left(RTrim(SITEN(J)) + String(48, "."), 45)
        
Print #2, Tab(5); SITEN(J); Tab(58); V1; Tab(69); V2; Tab(79); V3; Tab(91); V4

Write #3, SITEN(J); V1; V2; V3; V4

CONT_READ2:

Loop

Dim BAC_CAB

BAC_CAB = Left(msgtab(155), 3)
BAC_CAB = msgtab(297)

Dim CAB

Print #2, " "
Write #3, " "; " "; " "; " "; " "

If ss <> 0 Then

   FINAL_BAC = 100 * SA / ss
   Print #2, Tab(5); BAC_CAB; Tab(75); Format(100 * SA / ss, "###0.0##")
   Write #3, BAC_CAB; 100 * SA / ss; " "; " "; " "
   End If

EXIT_SUB:

Close #1

ACTIVE_ROUTINE:

'------------------ ACTIVE DAYS ----------------------
fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_ACTIVE.TXT"
      
If Dir(fnm) = "" Then GoTo FRAME_ROUTINE

'================================================================
Open fnm For Input As #1

Do Until EOF(1)

Line Input #1, XXX

Dim ACTMC, ACTBGC, ACTDAYS

XXX = XXX + Space(10)

ACTMC = Mid(XXX, 2, 4): ACTBGC = Mid(XXX, 8, 4): ACTDAYS = Mid(XXX, 26, 10)

yyy = ACTMC + ACTBGC

If yyy <> SELMBC Then GoTo CONT_READ3

' --------- found it ---------

FINAL_ACT = CDbl(ACTDAYS)

Print #2, ""
Print #2, Tab(5); msgtab(298); Tab(75); CDbl(ACTDAYS)

Write #3, " "; " "; " "; " "; " "
Write #3, msgtab(298); CDbl(ACTDAYS); " "; " "; " "

Close #1

GoTo FRAME_ROUTINE

CONT_READ3:

Loop

Close #1

GoTo FRAME_ROUTINE

FRAME_ROUTINE:
      
'------------------------------------------------------
FINAL_FRAME = 0

I = Val(Left(SELMBC, 4)): J = Val(Right(SELMBC, 4))

FINAL_FRAME = DOCTMNBG(I, J)

Print #2, ""
Print #2, Tab(5); msgtab(299); Tab(75); FINAL_FRAME
Print #2, " "

Write #3, " "; " "; " "; " "; " "
Write #3, msgtab(299); FINAL_FRAME; " "; " "; " "
Write #3, " "; " "; " "; " "; " "

TEST_RESULTS:

Dim EST_EFF, EST_CATCH

FINAL_EFFORT = 0: FINAL_CATCH = 0

If FINAL_CPUE * FINAL_BAC * FINAL_ACT * FINAL_FRAME = 0 Then GoTo EXIT_SUB_FINAL

EST_EFF = FINAL_BAC * FINAL_ACT * FINAL_FRAME / 100

Print #2, Tab(5); String(92, "=")
Print #2, Tab(5); msgtab(301); _
          Tab(75); Format(EST_EFF, "#######0")

Write #3, msgtab(301); EST_EFF; " "; " "; " "
          
Print #2, " "
Print #2, Tab(5); msgtab(302); Tab(75); Format(EST_EFF * FINAL_CPUE, "###########0")
Print #2, Tab(5); String(92, "=")

Write #3, " "; " "; " "; " "; " "
Write #3, msgtab(302); EST_EFF * FINAL_CPUE; " "; " "; " "

Write #3, " "; " "; " "; " "; " "
Write #3, " "; " "; " "; " "; " "

FINAL_EFFORT = EST_EFF: FINAL_CATCH = EST_EFF * FINAL_CPUE

'------------ LARGE LOOP FOR SPECIES ----------------

If NOMBS = 0 Then GoTo EXIT_SUB_FINAL

Print #2, Tab(5); msgtab(247) + " (" + msgtab(47) + ")"
Print #2, ""

Write #3, msgtab(247) + " (" + msgtab(47) + ")"; " "; " "; " "; " "
Write #3, " "; " "; " "; " "; " "

For KSP = 1 To NOMBS

SELSPC = LTrim(RTrim(TMBS(KSP)))

Call DOC_SPECIES

Next KSP

EXIT_SUB_FINAL:

Close #2
Close #3

rtsDISP.FileName = APPROOT + "\ARTBAS\RESULTS\WORK.TXT"
rtsDISP.Visible = True
cmdPRINT.Visible = True

cmdEXCEL_RAW.Visible = True

EXIT_SUB_FINAL2:

End Sub
Private Sub DOC_SPECIES()

Dim PQ, WQ

DOCS_FLAG = "SPECIES"

EXPFNM = APPROOT + "\ARTBAS\CURRENT_TABLES\RAWDATA.TXT"
'Open EXPFNM For Output As #3

Dim I, J, XXX, yyy, ZZZ, fnm, T_Q, T_F, T_V, AV_P, T_W, AV_W
Dim DOCDAT()

ReDim DOCDAT(1 To 99000)

fnm = APPROOT + "\ARTBAS\LANDINGS\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_LSAMPLES.TXT"

If Dir(fnm) = "" Then Exit Sub

Open fnm For Input As #1

Do Until EOF(1)

Line Input #1, XXX

J = Val(Left(XXX, 6))

DOCDAT(J) = Mid(XXX, 8, 2) + "/" + Format(current_month, "00") + "/" + Format(current_year, "0000")

Loop

Close #1

'Open APPROOT + "\ARTBAS\RESULTS\WORK.TXT" For Output As #2

I = Val(Left(SELMBC, 4))

SELMBN = RTrim(MINORN(I))

I = Val(Mid(SELMBC, 5, 4))

SELMBN = SELMBN + " " + RTrim(BGN(I))

I = Val(SELSPC)

Print #2, Tab(5); "--- " + RTrim(SPEN(I)) + " ---"
Print #2, " "

Write #3, "--- " + RTrim(SPEN(I)) + " ---"; " "; " "; " "; " "
Write #3, " "; " "; " "; " "; " "

ZZZ = ""
ZZZ = ZZZ + Right(Space(10) + RTrim(msgtab(97)), 10) + " "
ZZZ = ZZZ + Right(Space(10) + RTrim(msgtab(93)), 10) + " "
ZZZ = ZZZ + Right(Space(10) + RTrim(msgtab(94)), 10) + " "
ZZZ = ZZZ + Right(Space(10) + RTrim(msgtab(95)), 10) + " "

Print #2, Tab(5); msgtab(248); Tab(54); ZZZ
Print #2, Tab(5); String(92, "=")

Write #3, msgtab(248); msgtab(97); msgtab(93); msgtab(94); msgtab(95)
Write #3, " "; " "; " "; " "; " "

fnm = APPROOT + "\ARTBAS\LANDINGS\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_LSPECIES.TXT"

If Dir(fnm) = "" Then Exit Sub

Open fnm For Input As #1

Dim T_PQ, T_WQ

T_Q = 0: T_F = 0: T_V = 0: AV_P = 0: AV_W = 0: T_PQ = 0: T_WQ = 0

Do Until EOF(1)

Line Input #1, XXX

ZZZ = Right(XXX, 16)

J = Val(Mid(ZZZ, 7, 4))

yyy = Left(ZZZ, 4) + Right(ZZZ, 4)

If yyy <> SELMBC Then GoTo CONT_READ

If Mid(XXX, 10, 4) <> SELSPC Then GoTo CONT_READ

ZZZ = Space(4)

ZZZ = ZZZ + Mid(XXX, 2, 6) + " "

I = Val(Mid(XXX, 2, 6))

ZZZ = ZZZ + DOCDAT(I) + " "
ZZZ = ZZZ + SITEN(J) + " "

ZZZ = ZZZ + Right(Space(10) + Format(CDbl(Mid(XXX, 15, 19)), "#####0.0##"), 10) + " "
ZZZ = ZZZ + Right(Space(10) + Format(CDbl(Mid(XXX, 35, 19)), "#########0"), 10) + " "
ZZZ = ZZZ + Right(Space(10) + Format(CDbl(Mid(XXX, 55, 19)), "#####0.0##"), 10) + " "
ZZZ = ZZZ + Right(Space(10) + Format(CDbl(Mid(XXX, 75, 19)), "#####0.0##"), 10) + " "

Dim PPP, QQQ, WWW

WWW = CDbl(Mid(XXX, 35, 19))
PPP = CDbl(Mid(XXX, 75, 19))
QQQ = CDbl(Mid(XXX, 15, 19))

T_Q = T_Q + CDbl(Mid(XXX, 15, 19))
T_F = T_F + CDbl(Mid(XXX, 35, 19))
T_V = T_V + CDbl(Mid(XXX, 75, 19))

If QQQ <> 0 And PPP <> 0 Then T_PQ = T_PQ + CDbl(Mid(XXX, 15, 19))
If QQQ <> 0 And WWW <> 0 Then T_WQ = T_WQ + CDbl(Mid(XXX, 15, 19))

Print #2, ZZZ

Write #3, Mid(XXX, 2, 6) + " " + DOCDAT(I) + " " + SITEN(J); _
          CDbl(Mid(XXX, 15, 19)); CDbl(Mid(XXX, 35, 19)); CDbl(Mid(XXX, 55, 19)); _
          CDbl(Mid(XXX, 75, 19))

I = Val(Mid(XXX, 2, 6))

ZZZ = ZZZ + DOCDAT(I) + " "
ZZZ = ZZZ + SITEN(J) + " "

CONT_READ:

Loop

Close #1

If T_PQ <> 0 And T_V <> 0 Then AV_P = T_V / T_PQ
If T_F <> 0 And T_WQ <> 0 Then AV_W = T_WQ / T_F

ZZZ = Space(4)
ZZZ = ZZZ + Left(msgtab(314) + Space(60), 49)
ZZZ = ZZZ + Right(Space(10) + Format(T_Q, "#####0.0##"), 10) + " "
ZZZ = ZZZ + Right(Space(10) + Format(T_F, "#########0"), 10) + " "
ZZZ = ZZZ + Right(Space(10) + Format(AV_P, "#####0.0##"), 10) + " "
ZZZ = ZZZ + Right(Space(10) + Format(T_V, "#####0.0##"), 10) + " "

Dim PROP

PROP = 0

If SAMPLE_TOT <> 0 Then PROP = T_Q / SAMPLE_TOT

'--------- Print reminder of totals ------------------
Print #2, " "
Print #2, Tab(5); msgtab(315); Tab(75); Format(SAMPLE_TOT, "#####0.0##")

Write #3, " "; " "; " "; " "; " "
Write #3, msgtab(315); SAMPLE_TOT; " "; " "; " "

Print #2, Tab(5); msgtab(316); Tab(75); Format(FINAL_EFFORT, "#######0")

Write #3, msgtab(316); FINAL_EFFORT; " "; " "; " "
          
Print #2, Tab(5); msgtab(317); Tab(75); Format(FINAL_EFFORT * FINAL_CPUE, "###########0")
Print #2, " "

Write #3, msgtab(317); FINAL_EFFORT * FINAL_CPUE; " "; " "; " "
Write #3, " "; " "; " "; " "; " "
'-----------------------------------------------------------

If SAMPLE_TOT <> 0 Then PROP = T_Q / SAMPLE_TOT

Print #2, Tab(5); msgtab(303); Tab(75); LTrim(RTrim(Format(T_Q, "########0.0")))
Write #3, msgtab(303); T_Q; " "; " "; " "

Print #2, Tab(5); msgtab(304); Tab(75); LTrim(RTrim(Format(PROP, "#####0.0##")))
Write #3, msgtab(304); PROP; " "; " "; " "

Print #2, Tab(5); msgtab(305); Tab(75); LTrim(RTrim(Format(PROP * FINAL_CATCH, "#########0")))

Write #3, msgtab(305); PROP * FINAL_CATCH; " "; " "; " "

If FINAL_EFFORT <> 0 <> 0 Then
   Print #2, Tab(5); msgtab(306); Tab(75); LTrim(RTrim(Format(PROP * FINAL_CATCH / FINAL_EFFORT, "#####0.0##")))
   Print #2, " "

   Write #3, msgtab(306); PROP * FINAL_CATCH / FINAL_EFFORT; " "; " "; " "
   Write #3, " "; " "; " "; " "; " "
End If

'------------------------------
Print #2, Tab(5); msgtab(307); Tab(75); LTrim(RTrim(Format(T_V, "########0.0")))
Write #3, msgtab(307); T_V; " "; " "; " "

Print #2, Tab(5); msgtab(308); Tab(75); LTrim(RTrim(Format(AV_P, "########0.0##")))
Write #3, msgtab(308); AV_P; " "; " "; " "

Print #2, Tab(5); msgtab(309); Tab(75); LTrim(RTrim(Format(PROP * FINAL_CATCH * AV_P, "##########0")))
Write #3, msgtab(309); PROP * FINAL_CATCH * AV_P; " "; " "; " "
'--------------------------------------------
Print #2, " "
Write #3, " "; " "; " "; " "; " "

Print #2, Tab(5); msgtab(310); Tab(75); LTrim(RTrim(Format(T_F, "##########0")))
Write #3, msgtab(310); T_F; " "; " "; " "

Print #2, Tab(5); msgtab(311); Tab(75); LTrim(RTrim(Format(AV_W, "########0.0##")))
Write #3, msgtab(311); 1000 * AV_W / 1000; " "; " "; " "

If AV_W <> 0 Then

Print #2, Tab(5); msgtab(312); Tab(75); LTrim(RTrim(Format(PROP * FINAL_CATCH / AV_W, "##########0")))
Write #3, msgtab(312); Int(PROP * FINAL_CATCH / AV_W); " "; " "; " "

End If

Print #2, " "
Write #3, " "; " "; " "; " "; " "

Print #2, Tab(5); msgtab(313)
Print #2, Tab(5); String(92, "=")

Write #3, msgtab(313); " "; " "; " "; " "
Write #3, " "; " "; " "; " "; " "

EXIT_SUB:

End Sub
Private Sub FRAME_MINOR_TOTALS2()

Dim FSIC(1 To 1000), FMNC(1 To 200), FMNN(1 To 200), FBGC(1 To 200), FBGN(1 To 200)
Dim FMNBG(1 To 200, 1 To 200)
Dim I, J, K, L, XXX, fnm, V, W

' Load boat/gear table

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_BG.TXT"

Open fnm For Input As #5

For I = 1 To 200
FBGN(I) = " "
Next I

Do While Not EOF(5)
Line Input #5, XXX
J = Val(Left(XXX, 4))
FBGN(J) = Mid(XXX, 6, 30)
Loop

Close #5

' Load MN table
'===============

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_ASSOSI.TXT"

If Dir(fnm) = "" Then Exit Sub

Open fnm For Input As #5

For I = 1 To 200
FMNN(I) = " "
Next I

For I = 1 To 200
For J = 1 To 200
FMNBG(I, J) = 0
Next J
Next I

Do While Not EOF(5)

Line Input #5, XXX

I = Val(Mid(XXX, 1, 4)): FMNN(I) = Mid(XXX, 6, 30): K = Val(Mid(XXX, 37, 4))

For J = 1 To K
Line Input #5, XXX
L = Val(Mid(XXX, 6, 4))
FSIC(L) = I
Next J

Loop

Close #5

'========================================
' READ FRAME

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(current_year, "0000") + _
      "M" + Format(current_month, "00") + "_FRAME.TXT"

Open fnm For Input As #5

Do Until EOF(5)

Line Input #5, XXX

L = Val(Mid(XXX, 2, 4)): J = Val(Mid(XXX, 8, 4)): V = CDbl(Mid(XXX, 26, 15))

W = Format(V, "#####0.00")

I = FSIC(L)

FMNBG(I, J) = FMNBG(I, J) + V

DOCTMNBG(I, J) = FMNBG(I, J)

Loop

Close #5

End Sub

