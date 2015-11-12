VERSION 5.00
Begin VB.Form frmPLOT 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11880
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   1  'Arrow
   Moveable        =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdKPRICE 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   10560
      Picture         =   "frmPLOT.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3120
      Width           =   255
   End
   Begin VB.CommandButton cmdKVALUE 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   10560
      Picture         =   "frmPLOT.frx":010A
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2280
      Width           =   255
   End
   Begin VB.CommandButton cmdVALUE 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10920
      Picture         =   "frmPLOT.frx":0214
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1800
      Width           =   735
   End
   Begin VB.CommandButton cmdPRICE 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10920
      Picture         =   "frmPLOT.frx":080E
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton cmdKCPUE 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   10560
      Picture         =   "frmPLOT.frx":0E08
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3960
      Width           =   255
   End
   Begin VB.CommandButton cmdCPUE 
      BackColor       =   &H00FFFFFF&
      Caption         =   "CPUE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10920
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3480
      Width           =   735
   End
   Begin VB.CommandButton cmdKEFFORT 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   10560
      Picture         =   "frmPLOT.frx":0F12
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4800
      Width           =   255
   End
   Begin VB.CommandButton cmdKSPECIES 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   10560
      Picture         =   "frmPLOT.frx":101C
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5640
      Width           =   255
   End
   Begin VB.CommandButton cmdBG 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   10920
      Picture         =   "frmPLOT.frx":1126
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4320
      Width           =   735
   End
   Begin VB.CommandButton cmdSPECIES 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   10920
      Picture         =   "frmPLOT.frx":13A8
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5160
      Width           =   735
   End
   Begin VB.CommandButton cmdEXIT 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   11400
      MousePointer    =   1  'Arrow
      Picture         =   "frmPLOT.frx":1EAA
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7020
      Width           =   375
   End
   Begin VB.CommandButton cmdRETURN 
      BackColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   10800
      MousePointer    =   1  'Arrow
      Picture         =   "frmPLOT.frx":212C
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6420
      Width           =   975
   End
   Begin VB.Label lblTV 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   47
      Top             =   1080
      Width           =   10215
   End
   Begin VB.Label lblTP 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   46
      Top             =   840
      Width           =   10215
   End
   Begin VB.Label lblTU 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   45
      Top             =   600
      Width           =   10215
   End
   Begin VB.Label lblTE 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   44
      Top             =   360
      Width           =   10215
   End
   Begin VB.Label lblTC 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   43
      Top             =   120
      Width           =   10215
   End
   Begin VB.Label lblID 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   120
      TabIndex        =   42
      Top             =   7440
      Width           =   11535
   End
   Begin VB.Label lblPRICE 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Price"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   4200
      TabIndex        =   41
      Top             =   5520
      Width           =   1815
   End
   Begin VB.Label lblPRICE 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Price"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   4200
      TabIndex        =   40
      Top             =   5160
      Width           =   1815
   End
   Begin VB.Label lblPRICE 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Price"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   4200
      TabIndex        =   39
      Top             =   4800
      Width           =   1815
   End
   Begin VB.Label lblPRICE 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Price"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   4200
      TabIndex        =   38
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Label lblPRICE 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Price"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   4200
      TabIndex        =   37
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Label lblVALUE 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Value"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   2280
      TabIndex        =   36
      Top             =   5520
      Width           =   1815
   End
   Begin VB.Label lblVALUE 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Value"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   2280
      TabIndex        =   35
      Top             =   5160
      Width           =   1815
   End
   Begin VB.Label lblVALUE 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Value"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   2280
      TabIndex        =   34
      Top             =   4800
      Width           =   1815
   End
   Begin VB.Label lblVALUE 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Value"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   2280
      TabIndex        =   33
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Label lblVALUE 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Value"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   2280
      TabIndex        =   32
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Label lblCPUE 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "CPUE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   6120
      TabIndex        =   31
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Label lblCPUE 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "CPUE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   6120
      TabIndex        =   30
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label lblCPUE 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "CPUE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   6120
      TabIndex        =   29
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label lblCPUE 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "CPUE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   6120
      TabIndex        =   28
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label lblCPUE 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "CPUE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   6120
      TabIndex        =   27
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label lblEFFORT 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Effort"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   4200
      TabIndex        =   26
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Label lblEFFORT 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Effort"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   4200
      TabIndex        =   25
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label lblEFFORT 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Effort"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   4200
      TabIndex        =   24
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label lblEFFORT 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Effort"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   4200
      TabIndex        =   23
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label lblEFFORT 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Effort"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   4200
      TabIndex        =   22
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label lblCATCH 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Catch"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   2280
      TabIndex        =   21
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Label lblCATCH 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Catch"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   2280
      TabIndex        =   20
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label lblCATCH 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Catch"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   2280
      TabIndex        =   19
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label lblCATCH 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Catch"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   2280
      TabIndex        =   18
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label lblCATCH 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Catch"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   2280
      TabIndex        =   17
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label lblV 
      Appearance      =   0  'Flat
      BackColor       =   &H00C000C0&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   11640
      TabIndex        =   16
      Top             =   1800
      Width           =   135
   End
   Begin VB.Label lblP 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   11640
      TabIndex        =   15
      Top             =   2640
      Width           =   135
   End
   Begin VB.Label lblU 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   11640
      TabIndex        =   14
      Top             =   3480
      Width           =   135
   End
   Begin VB.Label lblE 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   11640
      TabIndex        =   13
      Top             =   4320
      Width           =   135
   End
   Begin VB.Label lblC 
      Appearance      =   0  'Flat
      BackColor       =   &H000000C0&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   11640
      TabIndex        =   12
      Top             =   5160
      Width           =   135
   End
End
Attribute VB_Name = "frmPLOT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private TCZ(), TEZ(), TUZ(), TPZ(), TVZ(), TWZ()
Private MINC, MAXC, MINE, MAXE, MINU, MAXU, MINP, MAXP, MINV, MAXV, MINW, MAXW
Private UPLX, UPLY, UPRX, UPRY, DNLX, DNLY, DNRX, DNRY
Private INP_MIN, INP_MAX, RES_MIN, RES_MAX
Private Sub cmdBG_Click()

If MINE = MAXE Then
   cmdKEFFORT.Visible = False
   cmdBG.Enabled = False
   Exit Sub
   End If

If cmdKEFFORT.Visible = True Then
   cmdKEFFORT.Visible = False
   Call DECIDE_PLOT
   Exit Sub
   End If

If cmdKEFFORT.Visible = False Then
   cmdKEFFORT.Visible = True
   Call DECIDE_PLOT
   Exit Sub
   End If

End Sub
Private Sub cmdCPUE_Click()

If MINU = MAXU Then
   cmdKCPUE.Visible = False
   cmdCPUE.Enabled = False
   Exit Sub
   End If

If cmdKCPUE.Visible = True Then
   cmdKCPUE.Visible = False
   Call DECIDE_PLOT
   Exit Sub
   End If

If cmdKCPUE.Visible = False Then
   cmdKCPUE.Visible = True
   Call DECIDE_PLOT
   Exit Sub
   End If

End Sub
Private Sub cmdEXIT_Click()

Call write_parms

End

End Sub
Private Sub cmdPRICE_Click()

If MINW = MAXW Then
   cmdKPRICE.Visible = False
   cmdPRICE.Enabled = False
   Exit Sub
   End If

If cmdKPRICE.Visible = True Then
   cmdKPRICE.Visible = False
   Call DECIDE_PLOT
   Exit Sub
   End If

If cmdKPRICE.Visible = False Then
   cmdKPRICE.Visible = True
   Call DECIDE_PLOT
   Exit Sub
   End If

End Sub

Private Sub cmdRETURN_Click()

Unload frmPLOT
frmREPORTS.Enabled = True
frmREPORTS.Show

End Sub
Private Sub cmdSPECIES_Click()

If MINC = MAXC Then
   cmdKSPECIES.Visible = False
   cmdSPECIES.Enabled = False
   Exit Sub
   End If

If cmdKSPECIES.Visible = True Then
   cmdKSPECIES.Visible = False
   Call DECIDE_PLOT
   Exit Sub
   End If

If cmdKSPECIES.Visible = False Then
   cmdKSPECIES.Visible = True
   Call DECIDE_PLOT
   Exit Sub
   End If

End Sub
Private Sub PLOT_FRAME()

DrawWidth = 2
ForeColor = vbWhite
FillColor = vbWhite
FillStyle = 0

UPLX = 120: UPLY = 1400: UPRX = 10320: UPRY = 1400
DNLX = 120: DNLY = 7400: DNRX = 10320: DNRY = 7400

Line (UPLX, UPLY)-(DNRX, DNRY), , BF

ForeColor = vbBlack
DrawWidth = 1

Line (UPLX, UPLY)-(UPRX, UPRY)
Line (UPRX, UPRY)-(DNRX, DNRY)
Line (DNRX, DNRY)-(DNLX, DNLY)
Line (DNLX, DNLY)-(UPLX, UPLY)

UPLX = 2280: UPLY = 1920: UPRX = 8160: UPRY = 1920
DNLX = 2280: DNLY = 6840: DNRX = 8160: DNRY = 6840

ForeColor = vbBlack
DrawWidth = 2

Line (UPLX, UPLY)-(UPRX, UPRY)
Line (UPRX, UPRY)-(DNRX, DNRY)
Line (DNRX, DNRY)-(DNLX, DNLY)
Line (DNLX, DNLY)-(UPLX, UPLY)

DrawWidth = 1

'PLOT Y_MARKINGS
'===============

Dim DY, I, YY

DY = DNLY - UPLY

For I = 0 To 4

YY = UPLY + Int(I * DY / 4)

Line (UPLX - 120, YY)-(UPRX + 120, YY)

Next I

'PLOT X_MARKINGS
'===============

Font.Size = 10
FontBold = True

Dim DX, XX

DX = DNRX - DNLX

For I = 0 To 12

XX = DNLX + Int(I * DX / 11)

If I < 12 Then Line (XX, DNRY)-(XX, DNRY + 120)

CurrentX = XX - 90: CurrentY = DNRY + 120

If I <= 11 Then
   If VALID_MONTHS(I + 1) <> "X" Then GoTo NEXT_I
   End If

If I <= 11 Then Print Format(I + 1, "00")

NEXT_I:

Next I

End Sub
Private Sub NORM_DATA_C()

If REPTC(13) = 0 Then
   MINC = 0: MAXC = 0
   Exit Sub
   End If

Dim I

ReDim TCZ(1 To 12)

'NORMALIZE CATCH
'===============

MINC = REPTC(1): MAXC = REPTC(1)

For I = 1 To 12

TCZ(I) = 0

If VALID_MONTHS(I) <> "X" Then GoTo NEXT_IC:

If REPTC(I) < MINC Then MINC = REPTC(I)
If REPTC(I) > MAXC Then MAXC = REPTC(I)

NEXT_IC:

Next I

INP_MIN = MINC

Call FIND_MINSCALE

MINC = Int(RES_MIN)

INP_MAX = MAXC

Call FIND_MAXSCALE

MAXC = Int(RES_MAX)

For I = 1 To 12

If VALID_MONTHS(I) <> "X" Then GoTo NEXT_IC2:

If (MAXC - MINC) <> 0 Then TCZ(I) = (REPTC(I) - MINC) / (MAXC - MINC)

NEXT_IC2:

Next I

End Sub
Private Sub NORM_DATA_V()

If REPTV(13) = 0 Then
   MINV = 0: MAXV = 0
   Exit Sub
   End If

Dim I

ReDim TVZ(1 To 12)

'NORMALIZE VALUE
'===============

MINV = REPTV(1): MAXV = REPTV(1)

For I = 1 To 12

TVZ(I) = 0

If VALID_MONTHS(I) <> "X" Then GoTo NEXT_IV:

If REPTV(I) < MINV Then MINV = REPTV(I)
If REPTV(I) > MAXV Then MAXV = REPTV(I)

NEXT_IV:

Next I

INP_MIN = MINV

Call FIND_MINSCALE

MINV = Int(RES_MIN)

INP_MAX = MAXV

Call FIND_MAXSCALE

MAXV = Int(RES_MAX)

For I = 1 To 12

If VALID_MONTHS(I) <> "X" Then GoTo NEXT_IV2:

If (MAXV - MINV) <> 0 Then TVZ(I) = (REPTV(I) - MINV) / (MAXV - MINV)

NEXT_IV2:

Next I

End Sub
Private Sub NORM_DATA_E()

If REPTE(13) = 0 Then
   MINE = 0: MAXE = 0
   Exit Sub
   End If

Dim I

ReDim TEZ(1 To 12)

'NORMALIZE EFFORT
'================

MINE = REPTE(1): MAXE = REPTE(1)

For I = 1 To 12

TEZ(I) = 0

If VALID_MONTHS(I) <> "X" Then GoTo NEXT_IE:

If REPTE(I) < MINE Then MINE = REPTE(I)
If REPTE(I) > MAXE Then MAXE = REPTE(I)

NEXT_IE:

Next I

INP_MIN = MINE

Call FIND_MINSCALE

MINE = Int(RES_MIN)

INP_MAX = MAXE

Call FIND_MAXSCALE

MAXE = Int(RES_MAX)

For I = 1 To 12

If VALID_MONTHS(I) <> "X" Then GoTo NEXT_IE2:

If (MAXE - MINE) <> 0 Then TEZ(I) = (REPTE(I) - MINE) / (MAXE - MINE)

NEXT_IE2:

Next I

End Sub
Private Sub NORM_DATA_U()

If REPTU(13) = 0 Then
   MINU = 0: MAXU = 0
   Exit Sub
   End If

Dim I

ReDim TUZ(1 To 12)

'NORMALIZE CPUE
'===============

MINU = REPTU(1): MAXU = REPTU(1)

For I = 1 To 12

TUZ(I) = 0

If VALID_MONTHS(I) <> "X" Then GoTo NEXT_IU:

If REPTU(I) < MINU Then MINU = REPTU(I)
If REPTU(I) > MAXU Then MAXU = REPTU(I)

NEXT_IU:

Next I

INP_MIN = Int(MINU)

Call FIND_MINSCALE

MINU = Int(RES_MIN)

INP_MAX = Int(MAXU) + 1

Call FIND_MAXSCALE

MAXU = Int(RES_MAX)

For I = 1 To 12

If VALID_MONTHS(I) <> "X" Then GoTo NEXT_IU2:

If (MAXU - MINU) <> 0 Then TUZ(I) = (REPTU(I) - MINU) / (MAXU - MINU)

NEXT_IU2:

Next I


End Sub
Private Sub NORM_DATA_P()

If REPTP(13) = 0 Then
   MINP = 0: MAXP = 0
   Exit Sub
   End If

Dim I

ReDim TPZ(1 To 12)

'NORMALIZE PRICE
'===============

MINP = REPTP(1): MAXP = REPTP(1)

For I = 1 To 12

TPZ(I) = 0

If VALID_MONTHS(I) <> "X" Then GoTo NEXT_IP:

If REPTP(I) < MINP Then MINP = REPTP(I)
If REPTP(I) > MAXP Then MAXP = REPTP(I)

NEXT_IP:

Next I

INP_MIN = Int(MINP)

Call FIND_MINSCALE

MINP = Int(RES_MIN)

INP_MAX = Int(MAXP) + 1

Call FIND_MAXSCALE

MAXP = Int(RES_MAX)

For I = 1 To 12

If VALID_MONTHS(I) <> "X" Then GoTo NEXT_IP2:

If (MAXP - MINP) <> 0 Then TPZ(I) = (REPTP(I) - MINP) / (MAXP - MINP)

NEXT_IP2:

Next I

End Sub
Private Sub NORM_DATA_W()

If REPTW(13) = 0 Then
   MINW = 0: MAXW = 0
   Exit Sub
   End If

Dim I

ReDim TWZ(1 To 12)

'NORMALIZE WEIGHT
'===============

MINW = 999999999999#: MAXW = -999999999999#

For I = 1 To 12

TWZ(I) = 0

If VALID_MONTHS(I) <> "X" Then GoTo NEXT_IW:

If REPTW(I) < MINW Then MINW = REPTW(I)
If REPTW(I) > MAXW Then MAXW = REPTW(I)

NEXT_IW:

Next I

INP_MIN = Int(MINW)

Call FIND_MINSCALE

'MINW = Int(RES_MIN)

INP_MAX = Int(MAXW) + 1

Call FIND_MAXSCALE

MINW = 0
'MAXW = Int(RES_MAX)

Dim MAXAUX

MAXAUX = MAXW

If MAXAUX < 0.1 Then MAXW = 0.1
If MAXAUX >= 0.1 And MAXAUX < 0.2 Then MAXW = 0.2
If MAXAUX >= 0.2 And MAXAUX < 0.3 Then MAXW = 0.3
If MAXAUX >= 0.3 And MAXAUX < 0.4 Then MAXW = 0.4
If MAXAUX >= 0.4 And MAXAUX < 0.5 Then MAXW = 0.5
If MAXAUX >= 0.5 And MAXAUX < 1 Then MAXW = 1
If MAXAUX >= 1 Then MAXW = Int(RES_MAX)


For I = 1 To 12

If VALID_MONTHS(I) <> "X" Then GoTo NEXT_IW2:

If (MAXW - MINW) <> 0 Then TWZ(I) = (REPTW(I) - MINW) / (MAXW - MINW)

NEXT_IW2:

Next I

End Sub
Private Sub cmdVALUE_Click()

If MINP = MAXP Then
   cmdKVALUE.Visible = False
   cmdVALUE.Enabled = False
   Exit Sub
   End If

If cmdKVALUE.Visible = True Then
   cmdKVALUE.Visible = False
   Call DECIDE_PLOT
   Exit Sub
   End If

If cmdKVALUE.Visible = False Then
   cmdKVALUE.Visible = True
   Call DECIDE_PLOT
   Exit Sub
   End If

End Sub
Private Sub Form_Load()

frmPLOT.Caption = msgtab(15) + " : " + Format(CURY, "0000") + _
                " - " + msgtab(90)

cmdEXIT.ToolTipText = msgtab(3)
cmdRETURN.ToolTipText = msgtab(91)

cmdSPECIES.ToolTipText = msgtab(63)
cmdBG.ToolTipText = msgtab(64)
cmdCPUE.ToolTipText = msgtab(65)
cmdPRICE.ToolTipText = msgtab(111)
cmdVALUE.ToolTipText = msgtab(66)

lblID = RTrim(REPMJN) + " " + RTrim(REPMNN) + " " + RTrim(REPBGN) + " " + RTrim(REPSPN)

Call NORM_DATA_C
Call NORM_DATA_E
Call NORM_DATA_U
'Call NORM_DATA_V
Call NORM_DATA_P
Call NORM_DATA_W

If MINC + MAXC = 0 Then
   cmdSPECIES.Visible = False
   lblC.Visible = False
   End If
   
If MINE + MAXE = 0 Then
   cmdBG.Visible = False
   lblE.Visible = False
   End If
   
If MINU + MAXU = 0 Then
   cmdCPUE.Visible = False
   lblU.Visible = False
   End If
   
If MINP + MAXP = 0 Then
   cmdVALUE.Visible = False
   lblV.Visible = False
   End If
   
If MINW + MAXW = 0 Then
   cmdPRICE.Enabled = False
   lblP.Visible = False
   End If

cmdPRICE.Enabled = True

If AW_OPTION = 0 Then
   cmdPRICE.Enabled = False
   End If

cmdKSPECIES.Visible = False
cmdKEFFORT.Visible = False
cmdKCPUE.Visible = False
cmdKPRICE.Visible = False
cmdKVALUE.Visible = False

lblTC.Visible = False
lblTE.Visible = False
lblTU.Visible = False
lblTP.Visible = False
lblTV.Visible = False

lblTC.ForeColor = lblC.BackColor
lblTE.ForeColor = lblE.BackColor
lblTU.ForeColor = lblU.BackColor
lblTP.ForeColor = lblP.BackColor
lblTV.ForeColor = lblV.BackColor

lblTC.Caption = msgtab(63) + " (" + UNW + ")"
lblTE.Caption = msgtab(64) + " (" + msgtab(71) + ")"
lblTU.Caption = msgtab(65) + " (" + UNW + "/" + msgtab(71) + ")"
lblTP.Caption = msgtab(111)
lblTV.Caption = msgtab(66) + " (" + UNV + "/" + UNW + ")"

Dim I

For I = 0 To 4

lblCATCH(I).Visible = False
lblEFFORT(I).Visible = False
lblCPUE(I).Visible = False
lblVALUE(I).Visible = False
lblPRICE(I).Visible = False

Next I

End Sub
Private Sub DECIDE_PLOT()

lblTC.Visible = False
lblTE.Visible = False
lblTU.Visible = False
lblTP.Visible = False
lblTV.Visible = False

Call PLOT_FRAME

If cmdKSPECIES.Visible = True Then
   lblTC.Visible = True
   Call PLOT_SCALE_C
   Call PLOT_CATCH
   End If
   
If cmdKEFFORT.Visible = True Then
   lblTE.Visible = True
   Call PLOT_SCALE_E
   Call PLOT_EFFORT
   End If
 
If cmdKCPUE.Visible = True Then
   lblTU.Visible = True
   Call PLOT_SCALE_U
   Call PLOT_CPUE
   End If

If cmdKVALUE.Visible = True Then
   lblTV.Visible = True
   Call PLOT_SCALE_P
   Call PLOT_PRICE
   End If

If cmdKPRICE.Visible = True Then
   lblTP.Visible = True
   Call PLOT_SCALE_W
   Call PLOT_WEIGHT
   End If

End Sub
Private Sub PLOT_CATCH()

Dim CCXX(), CCYY()

ReDim CCXX(1 To 12), CCYY(1 To 12)

ForeColor = lblC.BackColor
DrawWidth = 3

Dim I

For I = 1 To 12

If VALID_MONTHS(I) <> "X" Then GoTo NEXT_I

CurrentX = DNLX + Int((I - 1) * (DNRX - DNLX) / 11)
CurrentY = DNLY - (DNLY - UPLY) * (REPTC(I) - MINC) / (MAXC - MINC)

Circle (CurrentX, CurrentY), 50

CCXX(I) = CurrentX: CCYY(I) = CurrentY

NEXT_I:

Next I

DrawWidth = 2

For I = 1 To 11

If VALID_MONTHS(I) <> "X" Or VALID_MONTHS(I + 1) <> "X" Then GoTo NEXT_I2

Line (CCXX(I), CCYY(I))-(CCXX(I + 1), CCYY(I + 1))

NEXT_I2:

Next I

End Sub
Private Sub PLOT_VALUE()

Dim CCXX(), CCYY()

ReDim CCXX(1 To 12), CCYY(1 To 12)

ForeColor = lblV.BackColor
DrawWidth = 3

Dim I

For I = 1 To 12

If VALID_MONTHS(I) <> "X" Then GoTo NEXT_I

CurrentX = DNLX + Int((I - 1) * (DNRX - DNLX) / 11)
CurrentY = DNLY - (DNLY - UPLY) * (REPTV(I) - MINV) / (MAXV - MINV)

Circle (CurrentX, CurrentY), 50

CCXX(I) = CurrentX: CCYY(I) = CurrentY

NEXT_I:

Next I

DrawWidth = 2

For I = 1 To 11

If VALID_MONTHS(I) <> "X" Or VALID_MONTHS(I + 1) <> "X" Then GoTo NEXT_I2

Line (CCXX(I), CCYY(I))-(CCXX(I + 1), CCYY(I + 1))

NEXT_I2:

Next I


End Sub
Private Sub PLOT_PRICE()

Dim CCXX(), CCYY()

ReDim CCXX(1 To 12), CCYY(1 To 12)

ForeColor = lblV.BackColor
DrawWidth = 3

Dim I

For I = 1 To 12

If VALID_MONTHS(I) <> "X" Then GoTo NEXT_I

CurrentX = DNLX + Int((I - 1) * (DNRX - DNLX) / 11)
CurrentY = DNLY - (DNLY - UPLY) * (REPTP(I) - MINP) / (MAXP - MINP)

Circle (CurrentX, CurrentY), 50

CCXX(I) = CurrentX: CCYY(I) = CurrentY

NEXT_I:

Next I

DrawWidth = 2

For I = 1 To 11

If VALID_MONTHS(I) <> "X" Or VALID_MONTHS(I + 1) <> "X" Then GoTo NEXT_I2

Line (CCXX(I), CCYY(I))-(CCXX(I + 1), CCYY(I + 1))

NEXT_I2:

Next I

End Sub
Private Sub PLOT_WEIGHT()

Dim CCXX(), CCYY()

ReDim CCXX(1 To 12), CCYY(1 To 12)

ForeColor = lblP.BackColor
DrawWidth = 3

Dim I

For I = 1 To 12

If VALID_MONTHS(I) <> "X" Then GoTo NEXT_I

CurrentX = DNLX + Int((I - 1) * (DNRX - DNLX) / 11)
CurrentY = DNLY - (DNLY - UPLY) * (REPTW(I) - MINW) / (MAXW - MINW)

If REPTW(I) <> 0 Then Circle (CurrentX, CurrentY), 50

CCXX(I) = CurrentX: CCYY(I) = CurrentY

NEXT_I:

Next I

DrawWidth = 2

For I = 1 To 11

If VALID_MONTHS(I) <> "X" Or VALID_MONTHS(I + 1) <> "X" Then GoTo NEXT_I2

If REPTW(I) <> 0 And REPTW(I + 1) <> 0 Then
Line (CCXX(I), CCYY(I))-(CCXX(I + 1), CCYY(I + 1))
End If

NEXT_I2:

Next I

End Sub
Private Sub PLOT_EFFORT()

Dim CCXX(), CCYY()

ReDim CCXX(1 To 12), CCYY(1 To 12)

ForeColor = lblE.BackColor
DrawWidth = 3

Dim I

For I = 1 To 12

If VALID_MONTHS(I) <> "X" Then GoTo NEXT_I

CurrentX = DNLX + Int((I - 1) * (DNRX - DNLX) / 11)
CurrentY = DNLY - (DNLY - UPLY) * (REPTE(I) - MINE) / (MAXE - MINE)

Circle (CurrentX, CurrentY), 50

CCXX(I) = CurrentX: CCYY(I) = CurrentY

NEXT_I:

Next I

DrawWidth = 2

For I = 1 To 11

If VALID_MONTHS(I) <> "X" Or VALID_MONTHS(I + 1) <> "X" Then GoTo NEXT_I2

Line (CCXX(I), CCYY(I))-(CCXX(I + 1), CCYY(I + 1))

NEXT_I2:

Next I

End Sub
Private Sub PLOT_CPUE()

Dim CCXX(), CCYY()

ReDim CCXX(1 To 12), CCYY(1 To 12)

ForeColor = lblU.BackColor
DrawWidth = 3

Dim I

For I = 1 To 12

If VALID_MONTHS(I) <> "X" Then GoTo NEXT_I

CurrentX = DNLX + Int((I - 1) * (DNRX - DNLX) / 11)
CurrentY = DNLY - (DNLY - UPLY) * (REPTU(I) - MINU) / (MAXU - MINU)

Circle (CurrentX, CurrentY), 50

CCXX(I) = CurrentX: CCYY(I) = CurrentY

NEXT_I:

Next I

DrawWidth = 2

For I = 1 To 11

If VALID_MONTHS(I) <> "X" Or VALID_MONTHS(I + 1) <> "X" Then GoTo NEXT_I2

Line (CCXX(I), CCYY(I))-(CCXX(I + 1), CCYY(I + 1))

NEXT_I2:

Next I

End Sub
Private Sub PLOT_SCALE_C()

Dim DY, I, YY, DD

DD = MAXC - MINC

DY = DNLY - UPLY

For I = 0 To 4

YY = UPLY + Int(I * DY / 4)

lblCATCH(I).Top = YY - 100
lblCATCH(I).Left = UPLX - 150 - lblCATCH(I).Width
lblCATCH(I).ForeColor = lblC.BackColor
lblCATCH(I).Caption = MAXC - Int(I * DD / 4)
lblCATCH(I).Visible = True

Next I

End Sub
Private Sub PLOT_SCALE_V()

Dim DY, I, YY, DD

DD = MAXV - MINV

DY = DNLY - UPLY

For I = 0 To 4

YY = UPLY + Int(I * DY / 4)

lblVALUE(I).Top = YY - 100
lblVALUE(I).Left = UPRX + 100
lblVALUE(I).ForeColor = lblV.BackColor
lblVALUE(I).Caption = MAXV - Int(I * DD / 4)
lblVALUE(I).Visible = True

Next I

End Sub
Private Sub PLOT_SCALE_P()

Dim DY, I, YY, DD

DD = MAXP - MINP

DY = DNLY - UPLY

For I = 0 To 4

YY = UPLY + Int(I * DY / 4)

lblVALUE(I).Left = UPRX + 100
lblVALUE(I).Top = YY - 100
lblVALUE(I).Left = UPRX + 100
lblVALUE(I).ForeColor = lblV.BackColor
lblVALUE(I).Caption = Format(MAXP - (I * DD / 4), "########0.000")
lblVALUE(I).Visible = True

Next I

End Sub
Private Sub PLOT_SCALE_W()

Dim DY, I, YY, DD

DD = MAXW - MINW

DY = DNLY - UPLY

For I = 0 To 4

YY = UPLY + Int(I * DY / 4)

lblPRICE(I).Top = YY - lblVALUE(I).Height - 90
lblPRICE(I).Left = UPRX + 100
lblPRICE(I).ForeColor = lblP.BackColor
lblPRICE(I).Caption = Format(MAXW - (I * DD / 4), "########0.000")
lblPRICE(I).Visible = True

Next I

End Sub
Private Sub PLOT_SCALE_E()

Dim DY, I, YY, DD

DD = MAXE - MINE

DY = DNLY - UPLY

For I = 0 To 4

YY = UPLY + Int(I * DY / 4)

lblEFFORT(I).Top = YY - 100 - lblCATCH(I).Height
lblEFFORT(I).Left = UPLX - 150 - lblEFFORT(I).Width
lblEFFORT(I).ForeColor = lblE.BackColor
lblEFFORT(I).Caption = MAXE - Int(I * DD / 4)
lblEFFORT(I).Visible = True

Next I

End Sub
Private Sub PLOT_SCALE_U()

Dim DY, I, YY, DD

DD = MAXU - MINU

DY = DNLY - UPLY

For I = 0 To 4

YY = UPLY + Int(I * DY / 4)

lblCPUE(I).Top = YY + lblCATCH(I).Height - 100
lblCPUE(I).Left = UPLX - 150 - lblCPUE(I).Width
lblCPUE(I).ForeColor = lblU.BackColor
lblCPUE(I).Caption = Format(MAXU - (I * DD / 4), "########0.000")
lblCPUE(I).Visible = True

Next I

End Sub
Private Sub FIND_MINSCALE()

Dim X, I, J, Y, II, JJ

X = INP_MIN

If X < 1 Then
   RES_MIN = 0
   Exit Sub
   End If

For I = 1 To 20

Y = X / 10 ^ I

If Y < 1 Then
   II = I - 1
   GoTo CONT_PROC
   End If

Next I

CONT_PROC:

For J = 1 To 10

If J * 10 ^ II >= X Then
   JJ = J - 1
   GoTo CONT_PROC2:
   End If

Next J

CONT_PROC2:

RES_MIN = JJ * 10 ^ II
   
End Sub
Private Sub FIND_MAXSCALE()

Dim X, I, J, Y, II, JJ

X = INP_MAX

If X < 1 Then
   RES_MAX = 1
   Exit Sub
   End If

For I = 1 To 20

Y = X / 10 ^ I

If Y < 1 Then
   II = I - 1
   GoTo CONT_PROC
   End If

Next I

CONT_PROC:

For J = 1 To 10

If J * 10 ^ II >= X Then
   JJ = J
   GoTo CONT_PROC2:
   End If

Next J

CONT_PROC2:

RES_MAX = JJ * 10 ^ II

End Sub

