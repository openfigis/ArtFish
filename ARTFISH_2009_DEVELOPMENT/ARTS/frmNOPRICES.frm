VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmNOPRICES 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7425
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   11910
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7425
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRETURN 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   11040
      MousePointer    =   1  'Arrow
      Picture         =   "frmNOPRICES.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6600
      Width           =   735
   End
   Begin VB.CommandButton cmdGUIDE 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   9360
      MousePointer    =   1  'Arrow
      Picture         =   "frmNOPRICES.frx":0282
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6600
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
      Left            =   10200
      MousePointer    =   1  'Arrow
      Picture         =   "frmNOPRICES.frx":24E4
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6600
      Width           =   735
   End
   Begin RichTextLib.RichTextBox rtsNOPRICE 
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   5530
      _Version        =   393217
      BackColor       =   12648447
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      MousePointer    =   1
      TextRTF         =   $"frmNOPRICES.frx":2766
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox rtsNOFISH 
      Height          =   3255
      Left            =   0
      TabIndex        =   4
      Top             =   3240
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   5741
      _Version        =   393217
      BackColor       =   12648447
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      MousePointer    =   1
      TextRTF         =   $"frmNOPRICES.frx":27E6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FFFF&
      Caption         =   "05"
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   7200
      Width           =   255
   End
End
Attribute VB_Name = "frmNOPRICES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FNM5, FNM6

Private NMJ, NMN, NBG, NSP
Private TMJC(), TMJN(), TMNC(), TMNN(), TBGC(), TBGN(), TSPC(), TSPN(), TASSO()
Private SELMJ(), SELMN(), SELBG(), SELSP()
Private REPJ(), REPM(), REPB(), REPS()
Private Sub cmdPRINT_Click()

Printer.FontBold = True
Printer.FontName = "Courier"
Printer.FontName = "Courier New"
Printer.FontSize = 9
Printer.FontItalic = False

Dim XXX

If Dir(FNM5) = "" Then GoTo CONT_FISH

Printer.Print Tab(2); frmNOPRICES.Caption
Printer.Print " "

Open FNM5 For Input As #1

Do Until EOF(1)

Line Input #1, XXX

Printer.Print XXX

Loop

Close #1

CONT_FISH:

If Dir(FNM6) = "" Then
   Printer.EndDoc
   Exit Sub
   End If

Open FNM6 For Input As #1

Printer.NewPage

Printer.Print Tab(2); frmNOPRICES.Caption
Printer.Print " "

Do Until EOF(1)

Line Input #1, XXX

Printer.Print XXX

Loop

Close #1

Printer.EndDoc

End Sub
Private Sub Form_Load()

cmdGUIDE.Visible = False

Set Picture = LoadPicture(APPROOT + "\ARTS\PICS_RUNTIME\SCREEN_04.JPG")

FNM5 = APPROOT + "\ARTS\CONTROL\NOPRICES.TXT"
FNM6 = APPROOT + "\ARTS\CONTROL\NOFISH.TXT"

If Dir(FNM5) = "" And Dir(FNM6) = "" Then Call cmdRETURN_Click

If Dir(FNM5) <> "" Then
   rtsNOPRICE.FileName = FNM5
   rtsNOPRICE.Refresh
   End If
 
If Dir(FNM6) <> "" Then
   rtsNOFISH.FileName = FNM6
   rtsNOFISH.Refresh
   End If
 
frmNOPRICES.Caption = msgtab(15) + ": " + Format(CURY, "0000") + " - " + _
                 msgtab(54)

cmdRETURN.ToolTipText = msgtab(113)
cmdGUIDE.ToolTipText = msgtab(6)
cmdPRINT.ToolTipText = msgtab(110)

End Sub
Private Sub cmdGUIDE_Click()

HTYPE = "50"

HFNM = APPROOT + "\ARTS\HELP\" + current_language + "HELP" + HTYPE + ".rtf"

If Dir(HFNM) = "" Then Exit Sub

frmNOPRICES.Enabled = False
Load frmGUIDE
frmGUIDE.Show

End Sub
Private Sub cmdRETURN_Click()

'Load frmREPORTS
Unload frmNOPRICES
'frmREPORTS.Show

End Sub
Private Sub PRINT_NOPRICE()


End Sub


