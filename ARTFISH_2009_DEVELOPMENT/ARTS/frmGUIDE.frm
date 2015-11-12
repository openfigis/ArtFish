VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmGUIDE 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5115
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   5115
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox rtsGUIDE 
      CausesValidation=   0   'False
      Height          =   3375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   5953
      _Version        =   393217
      BackColor       =   16777215
      BorderStyle     =   0
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmGUIDE.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmGUIDE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()

frmGUIDE.Caption = msgtab(6)

frmGUIDE.MousePointer = 1

HFNM = APPROOT + "\ARTS\HELP\" + current_language + "HELP" + HTYPE + ".rtf"

rtsGUIDE.FileName = HFNM
rtsGUIDE.Refresh

End Sub
Private Sub rtsGUIDE_Click()

'rtsGUIDE.filename = HFNM
rtsGUIDE.Refresh

End Sub
Private Sub rtsGUIDE_dblClick()

'rtsGUIDE.SaveFile HFNM

If Left(HTYPE, 1) = 1 Then

Unload frmGUIDE
frmARTS00.Enabled = True
frmARTS00.Show
frmARTS00.Refresh
Exit Sub

End If

If Left(HTYPE, 1) = 2 Then

Unload frmGUIDE
frmSEL.Enabled = True
frmSEL.Show
frmSEL.Refresh

End If

If Left(HTYPE, 1) = 3 Then

Unload frmGUIDE
frmCOMP.Enabled = True
frmCOMP.Show
frmCOMP.Refresh

End If

If Left(HTYPE, 1) = 4 Then

Unload frmGUIDE
frmREPORTS.Enabled = True
frmREPORTS.Show
frmREPORTS.Refresh

End If

If Left(HTYPE, 1) = 5 Then

Unload frmGUIDE
frmNOPRICES.Enabled = True
frmNOPRICES.Show
frmNOPRICES.Refresh

End If


End Sub
