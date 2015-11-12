VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmGUIDE 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5025
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   5025
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox rtsGUIDE 
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
      Enabled         =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmGUIDE.frx":0000
   End
End
Attribute VB_Name = "frmGUIDE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

frmGUIDE.Caption = msgtab(243)

frmGUIDE.MousePointer = 1

HFNM = APPROOT + "\ARTBAS\HELP\" + current_language + "HELP" + HTYPE + ".rtf"

rtsGUIDE.FileName = HFNM
rtsGUIDE.Refresh

End Sub
Private Sub rtsGUIDE_Click()

rtsGUIDE.FileName = HFNM
rtsGUIDE.Refresh

End Sub

Private Sub rtsGUIDE_dblClick()

'rtsGUIDE.SaveFile HFNM

If Left(HTYPE, 1) = "1" Then

Unload frmGUIDE
frmARTB00.Enabled = True
frmARTB00.Show
frmARTB00.Refresh

End If

If Left(HTYPE, 1) = "2" Then

Unload frmGUIDE
frmARTB01.Enabled = True
frmARTB01.Show
frmARTB01.Refresh

End If

If Left(HTYPE, 1) = "3" Then

Unload frmGUIDE
frmTABLES.Enabled = True
frmTABLES.Show
frmTABLES.Refresh

End If

If Left(HTYPE, 1) = "4" Or Left(HTYPE, 1) = "5" Or Left(HTYPE, 1) = "6" Then

Unload frmGUIDE
frmEFFORT.Enabled = True
frmEFFORT.Show
frmEFFORT.Refresh

End If

If Left(HTYPE, 1) = "7" Or Left(HTYPE, 1) = "8" Then

Unload frmGUIDE
frmLAND.Enabled = True
frmLAND.Show
frmLAND.Refresh

End If

If Left(HTYPE, 1) = "9" Or Left(HTYPE, 1) = "A" Then

Unload frmGUIDE
frmLAND.Enabled = True
frmLAND.Show
frmLAND.Refresh

End If

If Left(HTYPE, 1) = "B" Then

Unload frmGUIDE
frmACTIVE.Enabled = True
frmACTIVE.Show
frmACTIVE.Refresh

End If

If Left(HTYPE, 1) = "C" Then

Unload frmGUIDE
frmESTIM.Enabled = True
frmESTIM.Show
frmESTIM.Refresh

End If

If Left(HTYPE, 1) = "D" Then

Unload frmGUIDE
frmREP.Enabled = True
frmREP.Show
frmREP.Refresh

End If

End Sub
