VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmDB02 
   Caption         =   "Form1"
   ClientHeight    =   7590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10710
   LinkTopic       =   "Form1"
   ScaleHeight     =   7590
   ScaleWidth      =   10710
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstYM 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   5580
      Left            =   0
      MultiSelect     =   1  'Simple
      TabIndex        =   0
      Top             =   0
      Width           =   3135
   End
   Begin ComctlLib.ProgressBar pgbFILES 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   6120
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   5760
      Width           =   9495
   End
End
Attribute VB_Name = "frmDB02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

Call CRDB_LOAD_YM

End Sub
Private Sub CRDB_LOAD_YM()

Dim I, NY, YYX, YSTR, BNK, XXX, YYY, YT(), YS(), NFIRST, NLAST
Dim YMSEL(), YYMMSEL, YMTAB()

'------- Create table with no-empty years and months -----------

NY = 0: YYMMSEL = 0

Open APPROOT + "\ARTBAS\CONTROL\CONTENTS.TXT" For Input As #1

Do Until EOF(1)

Line Input #1, XXX

YYX = Left(XXX, 4): YSTR = Mid(XXX, 6, 12)

YYY = LTrim(RTrim(YSTR))

If YYY = "" Then
   GoTo CONT_LOOP
   End If

NY = NY + 1

ReDim Preserve YT(1 To NY), YS(1 To NY)

YT(NY) = YYX: YS(NY) = YSTR

CONT_LOOP:

Loop

Close #1

If NY = 0 Then End

'=============================================================

Dim MY, JX, NYM, J, K

lstYM.Clear: NYM = 0

For I = 1 To NY
For J = 1 To 12

JX = Format(J, "00")

If Mid(YS(I), J, 1) = " " Then GoTo NEXT_J

lstYM.AddItem JX + "/" + YT(I)

NYM = NYM + 1

ReDim Preserve YMSEL(1 To NYM)
ReDim Preserve YMTAB(1 To NYM)

lstYM.ItemData(lstYM.NewIndex) = NYM
YMTAB(NYM) = JX + "/yt(i)": YMSEL(NYM) = 0

NEXT_J:

Next J
Next I

'If NYM = 0 Then End

pgbFILES.Min = 0
pgbFILES.Max = NYM
pgbFILES.Value = 0
pgbFILES.Visible = True
Label2.Visible = True

lstYM.ListIndex = 0
lstYM.Refresh

YYMMSEL = 0

END_SUB:

End Sub
