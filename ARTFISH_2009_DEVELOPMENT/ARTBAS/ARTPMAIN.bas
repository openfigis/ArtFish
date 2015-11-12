Attribute VB_Name = "Module1"
Option Explicit

' TNJ Note - all the code subroutines in this module are ordered by alphabetical name

Public OPT_LOCAL, OPT_UN

Public APPROOT, RUN_FILE, RUN_CODE, EXPFNM

Public CONVEX_YN, INACC, OUTACC, OUTACC1, OUTACC2, INSMP, OUTSMP1, OUTSMP2, POPSIZE, OUTSMP
Public SPST_IND

Public monthtab(), msgtab(), CY() As String
Public language As String
Public messages As Database

Public msgrec As Recordset
Public msgtabname As String
Public msgtab_e, msgtab_f, msgtab_s, msgtab_l As String
Public Nmsg, imsg As Integer

Public current_language, current_language_title

Public current_month, CURCAL, CTLMONTH As Integer
Public current_monthx As String * 12
Public current_nocal As Integer
Public current_year As Integer

Public CURMNC, CURMNN, CURMJC

Public UNW, UNM, CTLADMIN, BKPF, HTYPE, HFNM

'Module
Public Declare Function ReleaseMutex Lib "kernel32" (ByVal hMutex As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function CreateMutex Lib "kernel32" Alias "CreateMutexA" (lpMutexAttributes As SECURITY_ATTRIBUTES, ByVal bInitialOwner As Long, ByVal lpName As String) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function FlashWindowEx Lib "user32.dll" (ByRef pfwi As FLASHWINFO) As Long
 
 
Private Const FLASHW_STOP = 0
Private Const FLASHW_CAPTION = &H1
Private Const FLASHW_TRAY = &H2
Private Const FLASHW_ALL = (FLASHW_CAPTION Or FLASHW_TRAY)
Private Const FLASHW_TIMER = &H4
 
Private Const ERROR_ALREADY_EXISTS = 183&
Private Const SW_RESTORE = 9
Private Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type
 
' Structure for flashing the window to indicate that an instance is already existing
Private Type FLASHWINFO
    cbSize As Long
    hwnd As Long
    dwFlags As Long
    uCount As Long
    dwTimeout As Long
End Type
 
Public Mutex As Long

' Function for getting the 'short format' (DOS) for path names
Public Declare Function GetShortPathName Lib "kernel32" _
    Alias "GetShortPathNameA" _
    (ByVal lpszLongPath As String, _
    ByVal lpszShortPath As String, _
    ByVal cchBuffer As Long) As Long

Public Sub ACCURACY_FOR_GIVEN_SAMPLES()

Dim V, Z, AC, NPOP, POP_FLAG, A, G, S, K, A1, A2, W, NS

NS = INSMP

Z = 1.96

If CONVEX_YN = "Y" Then POP_FLAG = "CONVEX"
If CONVEX_YN = "N" Then POP_FLAG = "CONCAVE"

NPOP = POPSIZE

If NPOP < 10 Then
   OUTACC = 0
   Exit Sub
   End If

GoSub CALC_NS
GoSub CALC_NS2

GoTo END_SUB

CALC_NS:

If POP_FLAG = "CONVEX" Then
   W = 0.75 * (1 - 1 / NPOP)
   End If

If POP_FLAG = "CONCAVE" Then
   W = 1 - Log(1 + 0.5 * Exp(1 / NPOP))
   End If

A = 2 * W * NPOP ^ 2 / (NPOP - 1) ^ 2 - (NPOP + 1) / (NPOP - 1)
G = A + (1 - A) / NPOP
S = (1 - A) * (1 / Log(NPOP) - 1 / (NPOP * Log(NPOP)) - 1 / NPOP)
K = (-2 / Log(NPOP)) * Log(S / (1 - S - G))
A2 = (1 - S - G) ^ 2 / (2 * S + G - 1)
A1 = G - A2

OUTACC1 = A1 + A2 * NPOP ^ (-K * Log(NS) / Log(NPOP))
'NS = ((AC - A1) / (A2)) ^ (-1 / K)
'OUTSMP1 = NS

Return

CALC_NS2:

If POP_FLAG = "CONVEX" Then
   V = (2 * NPOP - 1) / (6 * (NPOP - 1)) - 0.25
   V = V ^ 0.5
   End If

If POP_FLAG = "CONCAVE" Then
   V = 0.25
   V = V ^ 0.5
   End If

OUTACC2 = 1 - Z * V * NS ^ (-0.5) * (1 - NS / NPOP) ^ -0.5
'NS = (((1 - AC) / (Z * V)) ^ 2 + 1 / NPOP) ^ (-1)
'OUTSMP2 = NS

Return

END_SUB:

OUTACC = OUTACC1: SPST_IND = "(SPST)"

If OUTACC2 > OUTACC1 Then
   OUTACC = OUTACC2
   SPST_IND = " "
   End If

End Sub

Public Sub CALC_CALDAYS()

Dim I, K

I = current_year
K = I - 4 * Int(I / 4)

I = current_month

CURCAL = 31

If I = 4 Or I = 6 Or I = 9 Or I = 11 Then CURCAL = 30

If I = 2 And K = 0 Then CURCAL = 29
If I = 2 And K <> 0 Then CURCAL = 28

End Sub

Public Sub CHECK_BACKUP_COMPLETE()

BKPF = APPROOT + "\ARTBAS\CONTROL\Y" + Format(current_year, "0000") + "M" + _
       Format(current_month, "00") + "BKP.TXT"

If Dir(BKPF) <> "" Then Kill BKPF

End Sub

Public Sub EXCEL_REQUEST_ERROR()

Dim resp, XXX

XXX = RTrim(msgtab(268)) + " " + RTrim(msgtab(269)) + RTrim(msgtab(272)) + " " + RTrim(msgtab(273))

resp = MsgBox(XXX, vbOKOnly, "Excel???  ")

End Sub

Public Function GetShortName(sFile As String) As String
    Dim sShortFile As String * 67
    Dim lResult As Long

    'Make a call to the GetShortPathName API
    lResult = GetShortPathName(sFile, sShortFile, _
    Len(sShortFile))

    'Trim out unused characters from the string.
    GetShortName = Left$(sShortFile, lResult)

End Function
 
Public Sub KILL_ARTBASIC_FOLDER()

' Commented out - TNJ - 14/1/08
' Also added the code below to close all open forms
' If Dir("C:\ARTBASIC_CURRENT_FOLDER.TXT") <> "" Then
'    Kill "C:\ARTBASIC_CURRENT_FOLDER.TXT"
'    End If
   
Dim frm As Form

  ' Loop through all open forms and close them
  For Each frm In Forms
              Unload frm
  Next frm
   
End Sub

Private Sub Main()
    Dim SA As SECURITY_ATTRIBUTES
    Dim hwnd As Long
    Dim FlashInfo As FLASHWINFO
 
    SA.bInheritHandle = 1
    SA.lpSecurityDescriptor = 0
    SA.nLength = Len(SA)
    
    ' This code ensures that only 1 instance of ArtBasic is running at a time
    Mutex = CreateMutex(SA, 1, App.Title)
    If (Err.LastDllError = ERROR_ALREADY_EXISTS) Then
        hwnd = GetSetting("ArtBasic", "ARTB00", "StartUp", 0)
        ShowWindow hwnd, SW_RESTORE
        SetForegroundWindow hwnd
        
        FlashInfo.cbSize = Len(FlashInfo)
        FlashInfo.dwFlags = FLASHW_ALL Or FLASHW_TIMER
        FlashInfo.dwTimeout = 0
        FlashInfo.hwnd = hwnd
        FlashInfo.uCount = 3
        FlashWindowEx FlashInfo
    Else
        frmARTB00.Show
    End If
End Sub

' Public Sub APPROOT_WRITE()

' Commented out by Tony - 14/1/08 - This was done because users
' do not have access to writing on the root of the C drive.
' This routine is no longer necessary but left for reference purposes

' Open "C:\ARTBASIC_CURRENT_FOLDER.TXT" For Output As #4
' Print #4, APPROOT
' Close #4

' End Sub

Public Sub MSGLOAD()

Dim DIRINP, DIROUT

ReDim msgtab(1 To 500)

Open APPROOT + "\ARTBAS\MESSAGES\ARTBMSG.TXT" For Input As #2

Dim XNO, INO, XENG, XFR, XSP, XLOC

Do Until EOF(2)

    Input #2, INO, XENG, XFR, XSP, XLOC

    If INO = 1 Then
        msgtab_e = XENG
        msgtab_f = XFR
        msgtab_s = XSP
        msgtab_l = XLOC
    End If

    If language = "E" Then msgtab(INO) = XENG       ' English
    If language = "F" Then msgtab(INO) = XFR        ' French
    If language = "S" Then msgtab(INO) = XSP        ' Spanish
    If language = "L" Then msgtab(INO) = XLOC       ' Local

Loop

Close #2


' Commented out by TNJ - 3/6/09 because all help files were moved to the
' same directory with 1-byte prefixes for languages i.e. E=English, F=French, etc.
'
' If language = "ENGLISH" Then DIRINP = APPROOT + "\ARTBAS\HELP_ENGLISH\"
' If language = "FRENCH" Then DIRINP = APPROOT + "\ARTBAS\HELP_FRENCH\"
' If language = "SPANISH" Then DIRINP = APPROOT + "\ARTBAS\HELP_SPANISH\"
' If language = "LOCAL" Then DIRINP = APPROOT + "\ARTBAS\HELP_LOCAL\"

' Dim WRKDIR, FNM1, fnm2

' DIROUT = APPROOT + "\ARTBAS\HELP\"

' FNM1 = Dir(DIRINP + "*.rtf")
' fnm2 = DIROUT + FNM1

' FileCopy DIRINP + FNM1, fnm2

' Do While FNM1 <> ""

'     FNM1 = Dir
'     fnm2 = DIROUT + FNM1

'     If FNM1 = "" Then GoTo EXIT_COPY

'     FileCopy DIRINP + FNM1, fnm2

' Loop

' EXIT_COPY:

End Sub
Public Sub SAMPLES_FOR_GIVEN_ACCURACY()

Dim V, Z, AC, NPOP, POP_FLAG, A, G, S, K, A1, A2, W, NS

AC = INACC

Z = 1.96

If CONVEX_YN = "Y" Then POP_FLAG = "CONVEX"
If CONVEX_YN = "N" Then POP_FLAG = "CONCAVE"

NPOP = POPSIZE

If NPOP < 10 Then
   OUTSMP = 0
   Exit Sub
   End If

GoSub CALC_NS
GoSub CALC_NS2

GoTo END_SUB

CALC_NS:

If POP_FLAG = "CONVEX" Then
   W = 0.75 * (1 - 1 / NPOP)
   End If

If POP_FLAG = "CONCAVE" Then
   W = 1 - Log(1 + 0.5 * Exp(1 / NPOP))
   End If

A = 2 * W * NPOP ^ 2 / (NPOP - 1) ^ 2 - (NPOP + 1) / (NPOP - 1)
G = A + (1 - A) / NPOP
S = (1 - A) * (1 / Log(NPOP) - 1 / (NPOP * Log(NPOP)) - 1 / NPOP)
K = (-2 / Log(NPOP)) * Log(S / (1 - S - G))
A2 = (1 - S - G) ^ 2 / (2 * S + G - 1)
A1 = G - A2
NS = ((AC - A1) / (A2)) ^ (-1 / K)
OUTSMP1 = NS

Return

CALC_NS2:

If POP_FLAG = "CONVEX" Then
   V = (2 * NPOP - 1) / (6 * (NPOP - 1)) - 0.25
   V = V ^ 0.5
   End If

If POP_FLAG = "CONCAVE" Then
   V = 0.25
   V = V ^ 0.5
   End If

NS = (((1 - AC) / (Z * V)) ^ 2 + 1 / NPOP) ^ (-1)
OUTSMP2 = NS

Return

END_SUB:

OUTSMP = OUTSMP1: SPST_IND = "(SPST)"

If OUTSMP2 < OUTSMP1 Then
   OUTSMP = OUTSMP2
   SPST_IND = " "
   End If

End Sub
Public Sub write_parms()

Dim fnm, ccy, ccm

fnm = APPROOT + "\ARTBAS\CONTROL\ARTBRUN.TXT"

If Dir(fnm) <> "" Then Kill fnm

ccy = current_year: ccm = CTLMONTH

fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(ccy, "0000") + "M" + Format(ccm, "00") + "_MAJOR.TXT"

If Dir(fnm) <> "" Then CY(ccy - 1989, ccm) = "X"

Dim XXX

XXX = APPROOT + "\ARTBAS\CONTROL\WCALL.TXT"

If Dir(XXX) <> "" Then Kill XXX

Open APPROOT + "\ARTBAS\CONTROL\SYSPARM.TXT" For Output As #1

Print #1, current_language
Print #1, Format(current_year, "0000")

Close #1

Open APPROOT + "\ARTBAS\CONTROL\CONTENTS.TXT" For Output As #1

Dim I, J

For I = 1990 To 2020

    XXX = ""

    For J = 1 To 12

        fnm = APPROOT + "\ARTBAS\TABLES\Y" + Format(I, "0000") + "M" + Format(J, "00") + "_MAJOR.TXT"

        CY(I - 1989, J) = " "

        If Dir(fnm) <> "" Then CY(I - 1989, J) = "X"

        XXX = XXX + CY(I - 1989, J)

    Next J

    Print #1, Format(I, "0000") + " " + XXX

Next I

Close #1

End Sub

