Attribute VB_Name = "Module1"
Option Explicit

Public AW_OPTION

Public TCATCH, FCATCH, PCATCH, TVALUE, TFISH
Public APPROOT, RUN_FILE, RUN_CODE, EXPFNM
Public SYSTEM_VALUES, SYSTEM_FISH, SEL_ONLY, COMPVAL, COMPFISH, COMPGEN
Public SHOW_VALUES, SHOW_FISH
Public CATCH_VALUES, PCATCH_VALUES, CATCH_FISH, PCATCH_FISH

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

Public UNW, UNV, CTLADMIN, BKPF, HTYPE, HFNM

Public CURY, CURM

Public C01, C02, C03, C04, C05, C06, C07, C08, C09, C10, C11, C12, C13
Public E01, E02, E03, E04, E05, E06, E07, E08, E09, E10, E11, E12, E13
Public U01, U02, U03, U04, U05, U06, U07, U08, U09, U10, U11, U12, U13
Public P01, P02, P03, P04, P05, P06, P07, P08, P09, P10, P11, P12, P13
Public V01, V02, V03, V04, V05, V06, V07, V08, V09, V10, V11, V12, V13
Public RANK, CUM, PER, RANKYN, RANK_CRIT
Public NCURS, NCURT
Public VALID_MONTHS(), VALIDC, VALIDE, VALIDU, VALIDP, VALIDV, VALIDW, VALIDF
Public REPMJN, REPMNN, REPBGN, REPSPN, REPTC(), REPTE(), REPTU(), REPTP(), REPTV(), REPTW()

' Function for getting the 'short format' (DOS) for path names
Public Declare Function GetShortPathName Lib "kernel32" _
    Alias "GetShortPathNameA" _
    (ByVal lpszLongPath As String, _
    ByVal lpszShortPath As String, _
    ByVal cchBuffer As Long) As Long

Public Function GetShortName(sFile As String) As String
    Dim sShortFile As String * 67
    Dim lResult As Long

    'Make a call to the GetShortPathName API
    lResult = GetShortPathName(sFile, sShortFile, _
    Len(sShortFile))

    'Trim out unused characters from the string.
    GetShortName = Left$(sShortFile, lResult)

End Function
 
Public Sub MSGLOAD()

ReDim msgtab(1 To 500)

Open APPROOT + "\ARTS\MESSAGES\ARTSMSG.TXT" For Input As #2

Dim XNO, INO, XENG, XFR, XSP, XLOC

Do Until EOF(2)

Input #2, INO, XENG, XFR, XSP, XLOC

If INO = 1 Then
   msgtab_e = XENG
   msgtab_f = XFR
   msgtab_s = XSP
   msgtab_l = XLOC
End If

If language = "E" Then msgtab(INO) = XENG
If language = "F" Then msgtab(INO) = XFR
If language = "S" Then msgtab(INO) = XSP
If language = "L" Then msgtab(INO) = XLOC

Loop

Close #2

' Commented out by TNJ because all the help files are going to be in 1 subdirectory
' Dim WRKDIR, fnm1, fnm2, DIRINP, DIROUT

' If language = "ENGLISH" Then DIRINP = APPROOT + "\ARTS\HELP_ENGLISH\"
' If language = "FRENCH" Then DIRINP = APPROOT + "\ARTS\HELP_FRENCH\"
' If language = "SPANISH" Then DIRINP = APPROOT + "\ARTS\HELP_SPANISH\"
' If language = "LOCAL" Then DIRINP = APPROOT + "\ARTS\HELP_LOCAL\"

' DIROUT = APPROOT + "\ARTS\HELP\"

' fnm1 = Dir(DIRINP + "*.rtf")
' fnm2 = DIROUT + fnm1

' FileCopy DIRINP + fnm1, fnm2

' Do While fnm1 <> ""

' fnm1 = Dir
' fnm2 = DIROUT + fnm1

' If fnm1 = "" Then GoTo EXIT_COPY

' FileCopy DIRINP + fnm1, fnm2

' Loop

' EXIT_COPY:

End Sub
Public Sub write_parms()

Dim fnm, ccy, ccm

fnm = APPROOT + "\ARTS\CONTROL\ARTSRUN.TXT"

If Dir(fnm) <> "" Then Kill fnm

ccy = current_year: ccm = CTLMONTH

fnm = APPROOT + "\ARTS\TABLES\Y" + Format(ccy, "0000") + "M" + Format(ccm, "00") + "_MAJOR.TXT"

If Dir(fnm) <> "" Then CY(ccy - 1989, ccm) = "X"

Dim XXX

XXX = APPROOT + "\ARTS\CONTROL\WCALL.TXT"

If Dir(XXX) <> "" Then Kill XXX

Open APPROOT + "\ARTS\CONTROL\SYSPARM.TXT" For Output As #1

Print #1, current_language
Print #1, Format(current_year, "0000")

Close #1

End Sub

' Public Sub APPROOT_WRITE()

' Dim XXX, LLL

' XXX = CurDir()

' LLL = InStr(XXX, "\ARTS")

' APPROOT = Left(XXX, LLL - 1)

' Open APPROOT + "\ARTSER_CURRENT_FOLDER.TXT" For Output As #4
' Print #4, APPROOT
' Close #4

' End Sub

Public Sub KILL_ARTSER_FOLDER()

' Commented out - TNJ - 22/6/09
' Also added the code below to close all open forms
' If Dir("C:\ARTSER_CURRENT_FOLDER.TXT") <> "" Then
'    Kill "C:\ARTSER_CURRENT_FOLDER.TXT"
'    End If
   
Dim frm As Form

  ' Loop through all open forms and close them
  For Each frm In Forms
              Unload frm
  Next frm
   
End Sub
Public Sub EXCEL_REQUEST_ERROR()

Dim resp, XXX

XXX = RTrim(msgtab(102)) + " " + RTrim(msgtab(103)) + " " + RTrim(msgtab(128)) + " " + RTrim(msgtab(129))

resp = MsgBox(XXX, vbOKOnly, "Excel???  ")

End Sub
