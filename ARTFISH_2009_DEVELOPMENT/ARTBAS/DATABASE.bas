Attribute VB_Name = "Module1"
Option Explicit

Public OPT_LOCAL, OPT_UN, LANG_IND, DBN, DBN2, YSEL1, YSEL2, MSEL1, MSEL2

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
