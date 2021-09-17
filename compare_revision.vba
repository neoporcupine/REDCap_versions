Option Explicit

#If VBA7 Then ' Excel 2010 or later
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As LongPtr)
Public Declare PtrSafe Function GetTickCount Lib "kernel32.dll" () As Long
#Else ' Excel 2007 or earlier
Public Declare Sub Sleep Lib "kernel32" (ByVal Milliseconds As Long)
Public Declare Function GetTickCount Lib "kernel32.dll" () As Long
#End If


' BUILD CODEBOOK VERSIONS HTML DOCUMENT
' Examine data dictionaries downloaded from REDCap "Project Revision History"
' REQUIRES:
' [1] each revision in a tab named "r1" ... "rNN" where NN is the number of the revision, these sheets contain the data dictionary for that revision.
' [2] "Revisions" tab contains columns revision name, tab name {ignored}, revision date, other stuff {ignored}
' Build CODEBOOK of each field in nice html format with historical changes
' Key sub = BuildCodebook() so link this to a button if you wish, or just run it
' TO DO: backward walk though revisions to see if/when any fields are removed

Const Base_Rev = 10 ' The most recent revision number
Const Revisions_Start_Row = 3 ' row on Revisions tab that contains r1
Const OutFileFolder = "c:\tmp\"  ' HTML output file is saved here, ensure ends with \
Const OutFileNamePrefix = "CodeBookVersions" ' the start of the file name
' Filename will be in the format: CodeBookVersions_yyyy-mm-dd.html

' ******************************************************
' ******************************************************

Const htmlhead = "<html>" & vbCrLf & "<head>" & vbCrLf & _
    "<TITLE>CodeBook Versions</TITLE><meta charset='utf-8' />" & vbCrLf & _
    "<STYLE>" & vbCrLf & _
    ".sfname {font-size:1.2em;font-weight:bold;background-color:#ccffff;}" & vbCrLf & _
    ".sfnamehide {font-size:1.2em;font-weight:bold;background-color:#ffcccc;}" & vbCrLf & _
    ".sfnamecalc {font-size:1.2em;font-weight:bold;background-color:#ffddcc;}" & vbCrLf & _
    ".sfnamedesc {font-size:1.2em;font-weight:bold;background-color:#eee;}" & vbCrLf & _
    ".sfheader {font-size:1.2em;font-weight:bold;background-color:#f5ef42;}" & vbCrLf & _
    ".sform {font-weight:bold;font-style: italic;background-color:#ccffdd;}" & vbCrLf & _
    ".sflabel, .schoices, .ifr {}" & vbCrLf & _
    ".sftype, .sfvalid, .ift, .ify {font-size:0.8em}" & vbCrLf & _
    ".sblogic {color:#A0B;}" & vbCrLf & _
    ".sfann {color:#06C;}" & vbCrLf & _
    "</STYLE>" & vbCrLf & _
    "</HEAD>" & vbCrLf & _
    "<BODY>" & vbCrLf
Const htmlkey = "<H1>Key</H1>" & vbCrLf & _
    "<table>" & vbCrLf & _
    "<TR><TD class='sfname'>field_name</TD><TD class='sform'>form_name</TD><TD class='sftype'>ValueType</TD><TD class='sfvalid'>Validation</TD><TD class='sftype'>Min</TD><TD class='sftype'>Max</TD><TD class='sftype'>Required</TD><TD class='sftype'>Idenifying</TD></TR>" & vbCrLf & _
    "<TR><TD class='sfnamedesc'>descriptive_field</TD><TD COLSPAN=7>Accepts no values, just displays fixed text on form.</TD></TR>" & vbCrLf & _
    "<TR><TD class='sfnamecalc'>calculated_field</TD><TD COLSPAN=7>Generates values based on calculation, no user entry.</TD></TR>" & vbCrLf & _
    "<TR><TD class='sfnamehide'>hidden_field</TD><TD COLSPAN=7>Field is hidden from user by Field Annotations or Branching Logic.</TD></TR>" & vbCrLf & _
    "<TR><TD class='sfheader'>-- SECTION HEADING --</TD><TD COLSPAN=7>Section headings are special because they are not seperate rows, these are display only, and no historical information is listed.</TD></TR>" & vbCrLf & _
    "<TR><TD>FormText:</TD><TD COLSPAN=7>Text that appears on the web forms.</TD></TR>" & vbCrLf & _
    "<TR><TD>FieldNote:</TD><TD COLSPAN=7>Text that appears underneath the data entry field.</TD></TR>" & vbCrLf & _
    "<TR><TD>Choices:</TD><TD COLSPAN=7>Categorical fields contains the category definitions. Calculated fields contains the formula.</TD></TR>" & vbCrLf & _
    "<TR><TD>BrLogic:</TD><TD COLSPAN=7>Branching Logic - only display the field if this logic evaluates to TRUE</TD></TR>" & vbCrLf & _
    "<TR><TD>FieldAnnot:</TD><TD COLSPAN=7>Special REDCap field annotation codes. Notably @HIDDEN hides the field.</TD></TR>" & vbCrLf & _
    "<TR><TD>History:</TD><TD COLSPAN=7>When was this field added to the project. How has this field been altered through the revisions.</TD></TR>" & vbCrLf & _
    "</table>" & vbCrLf

Dim OutFileName

Sub myDeleteFile(thisfile)
Dim fs
Set fs = CreateObject("Scripting.FileSystemObject")
If fs.FileExists(thisfile) Then
    fs.DeleteFile thisfile
End If
Set fs = Nothing
End Sub

Sub ryt(thisString) ' Write the string ... append to file
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Dim f, fso
Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.OpenTextFile(OutFileName, ForAppending, True, -1)
f.Write thisString
f.Close
Set f = Nothing
Set fso = Nothing
End Sub

Function i_Name(revNum)
i_Name = Worksheets("Revisions").Cells(revNum + Revisions_Start_Row - 1, 1).Value
End Function

Function i_Sheet(revNum) ' NOT USED
i_Sheet = Worksheets("Revisions").Cells(revNum + Revisions_Start_Row - 1, 2).Value
End Function

Function i_Date(revNum)
i_Date = Worksheets("Revisions").Cells(revNum + Revisions_Start_Row - 1, 3).Value
End Function

Function r_rc(rev, r, c)
r_rc = Worksheets("r" & rev).Cells(r, c).Value
End Function

' *** Note for "r_*" functions
' we add +1 to the row, so row 1 is actually the second row in order to skip the header row
' rev is the number NN 1..Base_Rev of the spreadsheet revNN
' r is the row, if r_name is blank then you've reached past the end

Function r_name(rev, r)
r_name = r_rc(rev, r + 1, 1)
End Function

Function r_form(rev, r)
r_form = r_rc(rev, r + 1, 2)
End Function

Function r_sect(rev, r)
r_sect = r_rc(rev, r + 1, 3)
End Function

Function r_ftype(rev, r)
r_ftype = r_rc(rev, r + 1, 4)
End Function

Function r_flabel(rev, r)
r_flabel = r_rc(rev, r + 1, 5)
End Function

Function r_fchoice(rev, r)
r_fchoice = r_rc(rev, r + 1, 6)
End Function

Function r_fnote(rev, r)
r_fnote = r_rc(rev, r + 1, 7)
End Function

Function r_fvalid(rev, r)
r_fvalid = r_rc(rev, r + 1, 8)
End Function

Function r_fvmin(rev, r)
r_fvmin = r_rc(rev, r + 1, 9)
End Function

Function r_fvmax(rev, r)
r_fvmax = r_rc(rev, r + 1, 10)
End Function

Function r_fident(rev, r)
r_fident = r_rc(rev, r + 1, 11)
End Function

Function r_fblogic(rev, r)
r_fblogic = r_rc(rev, r + 1, 12)
End Function

Function r_freq(rev, r)
r_freq = r_rc(rev, r + 1, 13)
End Function

Function r_fann(rev, r)
r_fann = r_rc(rev, r + 1, 18)
End Function


Function IfTitle(ttl, thisText)  ' only return title + thistext if thistext is not ""
If (thisText & "") <> "" Then
    IfTitle = "<span class='ift'>" & ttl & " " & thisText & "</span>"
Else
    IfTitle = ""
End If
End Function


Function Ify(ttl, thisText)  ' only return ttl if thisText is "y"
If thisText = "y" Then
    Ify = "<span class='ify'>" & ttl & "</span> "
Else
    Ify = ""
End If
End Function


Function IfRow(stl, ttl, thisText)  ' only return title + thistext if thistext is not ""
If (thisText & "") <> "" Then
    If stl = "" Then stl = "ifr"
    IfRow = "<TR><TD COLSPAN=8 class='" & stl & "'>" & ttl & thisText & "</TD></TR>" & vbCrLf
Else
    IfRow = ""
End If
End Function


Function fhidden(fldname, brlogic, annot) ' fieldname is being hidden by logic or annotation
If brlogic <> "" And InStr(brlogic, "[" & fldname & "]") > 0 Then
    fhidden = True
ElseIf annot <> "" And InStr(annot, "@HIDDEN") > 0 Then
    fhidden = True
Else
    fhidden = False
End If
End Function


Function FieldInRev(p_rev, p_fldname) ' returns row number, 0 if not found
Dim r, found, fname
r = 0
found = False
Do
    r = r + 1
    fname = r_name(p_rev, r)
    If p_fldname = fname Then found = True
Loop Until found Or fname = ""
If found Then FieldInRev = r Else FieldInRev = 0
End Function


Function FirstAppear(p_fldname, ByRef p_rownum) ' Returns revision number where fldname first appears, 0 if not found (?!)
Dim rev, found, rownum
rev = 0
Do
    rev = rev + 1
    rownum = FieldInRev(rev, p_fldname)
    ' MsgBox "FirstAppear (" & rev & "," & p_fldname & ")"
    If rownum > 0 Then found = True
Loop Until found Or rev = Base_Rev
If found Then
    p_rownum = rownum
    FirstAppear = rev
Else
    p_rownum = 0
    FirstAppear = 0
End If
End Function


Sub Testfa()
' 'subjid'
Dim rev, row
rev = FirstAppear("subjid", row)
'MsgBox "RESULT:: rev=" & rev & ", row=" & row
'row = FieldInRev(1, "subjid")
'MsgBox "FieldInRev(1, ""subjid"")=" & row
'MsgBox "r_name(2, 1) = " & r_name(2, 1)
End Sub


Function CompareCell(ttl, c1, c2)
Dim q1, q2
q1 = Replace(UCase(Trim(c1)), " ", "")
q2 = Replace(UCase(Trim(c2)), " ", "")
If q1 <> q2 Then
    CompareCell = "<BR><b>" & ttl & "</b> new value = '" & c2 & "', was value = '" & c1 & "'" & vbCrLf
Else
    CompareCell = ""
End If
End Function


Function RevHistory(p_fldname, p_revnum)
' compare p_fldname for p_revnum vs p_revnum+1
Dim a_row, b_row, html
a_row = FieldInRev(p_revnum, p_fldname)
b_row = FieldInRev(p_revnum + 1, p_fldname)
If (a_row = 0) Then
    html = "<p>Did not exist in " & i_Name(p_revnum) & "</p>"
ElseIf (b_row = 0) Then ' Don't report as this will come up during looping
    html = ""
Else
    html = ""
    ' compare cells
    html = html & CompareCell("Form Name", r_form(p_revnum, a_row), r_form(p_revnum + 1, b_row))
    html = html & CompareCell("Type", r_ftype(p_revnum, a_row), r_ftype(p_revnum + 1, b_row))
    html = html & CompareCell("Label", r_flabel(p_revnum, a_row), r_flabel(p_revnum + 1, b_row))
    html = html & CompareCell("Choices", r_fchoice(p_revnum, a_row), r_fchoice(p_revnum + 1, b_row))
    html = html & CompareCell("Note", r_fnote(p_revnum, a_row), r_fnote(p_revnum + 1, b_row))
    html = html & CompareCell("Type", r_fvalid(p_revnum, a_row), r_fvalid(p_revnum + 1, b_row))
    html = html & CompareCell("Min", r_fvmin(p_revnum, a_row), r_fvmin(p_revnum + 1, b_row))
    html = html & CompareCell("Max", r_fvmax(p_revnum, a_row), r_fvmax(p_revnum + 1, b_row))
    html = html & CompareCell("BrLogic", r_fblogic(p_revnum, a_row), r_fblogic(p_revnum + 1, b_row))
    html = html & CompareCell("FieldAnnot", r_fann(p_revnum, a_row), r_fann(p_revnum + 1, b_row))
    If html <> "" Then
        html = "<p>" & i_Name(p_revnum + 1) & " " & Format(i_Date(p_revnum + 1), "yyyy-mm-dd") & vbCrLf & html
    End If
End If
RevHistory = html
End Function


' STYLES  sfname sflabel schoices ift ify ifr sftype sfvalid sblogic sfann sfnamecalc sfnamedesc

Sub BuildCodebook()
If OutFileName & "" = "" Then OutFileName = OutFileFolder & OutFileNamePrefix & "_" & Format(Date, "yyyy-mm-dd") & ".html"
myDeleteFile OutFileName
Dim fNum, fname, html, tmpc, sect, fieldstyle
Dim rev1, row1 ' revision and row where this field first appears
Dim revcomp
Dim bench1, bench2
bench1 = Now
ryt htmlhead & htmlkey & "<H1>Codebook (versions)</H1>" & vbCrLf & "<TABLE>" & vbCrLf
fNum = 0

Do
    html = ""
    fNum = fNum + 1
    fname = r_name(Base_Rev, fNum)
    If fname <> "" Then
        sect = r_sect(Base_Rev, fNum)
        If sect <> "" Then ' There is a section header
            html = html & "<tr><td COLSPAN=8><hr></td></tr>" & vbCrLf & _
            "<TR><TD class='sfheader'>-- SECTION HEADING --</TD><TD class='sform'>" & _
            r_form(Base_Rev, fNum) & "</TD></TR>" & vbCrLf & _
            "<TR><TD COLSPAN=8 class='sflabel'>" & sect & "</TD></TR>"
        End If
        If fhidden(fname, r_fblogic(Base_Rev, fNum), r_fann(Base_Rev, fNum)) Then
            fieldstyle = "sfnamehide"
        ElseIf (r_ftype(Base_Rev, fNum) = "descriptive") Then
            fieldstyle = "sfnamedesc"
        ElseIf (r_ftype(Base_Rev, fNum) = "calc") Then
            fieldstyle = "sfnamecalc"
        Else
            fieldstyle = "sfname"
        End If
        html = html & "<tr><td COLSPAN=8><hr></td></tr>" & vbCrLf & _
        "<TR><TD class='" & fieldstyle & "'>" & fname & "</TD><TD class='sform'>" & _
        r_form(Base_Rev, fNum) & "</TD><TD class='sftype'>" & vbCrLf & _
        r_ftype(Base_Rev, fNum) & "</TD><TD class='sfvalid'>" & vbCrLf & _
        r_fvalid(Base_Rev, fNum) & "</TD><TD>" & vbCrLf & _
        IfTitle("Min:", r_fvmin(Base_Rev, fNum)) & "</TD><TD>" & vbCrLf & _
        IfTitle("Max:", r_fvmax(Base_Rev, fNum)) & "</TD><TD>" & vbCrLf & _
        Ify("Required", r_freq(Base_Rev, fNum)) & "</TD><TD>" & vbCrLf & _
        Ify("Identifying", r_fident(Base_Rev, fNum)) & "</TD>" & vbCrLf & _
        "</TR>" & vbCrLf
        html = html & "<TR><TD COLSPAN=8 class='sflabel'><b>FormText:</b> " & r_flabel(Base_Rev, fNum) & "</TD></TR>" & vbCrLf
        html = html + IfRow("sfnote", "<b>FieldNote:</b> ", r_fnote(Base_Rev, fNum))
        html = html + IfRow("schoices", "<b>Choices:</b> ", r_fchoice(Base_Rev, fNum))
        html = html + IfRow("sblogic", "<b>BrLogic:</b> ", r_fblogic(Base_Rev, fNum))
        html = html + IfRow("sfann", "<b>FieldAnnot:</b> ", r_fann(Base_Rev, fNum))
        rev1 = FirstAppear(fname, row1)
        html = html & "<tr><td COLSPAN=8><b>HISTORY:</b> First appears " & i_Name(rev1) & " " & i_Date(rev1) & vbCrLf
        If rev1 > 0 And rev1 <> Base_Rev Then
            For revcomp = rev1 To Base_Rev - 1
                html = html & RevHistory(fname, revcomp)
            Next
        html = html & "</td></tr>" & vbCrLf
        End If
    End If
    ryt html
Loop Until fname = ""

ryt "<tr><td COLSPAN=8><hr></td></tr><tr><td COLSPAN=8><hr></td></tr><tr><td COLSPAN=8><hr></td></tr>" & vbCrLf & "</TABLE>" & vbCrLf

bench2 = Now
ryt "<p>Processed in: " & Format((bench2 - bench1) * 86400, "0.0") & " seconds</p>" & vbCrLf & "</BODY>" & vbCrLf & "</HTML>" & vbCrLf

MsgBox "Completed in " & Format((bench2 - bench1) * 86400, "0.0") & " seconds." & vbCrLf & "File written to:" & vbCrLf & OutFileName, vbOKOnly, "Completed!"
End Sub



Sub test()
Dim n1, n2, t1, t2
n1 = Now
t1 = GetTickCount
Sleep 1234
n2 = Now
t2 = GetTickCount
MsgBox "diff Now = " & Format((n2 - n1) * 86400, "0.0") & vbCrLf & "diff ticker = " & Format((t2 - t1) / 1000, "0.0")
End Sub


