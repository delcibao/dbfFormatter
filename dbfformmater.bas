Attribute VB_Name = "DBFFormatter"
Sub DBFFormatter()
' by M.M. 2015
' Parse DBF file and Fix column header using
' the corresponding Tagname DBF file
'
Dim currentWorkbk As Workbook   'Wide DBF file
Dim tagWorkbk As Workbook       'Tagname DBF file
Dim cSheet As Worksheet
Dim tSheet As Worksheet
Dim total As Integer
Dim n As Integer
Dim m As Integer
Dim str1() As String

Dim varTagfile As String        'complete Tagname DBF file
Dim varFilename As String       'Wide DBF file
Dim varDirectory As String      'Current directory
Dim pref(2) As Boolean

varFilename = Application.ActiveWorkbook.Name
varDirectory = Application.ActiveWorkbook.Path

If (InStr(varFilename, "Wide") > 0) Then
    varTagfile = Replace(varFilename, "Wide", "Tagname")
Else
    MsgBox "Wrong file name", vbExclamation, "Error!"
    End
End If

 If Len(Dir(varDirectory & "\" & varTagfile)) = 0 Then
   MsgBox "File " & varTagfile & " does not exist!", vbExclamation, "Error!"
   End
 End If

Set currentWorkbk = ActiveWorkbook
Set tagWorkbk = Workbooks.Open(varDirectory & "\" & varTagfile)

Set cSheet = currentWorkbk.Sheets(1)
Set tSheet = tagWorkbk.Sheets(1)

pref(0) = Application.ScreenUpdating
pref(1) = Application.DisplayAlerts
Application.ScreenUpdating = False
Application.DisplayAlerts = False

n = 1
m = 4

Do Until IsEmpty((tSheet.Range("A1").Offset(n, 0).Value))
    str1 = Split(tSheet.Range("A1").Offset(n, 0).Value, "\")
    cSheet.Range("A1").Offset(0, m).Value = str1(1)
    n = n + 1
    m = m + 2
Loop
m = m + 1
Do Until m = 1
    'cSheet.Range("A1").Offset(0, m).Select
    'Columns("L:L").Select
    'Selection.Delete Shift:=xlToLeft
    'Columns("J:J").Select
    'Selection.Delete Shift:=xlToLeft
    cSheet.Range("A1").Offset(0, m).EntireColumn.Delete
    m = m - 2
Loop

Application.ScreenUpdating = pref(0)
Application.DisplayAlerts = pref(1)

tagWorkbk.Close False

End Sub
