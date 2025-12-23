Attribute VB_Name = "MMainProgram"
' MMainProgram Module - Multi-Class Functions
Option Explicit

' ====== CLASS MANAGEMENT ======
Public Sub CreateNewClassSheet()
    Dim ws As Worksheet
    Dim classCount As Integer
    Dim className As String
    
    ' Count existing class sheets
    classCount = CountClassSheets()
    
    ' Default name: Class1, Class2, etc.
    className = "Class " & (classCount + 1)
    
    ' Ask for custom name
    className = InputBox("Enter class name:", "New Class Sheet", className)
    If className = "" Then Exit Sub
    
    ' Create new sheet
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
    ws.name = className
    
    ' Initialize class sheet
    InitializeClassSheet ws
    
    ' Go to new class sheet
    ws.Activate
    
    MsgBox "New class '" & className & "' created successfully!", vbInformation
End Sub

Private Function CountClassSheets() As Integer
    Dim ws As Worksheet
    Dim count As Integer
    
    count = 0
    For Each ws In ThisWorkbook.Sheets
        If ws.name <> "Sheet1" And Left(ws.name, 5) <> "Chart" Then
            count = count + 1
        End If
    Next ws
    
    CountClassSheets = count
End Function

Private Sub InitializeClassSheet(ws As Worksheet)
    ' Clear sheet
    ws.Cells.Clear
    
    ' Add class header
    With ws.Range("A1")
        .Value = "CLASS: " & UCase(ws.name)
        .Font.Size = 16
        .Font.Bold = True
        .Font.Color = RGB(0, 100, 0)
    End With
    
    ' Add timestamp
    With ws.Range("A2")
        .Value = "Created: " & Now()
        .Font.Size = 9
        .Font.Color = RGB(100, 100, 100)
    End With
    
    ' Add column headers
    With ws.Range("A4:J4")
        .Value = Array("Roll No", "Student Name", "Mathematics", "Physics", _
                      "Chemistry", "Biology", "English", "Total", "Average", "Grade")
        .Font.Bold = True
        .Interior.Color = RGB(220, 230, 241)
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
    End With
    
    ' Add navigation back to main menu
    With ws.Range("M1")
        .Value = "? Back to Main"
        .Font.Color = RGB(0, 0, 255)
        .Font.Underline = xlUnderlineStyleSingle
    End With
    
    ' Add class statistics placeholder
    With ws.Range("M3:M10")
        .Value = Array("Class Statistics", "Students: 0", "Top Score: -", _
                      "Average: -", "Pass %: -", "Topper: -", "", "Click 'Calculate'")
        .Font.Size = 10
    End With
    
    ' Auto-fit columns
    ws.Columns.AutoFit
End Sub
Sub SmartAddStudents()
    ' Simple version without VBComponent reference
    
    ' 1. First, check if we're already on a class sheet
    If Left(ActiveSheet.name, 5) = "Class" Then
        ' We're on a class sheet, show the form
        ShowSimpleForm
        Exit Sub
    End If
    
    ' 2. Look for existing class sheets
    Dim classNum As Integer: classNum = 1
    
    Do While SheetExists("Class " & classNum)
        classNum = classNum + 1
    Loop
    
    ' 3. If no class sheets exist, create one
    If classNum = 1 Then
        Dim response As Integer
        response = MsgBox("No class sheets found. Create a new class?", vbYesNo + vbQuestion)
        
        If response = vbYes Then
            CreateNewClassSheet
            ThisWorkbook.Sheets("Class 1").Activate
            ShowSimpleForm
        End If
    Else
        ' 4. Ask which class to use
        Dim selectedClass As String
        selectedClass = InputBox("Enter class number to add students to:", "Select Class", "1")
        
        If selectedClass <> "" Then
            On Error Resume Next
            ThisWorkbook.Sheets("Class " & selectedClass).Activate
            On Error GoTo 0
            
            If Err.Number = 0 Then
                ShowSimpleForm
            Else
                MsgBox "Class " & selectedClass & " not found!", vbExclamation
            End If
        End If
    End If
End Sub

' ======== SIMPLE FORM SHOW ========
Sub ShowSimpleForm()
    On Error GoTo FormError
    
    ' Try to show the form with error handling
    frmMarkSheet.Show
    
    Exit Sub
    
FormError:
    If Err.Number = 424 Then
        MsgBox "Form 'frmMarkSheet' not found." & vbCrLf & _
               "Please create the form first.", vbExclamation
    Else
        MsgBox "Error showing form: " & Err.Description, vbCritical
    End If
End Sub

' ======== SHEET EXISTS FUNCTION ========
Function SheetExists(sheetName As String) As Boolean
    On Error Resume Next
    SheetExists = Not ThisWorkbook.Sheets(sheetName) Is Nothing
    On Error GoTo 0
End Function

' ======== NEW FUNCTION ========
Sub ShowMarkSheetForm()
    On Error GoTo FormError
    
    ' Check if form exists
    Dim formExists As Boolean
    formExists = False
    
    Dim comp As VBComponent
    For Each comp In ThisWorkbook.VBProject.VBComponents
        If comp.Type = vbext_ct_MSForm Then
            If comp.name = "frmMarkSheet" Then
                formExists = True
                Exit For
            End If
        End If
    Next comp
    
    If Not formExists Then
        MsgBox "Form frmMarkSheet not found. Please create it first.", vbExclamation
        Exit Sub
    End If
    
    ' Load and show the form
    Load frmMarkSheet
    frmMarkSheet.Show
    
    Exit Sub
    
FormError:
    MsgBox "Cannot load frmMarkSheet: " & Err.Description, vbCritical
End Sub






' ====== CALCULATION FUNCTIONS ======
Public Sub CalculateActiveClass()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    ' Get active class sheet
    Set ws = ActiveSheet
    If ws.name = "Sheet1" Then
        MsgBox "Please select a class sheet to calculate!", vbExclamation
        Exit Sub
    End If
    
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    
    If lastRow < 5 Then
        MsgBox "No student data found in '" & ws.name & "'!", vbExclamation
        Exit Sub
    End If
    
    ' Calculate for each student
    For i = 5 To lastRow
        If ws.Cells(i, 1).Value <> "" Then
            CalculateStudentRow ws, i
        End If
    Next i
    
    ' Update class statistics
    UpdateClassStatistics ws
    
    MsgBox "Calculations completed for " & ws.name & "!", vbInformation
End Sub

Private Sub CalculateStudentRow(ws As Worksheet, rowNum As Long)
    Dim total As Double
    Dim average As Double
    Dim grade As String
    
    ' Calculate total (columns C-G: Math to English)
    total = 0
    For col = 3 To 7
        If IsNumeric(ws.Cells(rowNum, col).Value) Then
            total = total + ws.Cells(rowNum, col).Value
        End If
    Next col
    
    ' Calculate average
    average = total / 5
    
    ' Determine grade
    Select Case average
        Case Is >= 90: grade = "A+"
        Case Is >= 80: grade = "A"
        Case Is >= 70: grade = "B+"
        Case Is >= 60: grade = "B"
        Case Is >= 50: grade = "C"
        Case Is >= 40: grade = "D"
        Case Else: grade = "F"
    End Select
    
    ' Write results
    ws.Cells(rowNum, 8).Value = total      ' Total
    ws.Cells(rowNum, 9).Value = average    ' Average
    ws.Cells(rowNum, 10).Value = grade     ' Grade
    
    ' Format
    With ws.Range(ws.Cells(rowNum, 8), ws.Cells(rowNum, 10))
        .HorizontalAlignment = xlCenter
        .NumberFormat = "0.00"
        .Borders.LineStyle = xlContinuous
    End With
End Sub

Private Sub UpdateClassStatistics(ws As Worksheet)
    Dim lastRow As Long
    Dim studentCount As Long
    Dim topScore As Double
    Dim classAverage As Double
    Dim topperName As String
    Dim i As Long
    
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    studentCount = 0
    topScore = 0
    classAverage = 0
    topperName = ""
    
    ' Calculate statistics
    For i = 5 To lastRow
        If ws.Cells(i, 1).Value <> "" Then
            studentCount = studentCount + 1
            classAverage = classAverage + ws.Cells(i, 9).Value
            
            If ws.Cells(i, 8).Value > topScore Then
                topScore = ws.Cells(i, 8).Value
                topperName = ws.Cells(i, 2).Value
            End If
        End If
    Next i
    
    If studentCount > 0 Then
        classAverage = classAverage / studentCount
    End If
    
    ' Update statistics box
    With ws.Range("M3:M10")
        .Value = Array("?? CLASS STATISTICS", _
                      "?? Students: " & studentCount, _
                      "?? Top Score: " & topScore, _
                      "?? Average: " & Format(classAverage, "0.00"), _
                      "? Pass %: " & CalculatePassPercentage(ws) & "%", _
                      "?? Topper: " & topperName, _
                      "", _
                      "?? Updated: " & Time())
        .Font.Size = 10
        .Font.Bold = True
        .Rows(1).Font.Color = RGB(0, 100, 0)
    End With
End Sub

Private Function CalculatePassPercentage(ws As Worksheet) As Double
    Dim lastRow As Long
    Dim passCount As Long
    Dim totalCount As Long
    Dim i As Long
    
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    passCount = 0
    totalCount = 0
    
    For i = 5 To lastRow
        If ws.Cells(i, 1).Value <> "" Then
            totalCount = totalCount + 1
            If ws.Cells(i, 9).Value >= 40 Then  ' Passing: 40%
                passCount = passCount + 1
            End If
        End If
    Next i
    
    If totalCount > 0 Then
        CalculatePassPercentage = Round((passCount / totalCount) * 100, 1)
    Else
        CalculatePassPercentage = 0
    End If
End Function

' ====== NAVIGATION FUNCTIONS ======
Public Sub GoToNextClass()
    Dim currentIndex As Integer
    Dim nextIndex As Integer
    
    currentIndex = ActiveSheet.index
    
    If currentIndex < ThisWorkbook.Sheets.count Then
        ThisWorkbook.Sheets(currentIndex + 1).Activate
    Else
        ThisWorkbook.Sheets(2).Activate  ' Go to first class sheet
    End If
End Sub

Public Sub GoToPrevClass()
    Dim currentIndex As Integer
    
    currentIndex = ActiveSheet.index
    
    If currentIndex > 2 Then
        ThisWorkbook.Sheets(currentIndex - 1).Activate
    Else
        ThisWorkbook.Sheets(ThisWorkbook.Sheets.count).Activate  ' Go to last class sheet
    End If
End Sub

Public Sub ViewAllClasses()
    Dim ws As Worksheet
    
    ' Create summary sheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("All Classes")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(2))
        ws.name = "All Classes"
    End If
    On Error GoTo 0
    
    ' Generate summary
    GenerateAllClassesSummary ws
    
    ws.Activate
End Sub

' ====== CLEAR/RESET ======
Public Sub ClearActiveClass()
    Dim ws As Worksheet
    
    Set ws = ActiveSheet
    
    If ws.name = "Sheet1" Then
        MsgBox "Cannot clear Main Menu!", vbExclamation
        Exit Sub
    End If
    
    If MsgBox("Clear all data from '" & ws.name & "'?", vbYesNo + vbQuestion) = vbYes Then
        InitializeClassSheet ws
        MsgBox "'" & ws.name & "' cleared successfully!", vbInformation
    End If
End Sub

Public Sub GenerateSummary()
    ' Create comprehensive report
    MsgBox "Summary report feature would show:" & vbCrLf & _
           "• All classes comparison" & vbCrLf & _
           "• Performance charts" & vbCrLf & _
           "• School/College statistics" & vbCrLf & _
           "• Export to PDF option", vbInformation
End Sub


Sub FixButtonMacros()
    Dim btn As Button
    
    For Each btn In ActiveSheet.Buttons
        Select Case btn.name
            Case "btnNewClass": btn.OnAction = "CreateNewClassSheet"
            Case "btnAddStudents": btn.OnAction = "SmartAddStudents"
            Case "btnCalculate": btn.OnAction = "CalculateActiveClass"
            Case "btnNextClass": btn.OnAction = "GoToNextClass"
            Case "btnPrevClass": btn.OnAction = "GoToPrevious"
            Case "btnClearClass": btn.OnAction = "ClearActiveClass"
            Case "btnViewAll": btn.OnAction = "ViewAllClasses"
            Case "btnSummary": btn.OnAction = "GenerateSummary"
        End Select
    Next btn
    
    MsgBox "Button macros fixed!", vbInformation
End Sub
