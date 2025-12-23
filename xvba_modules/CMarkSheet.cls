VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMarkSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Main MarkSheet Class
Private pStudents As Collection
Private pSubjects As Collection
Private pTotalStudents As Integer
Private pTotalSubjects As Integer

Private Sub Class_Initialize()
    Set pStudents = New Collection
    Set pSubjects = New Collection
    pTotalStudents = 0
    pTotalSubjects = 0
End Sub

' Add a student
Public Sub AddStudent(student As CStudent)
    pStudents.Add student
    pTotalStudents = pTotalStudents + 1
End Sub

' Add a subject
Public Sub AddSubject(subject As CSubject)
    pSubjects.Add subject
    pTotalSubjects = pTotalSubjects + 1
End Sub

' Get student by index
Public Function GetStudent(index As Integer) As CStudent
    Set GetStudent = pStudents(index)
End Function

' Get subject by index
Public Function GetSubject(index As Integer) As CSubject
    Set GetSubject = pSubjects(index)
End Function

' Calculate ranks based on total marks
Public Sub CalculateRanks()
    Dim i As Integer, j As Integer
    Dim tempStudent As CStudent
    Dim sortedStudents() As CStudent
    
    ' Create array for sorting
    ReDim sortedStudents(1 To pTotalStudents)
    
    ' Copy students to array
    For i = 1 To pTotalStudents
        Set sortedStudents(i) = pStudents(i)
    Next i
    
    ' Bubble sort by total marks (descending)
    For i = 1 To pTotalStudents - 1
        For j = i + 1 To pTotalStudents
            If sortedStudents(i).total < sortedStudents(j).total Then
                Set tempStudent = sortedStudents(i)
                Set sortedStudents(i) = sortedStudents(j)
                Set sortedStudents(j) = tempStudent
            End If
        Next j
    Next i
    
    ' Assign ranks
    Dim currentRank As Integer
    Dim previousTotal As Double
    
    currentRank = 1
    previousTotal = -1
    
    For i = 1 To pTotalStudents
        If sortedStudents(i).total <> previousTotal Then
            sortedStudents(i).Rank = currentRank
            previousTotal = sortedStudents(i).total
        Else
            sortedStudents(i).Rank = currentRank - 1
        End If
        currentRank = currentRank + 1
    Next i
End Sub

' Generate report on worksheet
Public Sub GenerateReport(ws As Worksheet)
    Dim i As Integer, j As Integer
    Dim student As CStudent
    Dim subject As CSubject
    
    ' Clear worksheet
    ws.Cells.Clear
    
    ' Set up headers
    With ws
        .Range("A1").Value = "Student ID"
        .Range("B1").Value = "Student Name"
        
        ' Subject headers
        For i = 1 To pTotalSubjects
            .Cells(1, 2 + i).Value = GetSubject(i).SubjectCode
        Next i
        
        .Cells(1, 3 + pTotalSubjects).Value = "Total"
        .Cells(1, 4 + pTotalSubjects).Value = "Average"
        .Cells(1, 5 + pTotalSubjects).Value = "Rank"
        
        ' Populate student data
        For i = 1 To pTotalStudents
            Set student = GetStudent(i)
            
            .Cells(i + 1, 1).Value = student.studentID
            .Cells(i + 1, 2).Value = student.studentName
            
            ' Subject marks
            For j = 1 To pTotalSubjects
                .Cells(i + 1, 2 + j).Value = student.marks(j - 1)
            Next j
            
            ' Totals and ranks
            .Cells(i + 1, 3 + pTotalSubjects).Value = student.total
            .Cells(i + 1, 4 + pTotalSubjects).Value = Format(student.average, "0.00")
            .Cells(i + 1, 5 + pTotalSubjects).Value = student.Rank
        Next i
        
        ' Auto-fit columns
        .Columns.AutoFit
        
        ' Apply formatting
        With .Range(.Cells(1, 1), .Cells(pTotalStudents + 1, 5 + pTotalSubjects))
            .Borders.LineStyle = xlContinuous
            .HorizontalAlignment = xlCenter
        End With
        
        ' Format header
        With .Rows(1)
            .Font.Bold = True
            .Interior.Color = RGB(200, 220, 250)
        End With
    End With
End Sub

' Get statistics
Public Function GetStatistics() As Variant
    Dim stats(1 To 5) As String
    Dim topTotal As Double
    Dim bottomTotal As Double
    
    If pTotalStudents > 0 Then
        ' Find top and bottom totals
        topTotal = GetStudent(1).total
        bottomTotal = GetStudent(1).total
        
        For i = 1 To pTotalStudents
            If GetStudent(i).total > topTotal Then topTotal = GetStudent(i).total
            If GetStudent(i).total < bottomTotal Then bottomTotal = GetStudent(i).total
        Next i
        
        stats(1) = "Total Students: " & pTotalStudents
        stats(2) = "Total Subjects: " & pTotalSubjects
        stats(3) = "Highest Total: " & topTotal
        stats(4) = "Lowest Total: " & bottomTotal
        stats(5) = "Average of Class: " & Format(GetClassAverage(), "0.00")
    End If
    
    GetStatistics = stats
End Function

' Calculate class average
Private Function GetClassAverage() As Double
    Dim sum As Double
    Dim i As Integer
    
    For i = 1 To pTotalStudents
        sum = sum + GetStudent(i).total
    Next i
    
    If pTotalStudents > 0 Then
        GetClassAverage = sum / pTotalStudents
    Else
        GetClassAverage = 0
    End If
End Function

Public Property Get studentCount() As Integer
    studentCount = pTotalStudents
End Property

Public Property Get subjectCount() As Integer
    subjectCount = pTotalSubjects
End Property
