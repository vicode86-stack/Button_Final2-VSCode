VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMarkSheet 
   Caption         =   "UserForm1"
   ClientHeight    =   12495
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19440
   OleObjectBlob   =   "frmMarkSheet.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMarkSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    ' This works with your current form (no combobox)
    Me.Caption = "Student Entry Form"
    
    ' Your form probably has these default names:
    ' TextBox1 = Student ID
    ' TextBox2 = Student Name
    ' TextBox3 = Marks
    
    TextBox1.Text = ""  ' Clear ID
    TextBox2.Text = ""  ' Clear Name
    TextBox3.Text = "Enter 5 marks separated by commas:" & vbCrLf & _
                    "Example: 85, 78, 92, 88, 90"
    TextBox3.MultiLine = True
    
    ' Focus on ID field
    TextBox1.SetFocus
End Sub

' "Add Student" button (probably CommandButton1)
Private Sub CommandButton1_Click()
    Dim ws As Worksheet
    
    ' Check which class sheet is active
    Set ws = ActiveSheet
    
    ' Make sure we're on a Class sheet
    If Left(ws.name, 5) <> "Class" Then
        MsgBox "ERROR: You must be on a Class sheet!" & vbCrLf & _
               "Example: Go to 'Class 1' or 'Class 2' sheet first.", _
               vbCritical, "Wrong Sheet"
        Exit Sub
    End If
    
    ' Add the student
    Call AddStudentToSheet(ws)
    
    MsgBox "Student added to " & ws.name & "!", vbInformation
End Sub

' "Generate Report" button (probably CommandButton2)
Private Sub CommandButton2_Click()
    MsgBox "Report feature - Coming soon!", vbInformation
End Sub

' "Close" button (probably CommandButton3)
Private Sub CommandButton3_Click()
    Unload Me
End Sub

' Helper function to add student data
Private Sub AddStudentToSheet(ws As Worksheet)
    Dim nextRow As Long
    Dim marksStr As String
    Dim marksArr() As String
    Dim i As Integer
    
    ' Validate inputs
    If TextBox1.Text = "" Then
        MsgBox "Please enter Student ID!", vbExclamation
        TextBox1.SetFocus
        Exit Sub
    End If
    
    If TextBox2.Text = "" Then
        MsgBox "Please enter Student Name!", vbExclamation
        TextBox2.SetFocus
        Exit Sub
    End If
    
    ' Parse marks
    marksStr = TextBox3.Text
    ' Remove the example text if still there
    If InStr(marksStr, "Example:") > 0 Then
        MsgBox "Please enter actual marks!", vbExclamation
        TextBox3.SetFocus
        Exit Sub
    End If
    
    ' Clean marks string
    marksStr = Replace(marksStr, vbCrLf, ",")
    marksStr = Replace(marksStr, " ", "")
    marksArr = Split(marksStr, ",")
    
    ' Check for 5 marks
    If UBound(marksArr) <> 4 Then
        MsgBox "Please enter exactly 5 marks!" & vbCrLf & _
               "Example: 85,78,92,88,90", vbExclamation
        TextBox3.SetFocus
        Exit Sub
    End If
    
    ' Find next empty row
    nextRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row + 1
    If nextRow < 5 Then nextRow = 5
    
    ' Write to sheet
    With ws
        .Cells(nextRow, 1).Value = TextBox1.Text  ' Roll No
        .Cells(nextRow, 2).Value = TextBox2.Text  ' Name
        .Cells(nextRow, 3).Value = Val(marksArr(0))  ' Math
        .Cells(nextRow, 4).Value = Val(marksArr(1))  ' Physics
        .Cells(nextRow, 5).Value = Val(marksArr(2))  ' Chemistry
        .Cells(nextRow, 6).Value = Val(marksArr(3))  ' Biology
        .Cells(nextRow, 7).Value = Val(marksArr(4))  ' English
    End With
    
    ' Clear form for next entry
    TextBox1.Text = ""
    TextBox2.Text = ""
    TextBox3.Text = "Enter 5 marks separated by commas:" & vbCrLf & _
                    "Example: 85, 78, 92, 88, 90"
    TextBox1.SetFocus
End Sub

