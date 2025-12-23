VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CStudent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Student Class
Private pID As String
Private pName As String
Private pMarks() As Double
Private pTotal As Double
Private pAverage As Double
Private pRank As Integer

Public Property Get studentID() As String
    studentID = pID
End Property

Public Property Let studentID(Value As String)
    pID = Value
End Property

Public Property Get studentName() As String
    studentName = pName
End Property

Public Property Let studentName(Value As String)
    pName = Value
End Property

Public Property Get marks() As Variant
    marks = pMarks
End Property

Public Property Let marks(Value As Variant)
    If IsArray(Value) Then
        pMarks = Value
    End If
End Property

Public Property Get total() As Double
    total = pTotal
End Property

Public Property Let total(Value As Double)
    pTotal = Value
End Property

Public Property Get average() As Double
    average = pAverage
End Property

Public Property Let average(Value As Double)
    pAverage = Value
End Property

Public Property Get Rank() As Integer
    Rank = pRank
End Property

Public Property Let Rank(Value As Integer)
    pRank = Value
End Property

' Calculate total and average
Public Sub CalculateResults()
    Dim i As Integer
    Dim sum As Double
    
    If Not IsArrayEmpty(pMarks) Then
        For i = LBound(pMarks) To UBound(pMarks)
            sum = sum + pMarks(i)
        Next i
        
        pTotal = sum
        pAverage = sum / (UBound(pMarks) - LBound(pMarks) + 1)
    End If
End Sub

' Check if array is empty
Private Function IsArrayEmpty(arr As Variant) As Boolean
    Dim i As Integer
    On Error GoTo ErrHandler
    
    i = UBound(arr, 1)
    IsArrayEmpty = False
    Exit Function
    
ErrHandler:
    IsArrayEmpty = True
End Function
