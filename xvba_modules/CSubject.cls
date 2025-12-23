VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSubject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Subject Class
Private pCode As String
Private pName As String
Private pMaxMarks As Double
Private pPassMarks As Double

Public Property Get SubjectCode() As String
    SubjectCode = pCode
End Property

Public Property Let SubjectCode(Value As String)
    pCode = Value
End Property

Public Property Get SubjectName() As String
    SubjectName = pName
End Property

Public Property Let SubjectName(Value As String)
    pName = Value
End Property

Public Property Get MaxMarks() As Double
    MaxMarks = pMaxMarks
End Property

Public Property Let MaxMarks(Value As Double)
    pMaxMarks = Value
End Property

Public Property Get PassMarks() As Double
    PassMarks = pPassMarks
End Property

Public Property Let PassMarks(Value As Double)
    pPassMarks = Value
End Property
