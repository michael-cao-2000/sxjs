VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ReviewPayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private strName As String
Private dblFee As Double
Private dblPostage As Double

Public Property Get Name() As String
     Name = strName
End Property

Public Property Let Name(Value As String)
    strName = Value
End Property

Public Property Get Fee() As Double
     Fee = dblFee
End Property

Public Property Let Fee(Value As Double)
    dblFee = Value
End Property

Public Property Get Postage() As Double
    Postage = dblPostage
End Property

Public Property Let Postage(Value As Double)
    dblPostage = Value
End Property

