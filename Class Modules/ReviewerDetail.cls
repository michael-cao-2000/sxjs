VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ReviewerDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private strName As String
Private strID As String
Private strAddress As String
Private strZipCode As String
Private strPhone As String
Private strMail As String
Private strCompany As String

Public Property Get Name() As String
     Name = strName
End Property

Public Property Let Name(Value As String)
    strName = Value
End Property

Public Property Get ID() As String
     ID = strID
End Property

Public Property Let ID(Value As String)
    strID = Value
End Property

Public Property Get Address() As String
    Address = strAddress
End Property

Public Property Let Address(Value As String)
    strAddress = Value
End Property

Public Property Get ZipCode() As String
    ZipCode = strZipCode
End Property

Public Property Let ZipCode(Value As String)
    strZipCode = Value
End Property

Public Property Get Phone() As String
    Phone = strPhone
End Property

Public Property Let Phone(Value As String)
    strPhone = Value
End Property

Public Property Get Mail() As String
    Mail = strMail
End Property

Public Property Let Mail(Value As String)
    strMail = Value
End Property

Public Property Get Company() As String
    Company = strCompany
End Property

Public Property Let Company(Value As String)
    strCompany = Value
End Property


