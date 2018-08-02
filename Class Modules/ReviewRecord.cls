VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ReviewRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private strName As String
Private strArticleNo As String
Private strTime As String
Private strTitle As String


Public Property Get Name() As String
     Name = strName
End Property

Public Property Let Name(Value As String)
    strName = Value
End Property

Public Property Get ArticleNo() As String
     ArticleNo = strArticleNo
End Property

Public Property Let ArticleNo(Value As String)
    strArticleNo = Value
End Property

Public Property Get ReviewTime() As String
    ReviewTime = strTime
End Property

Public Property Let ReviewTime(Value As String)
    strTime = Value
End Property

Public Property Get Title() As String
    Title = strTitle
End Property

Public Property Let Title(Value As String)
    strTitle = Value
End Property
