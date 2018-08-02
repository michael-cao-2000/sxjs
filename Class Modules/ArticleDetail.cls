VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArticleDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'编号
Public ArticleNo As String

'题目
Public Title As String

'联系人
Public Contact As String

'身份证号码
Public ID As String

'联系地址
Public Address As String

'邮编
Public ZipCode As String

'电话
Public Phone As String

'电子邮箱
Public Mail As String

'预计刊期
Public PlanIssueNo As String

'出版刊期
Public IssueNo As String

'版面费
Public PublicationFee As Integer

'录用通知日期
Public AccepteDate As Date

'到账日期
Public PublicationFeeReceivedDate As Date

'付款方式
Public PaymentMethod As String

'发票号
Public InvoiceNo As String

'发票寄出日期
Public InvoiceSentDate As String

'保密证明
Public Confidential As String

Public Reviewer1 As String
Public Reviewer1BackDate As Date
Public Reviewer1FeePayedDate As Date

Public Reviewer2 As String
Public Reviewer2BackDate As Date
Public Reviewer2FeePayedDate As Date

Public Reviewer3 As String
Public Reviewer3BackDate As Date
Public Reviewer3FeePayedDate As Date

'审稿人4
Public Reviewer4 As String

'审稿时间4
Public Reviewer4BackDate As Date

'审稿费4发放
Public Reviewer4FeePayedDate As Date

'作者单位
Public AuthorCompany As String

'基金论文
Public FundingsNo As String

'硕/博士论文
Public Dissertations As String

'第一作者职称
Public FirstAuthor As String



Public Function GetByName(Name As String) As String
    Select Case Name
        Case "ArticleNo":
            GetByName = ArticleNo
        Case "Title":
            GetByName = Title
        Case "Contact":
            GetByName = Contact
        Case "ZipCode":
            GetByName = ZipCode
        Case "Address":
            GetByName = Address
        Case "ID":
            GetByName = ID
        Case "Phone"
            GetByName = Phone
        
    End Select
End Function

