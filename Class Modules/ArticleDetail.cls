VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArticleDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'���
Public ArticleNo As String

'��Ŀ
Public Title As String

'��ϵ��
Public Contact As String

'����֤����
Public ID As String

'��ϵ��ַ
Public Address As String

'�ʱ�
Public ZipCode As String

'�绰
Public Phone As String

'��������
Public Mail As String

'Ԥ�ƿ���
Public PlanIssueNo As String

'���濯��
Public IssueNo As String

'�����
Public PublicationFee As Integer

'¼��֪ͨ����
Public AccepteDate As Date

'��������
Public PublicationFeeReceivedDate As Date

'���ʽ
Public PaymentMethod As String

'��Ʊ��
Public InvoiceNo As String

'��Ʊ�ĳ�����
Public InvoiceSentDate As String

'����֤��
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

'�����4
Public Reviewer4 As String

'���ʱ��4
Public Reviewer4BackDate As Date

'����4����
Public Reviewer4FeePayedDate As Date

'���ߵ�λ
Public AuthorCompany As String

'��������
Public FundingsNo As String

'˶/��ʿ����
Public Dissertations As String

'��һ����ְ��
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
