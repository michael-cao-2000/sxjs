VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ArticlePaymentForm 
   Caption         =   "稿费发放表"
   ClientHeight    =   2655
   ClientLeft      =   50
   ClientTop       =   380
   ClientWidth     =   6460
   OleObjectBlob   =   "ArticlePaymentForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ArticlePaymentForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Const ArticleNo_Col = 1

Const Title_Col = 2

Const Contact_Col = 3

Const ID_Col = 4

Const Address_Col = 5

Const ZipCode_Col = 6

Const Phone_Col = 7

Const Mail_Col = 8

Const PlanIssueNo_Col = 9

Const IssueNo_Col = 10

Const PublicationFee_Col = 11

Const AccepteDate_Col = 12

Const PublicationFeeReceivedDate_Col = 13

Const PaymentMethod_Col = 14

Const InvoiceNo_Col = 15

Const InvoiceSentDate_Col = 16

Const Confidential_Col = 17

Const Reviewer1_Col = 18

Const Reviewer1BackDate_Col = 19

Const Reviewer1FeePayedDate_Col = 20

Const Reviewer2_Col = 21

Const Reviewer2BackDate_Col = 22

Const Reviewer2FeePayedDate_Col = 23

Const Reviewer3_Col = 24

Const Reviewer3BackDate_Col = 25

Const Reviewer3FeePayedDate_Col = 26

Const Reviewer4_Col = 27

Const Reviewer4BackDate_Col = 28

Const Reviewer4FeePayedDate_Col = 29

Const AuthorCompany_Col = 30

Const FundingsNo_Col = 31

Const Dissertations_Col = 32

Const FirstAuthor_Col = 33

Private Articles(1000) As ArticleDetail

Private Count As Integer

Private OtherArticles(1000) As ArticleDetail

Private OtherCount As Integer

Private Sub UserForm_Initialize()
   
    txtIssueNo.SetFocus
    OKCommandButton.Enabled = True
    
End Sub


Private Sub CancelCommandButton_Click()
    Unload Me
End Sub


Private Sub OKCommandButton_Click()
    Dim IssueNo As String
    Dim IssueNoDate As Date
    Dim IssueYear As Integer
    Dim IssueMonth As Integer
    Dim strDate As String
    
    IssueNo = Trim(txtIssueNo.Value)
    If IssueNo = "" Then
        MsgBox "出版时间不能为空", vbExclamation
        txtIssueNo.SetFocus
    Else
        IssueNoDate = CDate(IssueNo)
        IssueYear = Year(IssueNoDate)
        IssueMonth = Month(IssueNoDate)
        strDate = Format(IssueNoDate, "yyyy/mm/dd")
        txtIssueNo.Text = strDate
        TotalCount = FindArticles(IssueYear, IssueMonth)
        If TotalCount = 0 Then
            MsgBox "没有找到出版刊期为（" & IssueNo & "）的文章", vbExclamation
        Else
            Set ReviewerSheet = FindWorksheetByName("审稿专家库")
            Set Sheet = FindWorksheetByName("稿费发放表")
            If Sheet Is Nothing Then
                Set Sheet = Workbooks.Application.Worksheets.Add(, ReviewerSheet)
                Sheet.Name = "稿费发放表"
            End If
            FillWorksheet Sheet
            Unload Me
        End If
    End If

End Sub

Private Sub FillWorksheet(Sheet As Variant)
    
    Sheet.Cells.ClearContents
    Sheet.Cells.WrapText = True
    
    FillWorksheetCol Sheet, 1, "Contact", "姓名", 10
    FillWorksheetCol Sheet, 2, "ArticleNo", "编号", 10
    FillWorksheetCol Sheet, 3, "Title", "题目", 50
    
    Sheet.Cells(1, 4).Value2 = "稿费金额"
    Sheet.Cells(1, 5).Value2 = "汇费金额"
    Sheet.Cells(1, 6).Value2 = "领款人签字或邮局回单号码"

    FillWorksheetCol Sheet, 7, "ID", "身份证号码", 20
    FillWorksheetCol Sheet, 8, "Address", "地址", 40
    FillWorksheetCol Sheet, 9, "ZipCode", "邮编", 10

'    Sheet.Cells(1, ID_Col).Value2 = "身份证号码"
'    Sheet.Cells(1, Address_Col).Value2 = "地址"
'    Sheet.Cells(1, ZipCode_Col).Value2 = "邮编"
'    Sheet.Cells(1, Phone_Col).Value2 = "电话"
'    Sheet.Cells(1, Mail_Col).Value2 = "电子邮箱"
'    Sheet.Cells(1, PlanIssueNo_Col).Value2 = "预计刊期"
'    Sheet.Cells(1, IssueNo_Col).Value2 = "出版刊期"
'    Sheet.Cells(1, PublicationFee_Col).Value2 = "版面费"
'    Sheet.Cells(1, AccepteDate_Col).Value2 = "录用通知日期"
'    Sheet.Cells(1, PublicationFeeReceivedDate_Col).Value2 = "到账日期"
'    Sheet.Cells(1, PaymentMethod_Col).Value2 = "付款方式"
'    Sheet.Cells(1, InvoiceNo_Col).Value2 = "发票号"
'    Sheet.Cells(1, InvoiceSentDate_Col).Value2 = "发票寄出日期"
'    Sheet.Cells(1, Confidential_Col).Value2 = "保密证明"
'    Sheet.Cells(1, Reviewer1_Col).Value2 = "审稿人1"
'    Sheet.Cells(1, Reviewer1BackDate_Col).Value2 = "审回时间1"
'    Sheet.Cells(1, Reviewer1FeePayedDate_Col).Value2 = "审稿费1发放"
'    Sheet.Cells(1, Reviewer2_Col).Value2 = "审稿人2"
'    Sheet.Cells(1, Reviewer2BackDate_Col).Value2 = "审回时间2"
'    Sheet.Cells(1, Reviewer2FeePayedDate_Col).Value2 = "审稿费2发放"
'    Sheet.Cells(1, Reviewer3_Col).Value2 = "审稿人3"
'    Sheet.Cells(1, Reviewer3BackDate_Col).Value2 = "审稿时间3"
'    Sheet.Cells(1, Reviewer3FeePayedDate_Col).Value2 = "审稿费3发放"
'    Sheet.Cells(1, Reviewer4_Col).Value2 = "审稿人4"
'    Sheet.Cells(1, Reviewer4BackDate_Col).Value2 = "审稿时间4"
'    Sheet.Cells(1, Reviewer4FeePayedDate_Col).Value2 = "审稿费4发放"
'    Sheet.Cells(1, AuthorCompany_Col).Value2 = "作者单位"
'    Sheet.Cells(1, FundingsNo_Col).Value2 = "基金论文"
'    Sheet.Cells(1, Dissertations_Col).Value2 = "硕/博士论文"
'    Sheet.Cells(1, FirstAuthor_Col) = "第一作者职称"
    
End Sub

Private Sub FillWorksheetCol(ByRef Sheet As Variant, Col As Integer, Name As String, Title As String, Optional Width As Double)
    Dim Value2 As String
    Dim hidden As Boolean
    
    hidden = Not Visible
    
    Set cell = Sheet.Cells(1, Col)
    cell.ColumnWidth = Width
    Sheet.Cells(1, Col).Value2 = Title
        
    For I = 1 To Count
        Set Detail = Articles(I)
        Value = Chr$(9) & Detail.GetByName(Name)
        Sheet.Cells(I + 1, Col).Value2 = Value
    Next I
    
    For I = 1 To OtherCount
        Set Detail = OtherArticles(I)
        Value = Chr$(9) & Detail.GetByName(Name)
        Sheet.Cells(Count + I + 1, Col).Value2 = Value
    Next I

End Sub


'根据出版刊期返回稿件清单
Private Function FindArticles(IssueYear As Integer, IssueMonth As Integer) As Integer
    Dim ActualIssueNo As Date
    Dim ActualYear As Integer
    Dim ActualMonth As Integer

    OKCommandButton.Enabled = False
    FrameProgress.Visible = True

    Count = FindArticlesFromSheet("来稿登记", IssueYear, IssueMonth, Articles)
    OtherCount = FindArticlesFromSheet("表外录用来稿登记 ", IssueYear, IssueMonth, OtherArticles)
    
    FindArticles = Count + OtherCount

    OKCommandButton.Enabled = True
    FrameProgress.Visible = False
    
End Function

Private Function FindArticlesFromSheet(SheetName As String, IssueYear As Integer, IssueMonth As Integer, ByRef Articles() As ArticleDetail) As Integer

    Set Sheet = FindWorksheetByName(SheetName)
    If Sheet Is Nothing Then
        Exit Function
    End If

    J = 0
    For I = 2 To 10000
        ArticleNo = Trim(Sheet.Cells(I, 1).Value2)
        Title = Trim(Sheet.Cells(I, 2).Value2)
        If ArticleNo = "" And Title = "" Then
            UpdateProgressBar (100)
            Exit For
        End If
            
        ActualIssueNo = Sheet.Cells(I, IssueNo_Col).Value2
        Debug.Print ActualIssueNo
        ActualYear = Year(ActualIssueNo)
        ActualMonth = Month(ActualIssueNo)
        Debug.Print ActualYear & "  " & ActualMonth
            
        If ActualYear = IssueYear And ActualMonth = IssueMonth Then
            Set Article = New ArticleDetail
            With Article
                .ArticleNo = ArticleNo
                .IssueNo = ActualIssueNo
                .Title = Title
                .Contact = Sheet.Cells(I, Contact_Col).Value2
                .ZipCode = Sheet.Cells(I, ZipCode_Col).Value2
                .Address = Sheet.Cells(I, Address_Col).Value2
                .ID = Sheet.Cells(I, ID_Col).Value2
            End With
            J = J + 1
            Set Articles(J) = Article
        End If
'姓名    证件类型    证件号码
'收款人邮编  收款人姓名  收款人地址
        
        
        If (I Mod 500) = 0 Then
            UpdateProgressBar (I / 10000)
        End If
    Next I
    
    FindArticlesFromSheet = J

End Function


Private Sub UpdateProgressBar(PctDone As Single)
    FrameProgress.Caption = Format(PctDone, "0%")
    LabelProgress.Width = PctDone * (FrameProgress.Width - 10)
    DoEvents
End Sub

