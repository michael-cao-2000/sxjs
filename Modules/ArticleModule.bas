Attribute VB_Name = "ArticleModule"

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


'从稿费发放表中提取稿件编号清单
Public Function GetArticleNosToBePaid(ByRef Articles() As String) As Integer

    Set Sheet = FindWorksheetByName("稿费发放表")
    If Sheet Is Nothing Then
        MsgBox "没有找到‘稿费发放表’，请先生成‘稿费发放表’", vbExclamation
        Exit Function
    End If
    
    I = 1
    EmptyCount = 0
    For N = 2 To 1000
        Set cell = Sheet.Cells(N, 2)
        Title = Trim(cell.Value2)
        If (Title = "") Then
            EmptyCount = EmptyCount + 1
            If (EmptyCount > 5) Then
                Exit For
            End If
        Else
            EmptyCount = 0
            Name = Trim(Sheet.Cells(N, 1).Value2)
            If Name = "" Then
            Else
                Articles(I) = Name
                I = I + 1
            End If
        End If
    Next N
    
    GetArticleNosToBePaid = I

End Function


