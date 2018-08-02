Attribute VB_Name = "ReportModule"
Public Sub OnGeneateArticlePaymentTable()

    ArticlePaymentForm.Show vbModal
    
End Sub

Public Sub OnGeneateReviewerFeeTable()

    ReviewerFeeForm.Show vbModal
    
End Sub

'大宗汇款-审稿费
Public Sub OnGeneateRemittanceReviewerTable()

    Dim I, Num, EmptyCount As Integer
    Dim Title, Name As String
    Dim Reviewers(1000) As ReviewPayment
    Dim currDate As Date
    Dim currDateStr As String
    Dim sumFee As Double, postageSum As Double
    Dim fullPathName As String
    Dim fileTypeName As String
        
    fileTypeName = "大宗汇款-审稿费"
    
    currDate = Date
    currDateStr = Format(currDate, "yyyy-mm-dd")
    
    Num = GetReviewers(Reviewers)
    If Num = 0 Then
        Exit Sub
    End If
    
    Set ReviewerDict = GetReviewerDict()
    
    '计算总金额加会费
    For I = 1 To Num - 1
        Name = Reviewers(I).Name
        If (Reviewers(I).Postage <> 0) Then
            sumFee = sumFee + Reviewers(I).Fee
            sumPostage = sumPostage + Reviewers(I).Postage
        End If
    Next I

    UserProfile = Environ("UserProfile") & "\Documents"
    Set fs = CreateObject("Scripting.FileSystemObject")
    fullPathName = UserProfile & "\大宗汇款-审稿费(" & currDateStr & ").csv"
    On Error Resume Next
    Set file = fs.CreateTextFile(fullPathName, True)
    If Err.Number > 0 Then
        MsgBox "大宗汇款-审稿费文件 " & fullPathName & " 已经被打开，无法继续当前操作，请先关闭文件后再执行 (错误代码：" & Err.Number & ")", vbCritical
        Exit Sub
    End If
    file.writeLine "商户代码,文件种类,总笔数,总金额"
    file.writeLine Chr$(9) & "310000000,0,0," & (sumFee + sumPostage)
    file.writeLine "汇款金额,收款人邮编,收款人姓名,收款人地址,附言"

    For I = 1 To Num - 1
        Name = Reviewers(I).Name
        Set Details = ReviewerDict.Item(Name)
        If (Reviewers(I).Postage <> 0) Then
            file.Write Reviewers(I).Fee & ","
            file.Write Chr$(9) & Details.ZipCode & ","
            file.Write Name & ","
            file.writeLine Details.Address & ","
        End If
    Next I
    
    file.Close
        
    PromptSuccess fileTypeName, fullPathName
    
End Sub

'大宗汇款-稿费
Public Sub OnGeneateRemittanceAuthorTable()

    Dim Title As String
    Dim currDate As Date
    Dim currDateStr As String
    Dim fullPathName As String
    Dim dblPaySum As Double, dblPostageSum As Double
    Dim fileTypeName As String
    
    fileTypeName = "大宗汇款-稿费"
    
    currDate = Date
    currDateStr = Format(currDate, "yyyy-mm-dd")
    
    Set Sheet = FindWorksheetByName("稿费发放表")
    If Sheet Is Nothing Then
        MsgBox "没有找到‘稿费发放表’，请先生成‘稿费发放表’", vbExclamation
        Exit Sub
    End If
    
    '计算稿费总金额 +　邮费
    For I = 1 To 1000
        Title = Trim(Sheet.Cells(I + 1, 3).Value2)
        If (Title = "") Then
            Exit For
        End If
        
        If (Sheet.Cells(I + 1, 5) > 0) Then
            dblPaySum = dblPaySum + Sheet.Cells(I + 1, 4)
            dblPostageSum = dblPostageSum + Sheet.Cells(I + 1, 5)
        End If
    Next I
    
    
    UserProfile = Environ("UserProfile") & "\Documents"
    Set fs = CreateObject("Scripting.FileSystemObject")
    fullPathName = UserProfile & "\大宗汇款-稿费(" & currDateStr & ").csv"
        
    On Error Resume Next
    Set file = fs.CreateTextFile(fullPathName, True)
    If Err.Number > 0 Then
        MsgBox "大宗汇款-稿费文件 " & fullPathName & " 已经被打开，无法继续当前操作，请先关闭文件后再执行 (错误代码：" & Err.Number & ")", vbCritical
        Exit Sub
    End If
    
    file.writeLine "商户代码,文件种类,总笔数,总金额"
    file.writeLine Chr$(9) & "310000000,0,0," & (dblPaySum + dblPostageSum)
    file.writeLine "汇款金额,收款人邮编,收款人姓名,收款人地址,附言"
    
    For I = 1 To 1000
        Title = Trim(Sheet.Cells(I + 1, 3).Value2)
        If (Title = "") Then
            Exit For
        End If
        
        file.Write Sheet.Cells(I + 1, 4) & ","
        file.Write Sheet.Cells(I + 1, 9).Value2 & ","
        file.Write Sheet.Cells(I + 1, 1).Value2 & ","
        file.writeLine Sheet.Cells(I + 1, 8).Value2 & ","
    Next I
    
    file.Close
    
    PromptSuccess fileTypeName, fullPathName
    
End Sub

'劳务发票申请表-审稿费
Public Sub OnGeneateServiceFeeReviewerTable()
    Dim I, Num, EmptyCount As Integer
    Dim Title, Name As String
    Dim currDate As Date
    Dim currDateStr As String
    Dim Reviewers(1000) As ReviewPayment
    Dim fileTypeName As String
    Dim fullPathName As String
    
    fileTypeName = "劳务发票申请表-审稿费"

    currDate = Date
    currDateStr = Format(currDate, "yyyy-mm-dd")

    Num = GetReviewers(Reviewers)
    If Num = 0 Then
        Exit Sub
    End If
    
    Set ReviewerDict = GetReviewerDict()

    UserProfile = Environ("UserProfile") & "\Documents"
    Set fs = CreateObject("Scripting.FileSystemObject")
    fullPathName = UserProfile & "\" & fileTypeName & "(" & currDateStr & ").csv"
    
    On Error Resume Next
    Set file = fs.CreateTextFile(fullPathName, True)
    If Err.Number > 0 Then
        MsgBox "大宗汇款-审稿费文件 " & fullPathName & " 已经被打开，无法继续当前操作，请先关闭文件后再执行 (错误代码：" & Err.Number & ")", vbCritical
        Exit Sub
    End If
    
    Set file = fs.CreateTextFile(UserProfile & "\劳务发票申请表-审稿费(" & currDateStr & ").csv", True)
    file.writeLine "中国科学院声学研究所东海研究站劳务发票申请表"
    file.writeLine "序号,姓名,证件类型,证件号码,劳务内容,所属期间,金额（元）"

    For I = 1 To Num - 1
        Name = Reviewers(I).Name
        file.Write I & ","
        file.Write Name & ","
        file.Write "身份证" & ","
        file.Write Chr$(9) & ReviewerDict.Item(Name).ID & ","
        file.writeLine ",,"
    Next I
    
    file.Close
    
    PromptSuccess fileTypeName, fullPathName
    
End Sub

'劳务发票申请表-稿费
Public Sub OnGeneateServiceFeeAuthorTable()
    Dim Title As String
    Dim currDate As Date
    Dim currDateStr As String
    Dim fileTypeName As String
    Dim fullPathName As String
    
    fileTypeName = "劳务发票申请表-稿费"
    
    currDate = Date
    currDateStr = Format(currDate, "yyyy-mm-dd")
    
    Set Sheet = FindWorksheetByName("稿费发放表")
    If Sheet Is Nothing Then
        MsgBox "没有找到‘稿费发放表’，请先生成‘稿费发放表’", vbExclamation
        Exit Sub
    End If
    
    UserProfile = Environ("UserProfile") & "\Documents"
    Set fs = CreateObject("Scripting.FileSystemObject")
    fullPathName = UserProfile & "\劳务发票申请表-稿费(" & currDateStr & ").csv"
    On Error Resume Next
    Set file = fs.CreateTextFile(fullPathName, True)
    If Err.Number > 0 Then
        MsgBox "劳务发票申请表-稿费文件 " & fullPathName & " 已经被打开，无法继续当前操作，请先关闭文件后再执行 (错误代码：" & Err.Number & ")", vbCritical
        Exit Sub
    End If
    
    file.writeLine "中国科学院声学研究所东海研究站劳务发票申请表"
    file.writeLine "序号,姓名,证件类型,证件号码,劳务内容,所属期间,金额（元）"
    
    For I = 1 To 1000
        Title = Trim(Sheet.Cells(I + 1, 3).Value2)
        If (Title = "") Then
            Exit For
        End If
        file.Write I & ","
        file.Write Sheet.Cells(I + 1, 1).Value2 & ","
        file.Write "身份证,"
        file.Write Sheet.Cells(I + 1, 7).Value2 & ","
        file.writeLine ",,"
    Next I
    
    file.Close
    
    PromptSuccess fileTypeName, fullPathName

End Sub


Sub PromptSuccess(ByRef fileTypeName As String, ByRef fullPathName As String)
    If (vbYes = MsgBox("生成了" & fileTypeName & "文件" & vbCrLf & fullPathName & vbCrLf & "需要现在打开吗？", vbQuestion & vbYesNo)) Then
        Workbooks.Open fullPathName
    End If
End Sub

