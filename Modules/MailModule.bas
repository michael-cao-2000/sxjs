Attribute VB_Name = "MailModule"
'审稿通知
Public Sub OnSendReviewEmail()

    Dim Row, Col As Integer
    Dim Email As String
    Dim Title As String
    Dim Reviewer As String, CC As String
    Dim currDate, expectDate As Date
    Dim currDateStr, expectDateStr As String
    
    currDate = Date
    currDateStr = Format(currDate, "yyyy-mm-dd")
    expectDate = DateAdd("m", 1, currDate)
    expectDateStr = Format(expectDate, "yyyy-mm-dd")
    
    Set cell = Worksheets.Application.ActiveCell
    Row = cell.Row
    Col = cell.Column
    Set Sheet = cell.Worksheet
    
    Set NoCell = Sheet.Cells(Row, Constants.NO_COL)
    ArticleNo = NoCell.Value2
    
    Set TitleCell = Sheet.Cells(Row, Constants.Title_Col)
    Title = TitleCell.Value2
    
    CC = FindCCByName("审稿抄送")
    
    If (Col <> Constants.Reviewer1_Col And Col <> Constants.Reviewer2_Col And Col <> Constants.Reviewer3_Col) Then
        MsgBox "请选择审稿人栏", vbExclamation
    Else
        Set TitleCell = Sheet.Cells(Row, Constants.Title_Col)
        Title = TitleCell.Value2
            
        If Title <> "" Then
            Set ReviewerCell = Sheet.Cells(Row, Col)
            Reviewer = ReviewerCell.Value2
            If Reviewer = "" Then
                MsgBox "审稿人栏为空，请先设置审稿人", vbExclamation
            Else
                Email = Get_Email_For_Reviewer(Reviewer)
                If (Email = "") Then
                    MsgBox "审稿人" & Reviewer & "没有设置邮箱地址", vbExclamation
                Else
                    Send_Email Email, "《声学技术》审稿通知：" & ArticleNo & " " & Title, _
                    "<html>" & _
                    "您好，<p><br/><p>" & _
                    "兹有《声学技术》来稿一份：" & Title & "，稿件编号：" & ArticleNo & "。<p>" & _
                    "呈送烦请代为审查，请于" & _
                    expectDateStr & _
                    "前将审稿意见快递（到付）或Email：sxjs@21cn.com 发回编辑部。<p><br/><p>" & _
                    "祝好!<p>" & _
                    currDateStr & _
                    "</html>", CC
                End If
            End If
        End If
    
    End If
End Sub

'自校通知
Public Sub OnSendSelfReviewEmail()
    
    Dim Row As Integer
    Dim Email As String, CC As String
    Dim ArticleNo, Title As String
    
    Set cell = Worksheets.Application.ActiveCell
    Row = cell.Row
    Set Sheet = cell.Worksheet
    
    Set NoCell = Sheet.Cells(Row, Constants.NO_COL)
    ArticleNo = NoCell.Value2
    
    Set TitleCell = Sheet.Cells(Row, Constants.Title_Col)
    Title = TitleCell.Value2
    
    Set EmailCell = Sheet.Cells(Row, EMAIL_COL)
    Email = EmailCell.Value2
    
    CC = FindCCByName("自校抄送")
    
    If Title <> "" Then
        If Email = "" Then
            MsgBox "没有电子邮箱"
        Else
            Send_Email Email, "《声学技术》自校通知：" & ArticleNo & " " & Title, _
            "<html>" & _
            "您好，<p><br/><p>" & _
            "您的文章马上就要发表了，附件中为校对通知和经过编辑加工的排版稿，请仔细校对文章内容并回答编辑所提出的问题，<p>" & _
            "用修订方式修改稿件，并保留修改痕迹，请尽快将校对稿通过电子邮件返回！<p><br/><p>" & _
            "由于邮寄稿费需要身份证号码，请您一并提供。<p>" & _
            "如联系地址、电话有变化，请及时告知。<p>" & _
            "</html>", CC
        End If
    End If
    
End Sub

'收稿通知
Public Sub OnSendAcceptEmail()
    
    Dim Row As Integer
    Dim Email As String, CC As String
    Dim ArticleNo, Title As String
    
    Set cell = Worksheets.Application.ActiveCell
    Row = cell.Row
    Set Sheet = cell.Worksheet
    
    Set NoCell = Sheet.Cells(Row, Constants.NO_COL)
    ArticleNo = NoCell.Value2
    
    Set TitleCell = Sheet.Cells(Row, Constants.Title_Col)
    Title = TitleCell.Value2
    
    Set EmailCell = Sheet.Cells(Row, EMAIL_COL)
    Email = EmailCell.Value2
    
    CC = FindCCByName("收稿抄送")
    
    If Title <> "" Then
        If Email = "" Then
            MsgBox "没有电子邮箱"
        Else
            Send_Email Email, "《声学技术》收稿通知：" & ArticleNo & " " & Title, _
            "<html>" & _
            "您好，<p><br/><p>" & _
            "您的文章稿件编号是：" & ArticleNo & _
            "。请记住此编号，以备稿件查询时使用。附件是投稿模板和转让协议，请按照模板重新排版，然后E-mail给编辑部。<p>" & _
            "请注意:<p>" & _
            "1)提供完整的首页脚注信息。<p>" & _
            "例：收稿日期：2016-06-05；修回日期：2016-07-20；<p>" & _
            "基金项目：（如无，此行删除）<p>" & _
            "作者简介：第一作者(1992-),男，浙江台州人，硕士，副教授，研究方向为…。<p>" & _
            "通讯作者：姓名，E-mail:<p>" & _
            "2)提供联系人的通讯地址，邮编，固定电话，手机，身份证号码（付稿费时使用）。<p>" & _
            "3)提供作者签名并单位盖章的保密证明和转让协议。<p>" & _
            "4)请在邮件中说明文章的创新点，以便审稿专家掌握了解。<p>" & _
            "</html>", CC
        End If
    End If
    
End Sub


'退修通知
Public Sub OnSendModifyEmail()
    
    Dim Row As Integer
    Dim Email As String
    Dim Title As String
    Dim CC As String
    
    
    Set cell = Worksheets.Application.ActiveCell
    Row = cell.Row
    Set Sheet = cell.Worksheet
    
    Set NoCell = Sheet.Cells(Row, Constants.NO_COL)
    ArticleNo = NoCell.Value2
    
    Set TitleCell = Sheet.Cells(Row, Constants.Title_Col)
    Title = TitleCell.Value2
    
    Set EmailCell = Sheet.Cells(Row, EMAIL_COL)
    Email = EmailCell.Value2
    
    CC = FindCCByName("退修抄送")
    
    If Title <> "" Then
        If Email = "" Then
            MsgBox "没有电子邮箱"
        Else
            Send_Email Email, "《声学技术》退修通知：" & ArticleNo & " " & Title, _
            "<html><body>" & _
            "您好，<p><br/><p>" & _
            "审稿意见发给您，请仔细阅读审稿意见，并逐条回答审稿人所提出的问题，请用修订方式修改稿件，保留修改痕迹以便审稿人了解您的修改之处。<p>" & _
            "<body></html>", CC
        End If
    End If
    
End Sub


Sub Send_Email(Address As String, subject As String, body As String, Optional CC As String)
    Dim Shell As Object
    Dim content As String
    Dim url As Variant
    
    Set Shell = CreateObject("Shell.Application")
    
    content = "mailto://" & Address & "?"
    If CC <> "" Then
        content = content & "cc=" & CC + "&"
    End If
    
    content = content & "subject=" & subject & " &&body=" & body
    
    Shell.Open (content)
    
End Sub


Function FindCCByName(DefName As String) As String
    Dim CCName As Name
    
    For I = 1 To Worksheets.Application.Names.Count
        Set CCName = Worksheets.Application.Names(I)
        If Trim(CCName.Name) = Trim(DefName) Then
            FindCCByName = CCName.RefersToRange.Value2
            Exit For
        End If
    Next I
    
End Function

