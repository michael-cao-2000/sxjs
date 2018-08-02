Attribute VB_Name = "ReportModule"
Public Sub OnGeneateArticlePaymentTable()

    ArticlePaymentForm.Show vbModal
    
End Sub

Public Sub OnGeneateReviewerFeeTable()

    ReviewerFeeForm.Show vbModal
    
End Sub

'���ڻ��-����
Public Sub OnGeneateRemittanceReviewerTable()

    Dim I, Num, EmptyCount As Integer
    Dim Title, Name As String
    Dim Reviewers(1000) As ReviewPayment
    Dim currDate As Date
    Dim currDateStr As String
    Dim sumFee As Double, postageSum As Double
    Dim fullPathName As String
    Dim fileTypeName As String
        
    fileTypeName = "���ڻ��-����"
    
    currDate = Date
    currDateStr = Format(currDate, "yyyy-mm-dd")
    
    Num = GetReviewers(Reviewers)
    If Num = 0 Then
        Exit Sub
    End If
    
    Set ReviewerDict = GetReviewerDict()
    
    '�����ܽ��ӻ��
    For I = 1 To Num - 1
        Name = Reviewers(I).Name
        If (Reviewers(I).Postage <> 0) Then
            sumFee = sumFee + Reviewers(I).Fee
            sumPostage = sumPostage + Reviewers(I).Postage
        End If
    Next I

    UserProfile = Environ("UserProfile") & "\Documents"
    Set fs = CreateObject("Scripting.FileSystemObject")
    fullPathName = UserProfile & "\���ڻ��-����(" & currDateStr & ").csv"
    On Error Resume Next
    Set file = fs.CreateTextFile(fullPathName, True)
    If Err.Number > 0 Then
        MsgBox "���ڻ��-�����ļ� " & fullPathName & " �Ѿ����򿪣��޷�������ǰ���������ȹر��ļ�����ִ�� (������룺" & Err.Number & ")", vbCritical
        Exit Sub
    End If
    file.writeLine "�̻�����,�ļ�����,�ܱ���,�ܽ��"
    file.writeLine Chr$(9) & "310000000,0,0," & (sumFee + sumPostage)
    file.writeLine "�����,�տ����ʱ�,�տ�������,�տ��˵�ַ,����"

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

'���ڻ��-���
Public Sub OnGeneateRemittanceAuthorTable()

    Dim Title As String
    Dim currDate As Date
    Dim currDateStr As String
    Dim fullPathName As String
    Dim dblPaySum As Double, dblPostageSum As Double
    Dim fileTypeName As String
    
    fileTypeName = "���ڻ��-���"
    
    currDate = Date
    currDateStr = Format(currDate, "yyyy-mm-dd")
    
    Set Sheet = FindWorksheetByName("��ѷ��ű�")
    If Sheet Is Nothing Then
        MsgBox "û���ҵ�����ѷ��ű����������ɡ���ѷ��ű�", vbExclamation
        Exit Sub
    End If
    
    '�������ܽ�� +���ʷ�
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
    fullPathName = UserProfile & "\���ڻ��-���(" & currDateStr & ").csv"
        
    On Error Resume Next
    Set file = fs.CreateTextFile(fullPathName, True)
    If Err.Number > 0 Then
        MsgBox "���ڻ��-����ļ� " & fullPathName & " �Ѿ����򿪣��޷�������ǰ���������ȹر��ļ�����ִ�� (������룺" & Err.Number & ")", vbCritical
        Exit Sub
    End If
    
    file.writeLine "�̻�����,�ļ�����,�ܱ���,�ܽ��"
    file.writeLine Chr$(9) & "310000000,0,0," & (dblPaySum + dblPostageSum)
    file.writeLine "�����,�տ����ʱ�,�տ�������,�տ��˵�ַ,����"
    
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

'����Ʊ�����-����
Public Sub OnGeneateServiceFeeReviewerTable()
    Dim I, Num, EmptyCount As Integer
    Dim Title, Name As String
    Dim currDate As Date
    Dim currDateStr As String
    Dim Reviewers(1000) As ReviewPayment
    Dim fileTypeName As String
    Dim fullPathName As String
    
    fileTypeName = "����Ʊ�����-����"

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
        MsgBox "���ڻ��-�����ļ� " & fullPathName & " �Ѿ����򿪣��޷�������ǰ���������ȹر��ļ�����ִ�� (������룺" & Err.Number & ")", vbCritical
        Exit Sub
    End If
    
    Set file = fs.CreateTextFile(UserProfile & "\����Ʊ�����-����(" & currDateStr & ").csv", True)
    file.writeLine "�й���ѧԺ��ѧ�о��������о�վ����Ʊ�����"
    file.writeLine "���,����,֤������,֤������,��������,�����ڼ�,��Ԫ��"

    For I = 1 To Num - 1
        Name = Reviewers(I).Name
        file.Write I & ","
        file.Write Name & ","
        file.Write "���֤" & ","
        file.Write Chr$(9) & ReviewerDict.Item(Name).ID & ","
        file.writeLine ",,"
    Next I
    
    file.Close
    
    PromptSuccess fileTypeName, fullPathName
    
End Sub

'����Ʊ�����-���
Public Sub OnGeneateServiceFeeAuthorTable()
    Dim Title As String
    Dim currDate As Date
    Dim currDateStr As String
    Dim fileTypeName As String
    Dim fullPathName As String
    
    fileTypeName = "����Ʊ�����-���"
    
    currDate = Date
    currDateStr = Format(currDate, "yyyy-mm-dd")
    
    Set Sheet = FindWorksheetByName("��ѷ��ű�")
    If Sheet Is Nothing Then
        MsgBox "û���ҵ�����ѷ��ű����������ɡ���ѷ��ű�", vbExclamation
        Exit Sub
    End If
    
    UserProfile = Environ("UserProfile") & "\Documents"
    Set fs = CreateObject("Scripting.FileSystemObject")
    fullPathName = UserProfile & "\����Ʊ�����-���(" & currDateStr & ").csv"
    On Error Resume Next
    Set file = fs.CreateTextFile(fullPathName, True)
    If Err.Number > 0 Then
        MsgBox "����Ʊ�����-����ļ� " & fullPathName & " �Ѿ����򿪣��޷�������ǰ���������ȹر��ļ�����ִ�� (������룺" & Err.Number & ")", vbCritical
        Exit Sub
    End If
    
    file.writeLine "�й���ѧԺ��ѧ�о��������о�վ����Ʊ�����"
    file.writeLine "���,����,֤������,֤������,��������,�����ڼ�,��Ԫ��"
    
    For I = 1 To 1000
        Title = Trim(Sheet.Cells(I + 1, 3).Value2)
        If (Title = "") Then
            Exit For
        End If
        file.Write I & ","
        file.Write Sheet.Cells(I + 1, 1).Value2 & ","
        file.Write "���֤,"
        file.Write Sheet.Cells(I + 1, 7).Value2 & ","
        file.writeLine ",,"
    Next I
    
    file.Close
    
    PromptSuccess fileTypeName, fullPathName

End Sub


Sub PromptSuccess(ByRef fileTypeName As String, ByRef fullPathName As String)
    If (vbYes = MsgBox("������" & fileTypeName & "�ļ�" & vbCrLf & fullPathName & vbCrLf & "��Ҫ���ڴ���", vbQuestion & vbYesNo)) Then
        Workbooks.Open fullPathName
    End If
End Sub

