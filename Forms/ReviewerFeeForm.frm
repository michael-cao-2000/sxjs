VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ReviewerFeeForm 
   Caption         =   "审稿费发放表"
   ClientHeight    =   6550
   ClientLeft      =   50
   ClientTop       =   380
   ClientWidth     =   9690
   OleObjectBlob   =   "ReviewerFeeForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ReviewerFeeForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private ReviewRecords() As ReviewRecord

Private OtherReviewRecords() As ReviewRecord


Private Sub CancelButton_Click()

    Unload Me

End Sub


Private Sub DeleteButton_Click()
    If ReviewerListBox.ListIndex >= 0 Then
        ReviewerListBox.RemoveItem ReviewerListBox.ListIndex
    End If
End Sub

Private Sub OKButton_Click()
    Dim Value As String
    Dim Items() As String
    Dim Name As String
    Dim Order As String
    Dim ArticleNo As String
    Dim Title As String
    Dim Source As String
    Dim Row As Integer
    Dim Col As Integer
    Dim NewValue As String
    
    Set ReviewerSheet = FindWorksheetByName("审稿专家库")
    
    Set Sheet0 = FindWorksheetByName("来稿登记")
    Set Sheet1 = FindWorksheetByName("表外录用来稿登记")
    
    Set Sheet = FindWorksheetByName("审稿费发放表")
    If Sheet Is Nothing Then
        Set Sheet = Workbooks.Application.Worksheets.Add(, ReviewerSheet)
        Sheet.Name = "审稿费发放表"
    End If
    
    Sheet.Cells.ClearContents
    Sheet.Cells.WrapText = True
    
    Sheet.Cells(1, 1).Value2 = "姓名"
    Sheet.Cells(1, 1).ColumnWidth = 10
    Sheet.Cells(1, 2).Value2 = "稿件编号及文章题目"
    Sheet.Cells(1, 2).ColumnWidth = 80
    Sheet.Cells(1, 3).Value2 = "审稿费金额"
    Sheet.Cells(1, 3).ColumnWidth = 10
    Sheet.Cells(1, 4).Value2 = "汇费金额"
    Sheet.Cells(1, 4).ColumnWidth = 10
    Sheet.Cells(1, 5).Value2 = "审稿人签字或邮局回单号码"
    Sheet.Cells(1, 5).ColumnWidth = 10
    
    NewValue = Format(Date, "yyyy-mm-dd") & "已付"
    
    K = 2
    For I = 0 To ReviewerListBox.ListCount - 1
        Value = ReviewerListBox.List(I)
        Debug.Print Value
        Items = Split(Value, ":")
        Name = Items(0)
        OrderNo = Trim(Items(1))
        ArticleNo = Trim(Items(2))
        Title = Trim(Items(3))
        N = UBound(Items)
        Col = CInt(Items(N))
        Row = CInt(Items(N - 1))
        Source = Trim(Items(N - 2))
        
        If OrderNo = "(1)" Then
            Sheet.Cells(K, 1).Value2 = Name
        End If
        
        Sheet.Cells(K, 2).Value2 = Chr$(9) & ArticleNo & " " & Title
        
        If Source = "0" Then
            Sheet0.Cells(Row, Col).Value2 = NewValue
        ElseIf Source = "1" Then
            Sheet1.Cells(Row, Col).Value2 = NewValue
        End If
        
        K = K + 1
    Next I
    
    Unload Me

End Sub

Private Sub ReviewerListBox_Change()
    Dim Count As Integer
        
    For I = 0 To ReviewerListBox.ListCount - 1
        If ReviewerListBox.Selected(I) = True Then
            Count = Count + 1
        End If
    Next I
    
    DeleteButton.Enabled = (Count >= 1)
    
End Sub

Private Sub UserForm_Activate()
    Dim ArticleList As Variant
    Dim Name As String
    Dim Text As String
    
    Set ReviewerDict = GetReviewerDict
    Set Dict = CreateObject("Scripting.Dictionary")
    
    FrameProgress.Visible = True
    FindUnpaidReviewRecords "来稿登记", Dict, ReviewerDict, "0"
    FindUnpaidReviewRecords "表外录用来稿登记", Dict, ReviewerDict, "1"
    FrameProgress.Visible = False
    
    Keys = Dict.Keys
    Values = Dict.Items
    For I = 0 To Dict.Count - 1
        Name = Keys(I)
        ArticleList = Values(I)
        For K = 0 To UBound(ArticleList) - 1
            X = ArticleList(K)
            If X <> "" Then
                ReviewerListBox.AddItem Name & " : (" & (K + 1) & ") : " & X
            End If
        Next
    Next I
    
End Sub


Sub UpdateProgressBar(PctDone As Single)

    With ReviewerFeeForm
        .FrameProgress.Caption = Format(PctDone, "0%")
        .LabelProgress.Width = PctDone * (.FrameProgress.Width - 2)
    End With
    
    DoEvents
End Sub


Function FindUnpaidReviewRecords(ByRef SheetName As String, ByRef Result As Variant, ByRef ReviewerDict As Variant, Source As String)
    Dim ReviewerCols(10) As Integer
    Dim ReviewerCount As Integer
    Dim EmptyCount As Integer
    Dim Reviewer As String
    Dim ReviewDate As String
    Dim Paid As String
    Dim ExistArray() As String

    Set Sheet = FindWorksheetByName(SheetName)

    ReviewerCount = FindReviewersCol(Sheet, ReviewerCols)
    
    For I = 2 To 20000
        Title = Sheet.Cells(I, 2)
        ArticleNo = CStr(Sheet.Cells(I, 1))
        If Title = "" And ArticleNo = "" Then
            EmptyCount = EmptyCount + 1
            If (EmptyCount > 5) Then
                UpdateProgressBar 1
                Exit For
            End If
        Else
            EmptyCount = 0
        End If
        UpdateProgressBar (I / 20000)
        
        For J = 0 To ReviewerCount - 1
            K = ReviewerCols(J)
            Paid = Sheet.Cells(I, K + 2).Value2
            If (Paid = "") Then
                Reviewer = Sheet.Cells(I, K).Value2
                If (Reviewer <> "") Then
                    ReviewDate = Sheet.Cells(I, K + 1).Value2
                    If (ReviewDate <> "") Then
                        If Result.Exists(Reviewer) Then
                            ExistArray = Result.Item(Reviewer)
                            N = UBound(ExistArray)
                            For M = 0 To N - 1
                                If ExistArray(M) = "" Then
                                   Exit For
                                End If
                            Next M
                            If M < 4 Or IsReviewerFromSAL(Reviewer, ReviewerDict) Then
                                If (M >= N) Then
                                    ReDim Preserve ExistArray(N + 5)
                                End If
                                ExistArray(M) = ArticleNo + " : " + Title & ":" & Source & ":" & I & ":" & (K + 2)
                                Result.Item(Reviewer) = ExistArray
                            End If
                        Else
                            Dim ArticleList(4) As String
                            ArticleList(0) = ArticleNo & " : " & Title & ":" & Source & ":" & I & ":" & (K + 2)
                            Result.Item(Reviewer) = ArticleList
                        End If
                    End If
                End If
            End If
        Next J
    Next I

End Function

Private Function FindReviewersCol(ByRef Sheet As Variant, ByRef ReviewerCols() As Integer) As Integer

    J = 0
    For I = 1 To 100
        StrValue = Sheet.Cells(1, I).Value2
        If StrValue <> "" Then
            If InStr(1, StrValue, "审稿人") Then
                ReviewerCols(J) = I
                J = J + 1
            End If
        End If
    Next I
    FindReviewersCol = J
    
End Function

Private Function IsReviewerFromSAL(ByRef Name As String, ByRef ReviewerDict As Variant) As Boolean

    Set Detail = ReviewerDict.Item(Name)
    If (Detail.Company = "东海站") Then
        IsReviewerFromSAL = True
    Else
        IsReviewerFromSAL = False
    End If
 
End Function
