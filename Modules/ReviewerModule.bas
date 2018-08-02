Attribute VB_Name = "ReviewerModule"
'ר�ҿ����Ŀ���
Const NAME_COL = 1
Const ID_Col = 2
Const Address_Col = 3
Const ZipCode_Col = 4
Const Phone_Col = 5
Const EMAIL_COL = 6
Const COMPANY_COL = 7

'�������ר�����Ʒ����������ַ
Public Function Get_Email_For_Reviewer(Reviewer As String)
    Dim Name, Email As String
        
    Set Sheet = Worksheets.Application.Sheets("���ר�ҿ�")
    For I = 1 To 100
        Set NameCell = Sheet.Cells(I, NAME_COL)
        Name = NameCell.Value2
        If Name = Reviewer Then
            Set MailCell = Sheet.Cells(I, EMAIL_COL)
            Get_Email_For_Reviewer = MailCell.Value2
            Exit For
        End If
    Next I

End Function


'�����ר�ҿ��з���ר���ֵ䣬key��ר��������value��ר����ϸ��Ϣ�������֤�ţ���ַ���绰���ʱ��
Public Function GetReviewerDict() As Object
    Dim EmptyCount As Integer
    Dim Name, ID, Address, ZipCode, Phone, Mail, Company As String
    
    Set Dict = CreateObject("Scripting.Dictionary")
    Set GetReviewerDict = Dict
    Set Sheet = FindWorksheetByName("���ר�ҿ�")
    For I = 1 To 1000
        Set cell = Sheet.Cells(I, 1)
        Name = Trim(cell.Value2)
        If (Name = "") Then
            EmptyCount = EmptyCount + 1
            If (EmptyCount > 5) Then
                Exit For
            End If
        Else
            EmptyCount = 0
            ID = Trim(Sheet.Cells(I, 2).Value2)
            Address = Trim(Sheet.Cells(I, 3).Value2)
            ZipCode = Trim(Sheet.Cells(I, 4).Value2)
            Phone = Trim(Sheet.Cells(I, 5).Value2)
            Mail = Trim(Sheet.Cells(I, 6).Value2)
            Company = Trim(Sheet.Cells(I, 7).Value2)
            
            Set Detail = New ReviewerDetail
            With Detail
                .ID = ID
                .Address = Address
                .ZipCode = ZipCode
                .Phone = Phone
                .Mail = Mail
                .Company = Company
            End With
            
            Dict.Add Name, Detail
        End If
    Next I
    
End Function


'�����ѷ��ű�����ȡ������嵥
Public Function GetReviewers(ByRef Reviewers() As ReviewPayment) As Integer
    Dim Name As String
    Dim Fee As Double
    Dim Postage As Double
    
    Set Sheet = FindWorksheetByName("���ѷ��ű�")
    If Sheet Is Nothing Then
        MsgBox "û���ҵ������ѷ��ű����������ɡ����ѷ��ű�", vbExclamation
        GetReviewers = 0
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
                Set payment = New ReviewPayment
                payment.Name = Name
                payment.Fee = Sheet.Cells(N, 3).Value2
                payment.Postage = Sheet.Cells(N, 4).Value2
                Set Reviewers(I) = payment
                I = I + 1
            End If
        End If
    Next N
    
    GetReviewers = I

End Function


