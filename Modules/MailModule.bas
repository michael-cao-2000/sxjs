Attribute VB_Name = "MailModule"
'���֪ͨ
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
    
    CC = FindCCByName("��峭��")
    
    If (Col <> Constants.Reviewer1_Col And Col <> Constants.Reviewer2_Col And Col <> Constants.Reviewer3_Col) Then
        MsgBox "��ѡ���������", vbExclamation
    Else
        Set TitleCell = Sheet.Cells(Row, Constants.Title_Col)
        Title = TitleCell.Value2
            
        If Title <> "" Then
            Set ReviewerCell = Sheet.Cells(Row, Col)
            Reviewer = ReviewerCell.Value2
            If Reviewer = "" Then
                MsgBox "�������Ϊ�գ��������������", vbExclamation
            Else
                Email = Get_Email_For_Reviewer(Reviewer)
                If (Email = "") Then
                    MsgBox "�����" & Reviewer & "û�����������ַ", vbExclamation
                Else
                    Send_Email Email, "����ѧ���������֪ͨ��" & ArticleNo & " " & Title, _
                    "<html>" & _
                    "���ã�<p><br/><p>" & _
                    "���С���ѧ����������һ�ݣ�" & Title & "�������ţ�" & ArticleNo & "��<p>" & _
                    "���ͷ����Ϊ��飬����" & _
                    expectDateStr & _
                    "ǰ����������ݣ���������Email��sxjs@21cn.com ���ر༭����<p><br/><p>" & _
                    "ף��!<p>" & _
                    currDateStr & _
                    "</html>", CC
                End If
            End If
        End If
    
    End If
End Sub

'��У֪ͨ
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
    
    CC = FindCCByName("��У����")
    
    If Title <> "" Then
        If Email = "" Then
            MsgBox "û�е�������"
        Else
            Send_Email Email, "����ѧ��������У֪ͨ��" & ArticleNo & " " & Title, _
            "<html>" & _
            "���ã�<p><br/><p>" & _
            "�����������Ͼ�Ҫ�����ˣ�������ΪУ��֪ͨ�;����༭�ӹ����Ű�壬����ϸУ���������ݲ��ش�༭����������⣬<p>" & _
            "���޶���ʽ�޸ĸ�����������޸ĺۼ����뾡�콫У�Ը�ͨ�������ʼ����أ�<p><br/><p>" & _
            "�����ʼĸ����Ҫ���֤���룬����һ���ṩ��<p>" & _
            "����ϵ��ַ���绰�б仯���뼰ʱ��֪��<p>" & _
            "</html>", CC
        End If
    End If
    
End Sub

'�ո�֪ͨ
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
    
    CC = FindCCByName("�ո峭��")
    
    If Title <> "" Then
        If Email = "" Then
            MsgBox "û�е�������"
        Else
            Send_Email Email, "����ѧ�������ո�֪ͨ��" & ArticleNo & " " & Title, _
            "<html>" & _
            "���ã�<p><br/><p>" & _
            "�������¸������ǣ�" & ArticleNo & _
            "�����ס�˱�ţ��Ա������ѯʱʹ�á�������Ͷ��ģ���ת��Э�飬�밴��ģ�������Ű棬Ȼ��E-mail���༭����<p>" & _
            "��ע��:<p>" & _
            "1)�ṩ��������ҳ��ע��Ϣ��<p>" & _
            "�����ո����ڣ�2016-06-05���޻����ڣ�2016-07-20��<p>" & _
            "������Ŀ�������ޣ�����ɾ����<p>" & _
            "���߼�飺��һ����(1992-),�У��㽭̨���ˣ�˶ʿ�������ڣ��о�����Ϊ����<p>" & _
            "ͨѶ���ߣ�������E-mail:<p>" & _
            "2)�ṩ��ϵ�˵�ͨѶ��ַ���ʱ࣬�̶��绰���ֻ������֤���루�����ʱʹ�ã���<p>" & _
            "3)�ṩ����ǩ������λ���µı���֤����ת��Э�顣<p>" & _
            "4)�����ʼ���˵�����µĴ��µ㣬�Ա����ר�������˽⡣<p>" & _
            "</html>", CC
        End If
    End If
    
End Sub


'����֪ͨ
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
    
    CC = FindCCByName("���޳���")
    
    If Title <> "" Then
        If Email = "" Then
            MsgBox "û�е�������"
        Else
            Send_Email Email, "����ѧ����������֪ͨ��" & ArticleNo & " " & Title, _
            "<html><body>" & _
            "���ã�<p><br/><p>" & _
            "������������������ϸ�Ķ����������������ش����������������⣬�����޶���ʽ�޸ĸ���������޸ĺۼ��Ա�������˽������޸�֮����<p>" & _
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

