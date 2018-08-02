Attribute VB_Name = "MenuModule"
Public Sub AddSubMenu()
    Dim oMainMenuBar  As CommandBar
    Dim oNewMenu  As CommandBarPopup
    
    Set oMainMenuBar = Application.CommandBars.Item("Worksheet Menu Bar")
    
    Set oNewMenu = oMainMenuBar.Controls.Add(Type:=msoControlPopup)
    oNewMenu.Caption = "��ѧ����[&X]"
    oNewMenu.Tag = "sxjs"
    
    AddMenuItem oNewMenu, "��ѷ��ű�", "onArticalPaymentTable"
    AddMenuItem oNewMenu, "���ѷ��ű�", "onReviewFeeTable"
          
End Sub


Public Sub RemoveSubMenu()
    Dim MainMenuBar As CommandBar
    Dim SubMenu  As CommandBarControl
    Dim I As Integer
    
    Set MainMenuBar = Application.CommandBars.Item("Worksheet Menu Bar")
    
    For I = MainMenuBar.Controls.Count To 1 Step -1
        Set SubMenu = MainMenuBar.Controls.Item(I)
        If SubMenu.Tag = "sxjs" Then
            SubMenu.Delete
        End If
    Next
End Sub


Private Sub AddMenuItem(oNewMenu As Object, Name As String, Action As String, Optional bnlNewGroup As Boolean = False)
    Dim oSubMenu As CommandBarControl
    
    Set oSubMenu = oNewMenu.Controls.Add(Type:=msoBarTypeMenuBar)
    With oSubMenu
        .Caption = Name
        .BeginGroup = bnlNewGroup
        If Action <> "" Then
            .OnAction = Action
            .Tag = Action
        End If
    End With
    Set oSubMenu = Nothing
End Sub


Public Sub AddToCellMenu()
    Dim ContextMenu As CommandBar
    Dim MySubMenu As CommandBarControl
    Dim UserProfile As String
    
    UserProfile = Environ("UserProfile") & "\Documents"
    Set fs = CreateObject("Scripting.FileSystemObject")
    NeedReport = fs.FileExists(UserProfile + "\���.�˵�")

    'Delete the controls first to avoid duplicates
    Call DeleteFromCellMenu

    'Set ContextMenu to the Cell menu
    Set ContextMenu = Application.CommandBars("Cell")
    
    'Add custom menu with three buttons
    Set MySubMenu = ContextMenu.Controls.Add(Type:=msoControlPopup, before:=1)
    

    With MySubMenu
        .Caption = "��ѧ����"
        .Tag = "SXJS_Cell_Control_Tag"

        With .Controls.Add(Type:=msoControlButton)
            .OnAction = "'" & ThisWorkbook.Name & "'!" & "OnSendReviewEmail"
            .Caption = "��������ʼ�"
        End With
        With .Controls.Add(Type:=msoControlButton)
            .BeginGroup = True
            .OnAction = "'" & ThisWorkbook.Name & "'!" & "OnSendAcceptEmail"
            .Caption = "�����ո��ʼ�"
        End With
        With .Controls.Add(Type:=msoControlButton)
            .OnAction = "'" & ThisWorkbook.Name & "'!" & "OnSendModifyEmail"
            .Caption = "���������ʼ�"
        End With
        With .Controls.Add(Type:=msoControlButton)
            .OnAction = "'" & ThisWorkbook.Name & "'!" & "OnSendSelfReviewEmail"
            .Caption = "������У�ʼ�"
        End With
        
        If NeedReport Then
            With .Controls.Add(Type:=msoControlButton)
                .BeginGroup = True
                .OnAction = "'" & ThisWorkbook.Name & "'!" & "OnGeneateArticlePaymentTable"
                .Caption = "��ѷ��ű�"
            End With
            With .Controls.Add(Type:=msoControlButton)
                .OnAction = "'" & ThisWorkbook.Name & "'!" & "OnGeneateRemittanceAuthorTable"
                .Caption = "���ڻ��-���"
            End With
            With .Controls.Add(Type:=msoControlButton)
                .OnAction = "'" & ThisWorkbook.Name & "'!" & "OnGeneateServiceFeeAuthorTable"
                .Caption = "����Ʊ�����-���"
            End With
            With .Controls.Add(Type:=msoControlButton)
                .BeginGroup = True
                .OnAction = "'" & ThisWorkbook.Name & "'!" & "OnGeneateReviewerFeeTable"
                .Caption = "���ѷ��ű�"
            End With
            With .Controls.Add(Type:=msoControlButton)
                .OnAction = "'" & ThisWorkbook.Name & "'!" & "OnGeneateRemittanceReviewerTable"
                .Caption = "���ڻ��-����"
            End With
            With .Controls.Add(Type:=msoControlButton)
                .OnAction = "'" & ThisWorkbook.Name & "'!" & "OnGeneateServiceFeeReviewerTable"
                .Caption = "����Ʊ�����-����"
            End With
        End If
    End With

    'Add seperator to the Cell menu
    ContextMenu.Controls(4).BeginGroup = True

    
End Sub


Public Sub DeleteFromCellMenu()
    Dim ContextMenu As CommandBar
    Dim ctrl As CommandBarControl

    'Set ContextMenu to the Cell menu
    Set ContextMenu = Application.CommandBars("Cell")

    'Delete custom controls with the Tag : SXJS_Cell_Control_Tag
    For Each ctrl In ContextMenu.Controls
        Debug.Print ctrl.Tag & ": " & ctrl.Caption
        
        If ctrl.Tag = "" And ctrl.Caption = "" Then
            ctrl.Delete
        End If
        
        If ctrl.Tag = "SXJS_Cell_Control_Tag" Then
            ctrl.Delete
        End If
        
        If ctrl.Tag = "My_Cell_Control_Tag" Then
            ctrl.Delete
        End If
    Next ctrl

    'Delete built-in Save button
    On Error Resume Next
    ContextMenu.FindControl(ID:=3).Delete
    On Error GoTo 0
End Sub

