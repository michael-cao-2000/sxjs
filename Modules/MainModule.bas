Attribute VB_Name = "MainModule"

Public Function FindWorksheetByName(Name As String) As Worksheet
    Dim I As Integer

    For I = 1 To Workbooks.Application.Worksheets.Count
        Set Sheet = Workbooks.Application.Worksheets(I)
        If (Trim(Sheet.Name) = Trim(Name)) Then
            Set FindWorksheetByName = Sheet
            Exit Function
        End If
    Next I

End Function


Public Function FileLocked(strFileName As String) As Boolean
   On Error Resume Next
   ' If the file is already opened by another process,
   ' and the specified type of access is not allowed,
   ' the Open operation fails and an error occurs.
   Open strFileName For Binary Access Read Write Lock Read Write As #1
   Close #1
   ' If an error occurs, the document is currently open.
   If Err.Number <> 0 Then
      ' Display the error number and description.
      MsgBox "文件" & strFileName & "已经被打开" & Str(Err.Number) & " - " & Err.Description
      FileLocked = True
      Err.Clear
   End If
End Function
