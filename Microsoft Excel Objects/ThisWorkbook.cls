VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_AddinInstall()
    MsgBox "Installing 声学技术 Add-ins..."
    MenuModule.AddToCellMenu
    MsgBox "Installed 声学技术 Add-ins."
End Sub


Private Sub Workbook_AddinUninstall()
    MenuModule.DeleteFromCellMenu
    MsgBox "Uninstalled 声学技术 Add-ins."
End Sub

