Attribute VB_Name = "ThisWorkbook"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_Open()
    AdjustProtocol
End Sub

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    If IsProtocolSheet(Sh) Then AdjustProtocol
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    ' Changes to the selector, results, or thresholds must refresh the layout.
    AdjustProtocol
End Sub

Private Sub Workbook_BeforePrint(Cancel As Boolean)
    AdjustProtocol
End Sub
