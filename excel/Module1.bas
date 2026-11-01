Attribute VB_Name = "Module1"
Option Explicit
Private updating As Boolean

Public Function IsProtocolSheet(ByVal Sh As Object) As Boolean
    Dim data As Range
    On Error GoTo NotProtocol
    Set data = ThisWorkbook.Names("GTO_ProtocolData").RefersToRange
    IsProtocolSheet = (Sh Is data.Worksheet)
NotProtocol:
End Function

Public Sub AdjustProtocol()
    Dim data As Range, signatures As Range, sheet As Worksheet
    Dim count As Long, hideFrom As Long, oldEvents As Boolean, oldScreen As Boolean
    If updating Then Exit Sub
    oldEvents = Application.EnableEvents
    oldScreen = Application.ScreenUpdating
    On Error GoTo Failed
    updating = True
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Set data = ThisWorkbook.Names("GTO_ProtocolData").RefersToRange
    Set signatures = ThisWorkbook.Names("GTO_ProtocolSignatures").RefersToRange
    Set sheet = data.Worksheet
    ' Recalculate dependent helper sheets before counting numeric row numbers.
    Application.CalculateFull
    count = Application.WorksheetFunction.Count(data.Columns(1))
    sheet.Rows(data.Row & ":" & (signatures.Row - 1)).Hidden = False
    hideFrom = data.Row + count
    If hideFrom < signatures.Row Then
        sheet.Rows(hideFrom & ":" & (signatures.Row - 1)).Hidden = True
    End If
    signatures.EntireRow.Hidden = False
    sheet.PageSetup.PrintArea = sheet.Range("A1", signatures.Cells(1, signatures.Columns.Count)).Address
Done:
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    updating = False
    Exit Sub
Failed:
    MsgBox "The protocol layout could not be refreshed. Error " & Err.Number & ": " & Err.Description, vbExclamation, "GTO protocol"
    Resume Done
End Sub
