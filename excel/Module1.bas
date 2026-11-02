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

Private Sub RestoreProtocolFormulas(ByVal data As Range)
    Dim template As Range, column As Variant, destination As Range
    Set template = ThisWorkbook.Names("GTO_ProtocolTemplate").RefersToRange
    If template.Rows.Count <> data.Rows.Count Then
        Err.Raise vbObjectError + 1000, "GTO protocol", "The protocol template has a different row count."
    End If
    ' Formula strings are kept separately from the editable printout. Restore
    ' only generated columns; school, bib number, level and rank stay editable.
    For Each column In Array(1, 2, 3, 4, 6, 8, 12, 13, 14)
        Set destination = data.Worksheet.Cells(data.Row, column).Resize(data.Rows.Count, 1)
        destination.NumberFormat = "General"
        destination.Formula = template.Columns(column).Value2
    Next column
End Sub

Public Sub AdjustProtocol(Optional ByVal restoreData As Boolean = False)
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
    If restoreData Then RestoreProtocolFormulas data
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
