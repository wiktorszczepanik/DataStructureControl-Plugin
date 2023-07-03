Attribute VB_Name = "Clear_Errors"
Option Explicit

Sub clearErrors()
    '
    'clear added errors during checking
    '
    
    Dim Suppress As New Main_DSC
    Dim Restore As New Main_DSC
    
    Const errorIntro = "DSC - hint"
    Dim workBookPath As String
    Dim workbookAction As Workbook
    Dim worksheetsCount As Integer
    Dim counterSh As Integer
    Dim cell As Object
    Dim wsRange As Range
    Dim cellColor As Long
    Dim cellNote As String
    Dim noteChcek As Boolean: noteChcek = False
    Dim cellNoteStrip As String
    Dim noteSplitByConst() As String
    
    Set Suppress = New Main_DSC
    Suppress.ApplicationSuppress
    
    Let workBookPath = Excel.Application.ActiveWorkbook.FullName
    Set workbookAction = Workbooks.Open(workBookPath)
    
    Let worksheetsCount = workbookAction.Worksheets.Count
    For counterSh = 1 To worksheetsCount
        workbookAction.Worksheets(Sheets(counterSh).Name).Activate
        Set wsRange = ActiveSheet.usedRange
        For Each cell In wsRange
            noteChcek = False
            cellNote = cell.NoteText
            cellColor = cell.Interior.Color
            
            'delete error RGB color
            If cellColor = RGB(255, 146, 145) Then
                cell.Interior.Color = RGB(255, 255, 255)
                cell.BorderAround LineStyle:=xlContinuous, _
                Weight:=xlThin, ColorIndex:=15, Color:=0
                cell.Font.Color = vbBlack
            End If
            
            'delete/cut error notes
            If cellNote <> "" Then noteChcek = True
            If noteChcek = True Then
                cellNoteStrip = Left(cell.NoteText, 10)
                If cellNoteStrip = errorIntro Then
                    cell.Comment.Delete
                ElseIf InStr(cellNote, errorIntro) Then
                    noteSplitByConst = Split(cellNote, errorIntro)
                    cell.NoteText Trim(noteSplitByConst(0))
                End If
            End If
        Next cell
    Next counterSh
    
    Set Restore = New Main_DSC
    Restore.ApplicationRestore
    
End Sub
