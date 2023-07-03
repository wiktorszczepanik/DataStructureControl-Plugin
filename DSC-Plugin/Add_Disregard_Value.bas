Attribute VB_Name = "Add_Disregard_Value"
Option Explicit

Sub AddDisregardValue()
    '
    'Creates disregard value for xlcg file
    '
    
    Dim Suppress As New Main_DSC
    Dim Restore As New Main_DSC
    
    Dim disregardValueStr As String
    Dim cellNote As String
    Dim cellNoteStrip As String
    Dim cell As Object
    
    Set Suppress = New Main_DSC
    Suppress.ApplicationSuppress
    
    'Main
    Let disregardValueStr = "DXLCGDVD"
    For Each cell In Selection
        cellNote = cell.NoteText
        cellNoteStrip = Right(cell.NoteText, 6)
        
        If cellNote = "" Then
            cell.NoteText disregardValueStr
        ElseIf cellNoteStrip <> disregardValueStr Then
            cell.NoteText cellNote & vbNewLine & disregardValueStr
        End If
        
        cell.BorderAround LineStyle:=xlContinuous, _
        Weight:=2, ColorIndex:=5
    
    Next cell
    
    Set Restore = New Main_DSC
    Restore.ApplicationRestore
    
End Sub
