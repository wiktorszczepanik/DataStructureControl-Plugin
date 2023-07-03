Attribute VB_Name = "Add_Hard_Value"
Option Explicit

Sub AddHardValue()
    '
    'Creates stable value for xlcg file
    '
    
    Dim Suppress As New Main_DSC
    Dim Restore As New Main_DSC
    
    Dim hardValueStr As String
    Dim cellNote As String
    Dim cellNoteStrip As String
    Dim cell As Object
    
    Set Suppress = New Main_DSC
    Suppress.ApplicationSuppress
    
    'Main
    Let hardValueStr = "HXLCGHVH"
    For Each cell In Selection
        cellNote = cell.NoteText
        cellNoteStrip = Right(cell.NoteText, 6)
        
        If cellNote = "" Then
            cell.NoteText hardValueStr
        ElseIf cellNoteStrip <> hardValueStr Then
            cell.NoteText cellNote & vbNewLine & hardValueStr
        End If
        
        cell.BorderAround LineStyle:=xlContinuous, _
        Weight:=2, ColorIndex:=3
        
    Next cell
    
    Set Restore = New Main_DSC
    Restore.ApplicationRestore
    
End Sub
