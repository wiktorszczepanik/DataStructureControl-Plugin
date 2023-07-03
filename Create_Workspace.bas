Attribute VB_Name = "Create_Workspace"
Option Explicit

Public Sub CreateWorkSpace()
    '
    'Create workspace based on folders
    '
    Dim Suppress As New Main_DSC
    Dim Restore As New Main_DSC
    
    Dim newConfigWorkbook As Workbook
    Dim templateConfigSheet As Worksheet
    Dim selectedFilePath As String
    Dim workSpaceMain As String
    Dim confirmedPath As String
    Dim unConfirmedPath As String
    Dim deletedPath As String
    Dim templatePath As String
    Dim wsFolderName As String
    Dim isInputCorrect As Boolean
    Dim extSub As Boolean
    
    Set Suppress = New Main_DSC
    Suppress.ApplicationSuppress
    
    On Error GoTo ErrorHandler
    extSub = False
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select folder where you want to create data structure control workspace"
        .ButtonName = "Select Folder"
        If .Show = 0 Then Exit Sub
        selectedFilePath = .SelectedItems(1) & "\"
    End With
    
    'Main folder (workspace)
    Let isInputCorrect = False
    Do While isInputCorrect = False
        Call MfInputBox(wsFolderName, isInputCorrect, extSub)
        If extSub = True Then Exit Sub
    Loop
    
    Let workSpaceMain = selectedFilePath & wsFolderName
    MkDir (workSpaceMain)
    
    'Folder with confirmed files
    Let confirmedPath = workSpaceMain & "\Correct_Files"
    MkDir (confirmedPath)
    
    'Folder with deleted files (bin)
    Let deletedPath = workSpaceMain & "\Deleted_Files"
    MkDir (deletedPath)
    
    'Folder for templates files
    Let templatePath = workSpaceMain & "\Template_Files"
    MkDir (templatePath)
    
    'Folder with UNconfirmed (Input_Files) files
    Let unConfirmedPath = workSpaceMain & "\Input_Files"
    MkDir (unConfirmedPath)
    
    MsgBox Prompt:="Workspace created correctly!" & vbNewLine & vbNewLine _
    & "Path: " & workSpaceMain, _
    Buttons:=vbOKOnly, Title:="Workspace"
    
    Set Restore = New Main_DSC
    Restore.ApplicationRestore
    
ErrorHandler:
    Select Case Err.Number
        Case 75
            MsgBox Prompt:="Name of the workspace may already exist in selected folder. Try other name.", _
            Buttons:=vbOKOnly, Title:="Error"
    End Select
    
End Sub

Private Sub MfInputBox(folderName, isCorrect, extSub)
    '
    'Sub is handlig inputbox errors
    '
    
    Let folderName = InputBox(Prompt:="Set the name of Workspace folder:", _
    Title:="Workspace name", Default:="workspace_folder")

    If StrPtr(folderName) = 0 Then
        extSub = True
        Exit Sub
    ElseIf folderName = "" Then
        MsgBox Prompt:="Wrong input!"
        isCorrect = False
    Else
        isCorrect = True
    End If
    
End Sub
