Attribute VB_Name = "Create_Template_From_File"
Option Explicit

Public Sub ConfigListType()
    '
    'Create list for template/config file
    '
    
    Dim Suppress As New Main_DSC
    Dim Restore As New Main_DSC
    
    Dim selectedFilePath, workBookPath As String
    Dim worksheetsCount As Byte
    Dim templateWorkbook As Workbook
    Dim counterSh As Byte
    Dim templateFile As String
    Dim xlcgFileName As String
    Dim worksheetName As String
    Dim allWorksheetsTemplateCells() As Variant
    Dim countForTemplate As Integer
    Dim stringForTemplate As String
    Dim isATableObj, isNoteObj As Boolean
    Dim notesCheck As Integer
    Dim usedRangeCol, usedRangeRow As Integer
    Dim tableNum As Byte
    
    Set Suppress = New Main_DSC
    Suppress.ApplicationSuppress
    
    'Select file for template
    If MsgBox(Prompt:="Do you want to create template file from current open workbook?", _
    Buttons:=vbYesNo, Title:="Template file") = vbYes Then
        workBookPath = Excel.Application.ActiveWorkbook.FullName
    Else
        With Application.FileDialog(msoFileDialogOpen)
            .Title = "Select Excel File"
            .ButtonName = "Select File"
            .Filters.Add "Excel Files", "*.xlsx?", 1
            .AllowMultiSelect = False
            If .Show = 0 Then Exit Sub
            workBookPath = .SelectedItems(1)
        End With
    End If
    
    'Location for save the template
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select template location"
        .ButtonName = "Select template location"
        If .Show = 0 Then Exit Sub
        selectedFilePath = .SelectedItems(1) & "\"
    End With
    
    'Create empty template file
    Let xlcgFileName = InputBox(Prompt:="Set the name of template file", _
    Title:="Template file name", Default:="Template")
    Let templateFile = selectedFilePath & xlcgFileName & ".xlcg"
    Open templateFile For Output As #1
    Print #1, "*" & Format(Now, "DD/MM/YY") & " " & Format(Now, "HH:MM Am/Pm") & "*"
        
    'Main
    Set templateWorkbook = Workbooks.Open(workBookPath)
    Let worksheetsCount = templateWorkbook.Worksheets.Count
    For counterSh = 1 To worksheetsCount
        templateWorkbook.Worksheets(Sheets(counterSh).Name).Activate
        worksheetName = ActiveSheet.Name
        countForTemplate = ActiveSheet.usedRange.Count
        stringForTemplate = ActiveSheet.usedRange.Address(RowAbsolute:=True, ColumnAbsolute:=False)
        usedRangeCol = ActiveSheet.usedRange.Columns.Count
        usedRangeRow = ActiveSheet.usedRange.Rows.Count
        
        'Checking is table object in worksheet
        Let tableNum = ActiveSheet.ListObjects.Count
        If tableNum > 0 Then
            isATableObj = True
        Else
            isATableObj = False
        End If
        
        'Checking is table countains any notes with HXLCGHVH
        Let isNoteObj = False
        Let notesCheck = 0
        On Error Resume Next
        notesCheck = Range("A1").SpecialCells(xlCellTypeComments).Count
        If notesCheck > 0 Then isNoteObj = True
        
        'Set size of list
        ReDim allWorksheetsTemplateCells(1 To countForTemplate, 1 To 5)
        
        'Base loop over cells and add hard value
        Call BaseLoop(countForTemplate, allWorksheetsTemplateCells, _
        stringForTemplate, isNoteObj)
        
        'Add header loop over cells
        If isATableObj = True Then
            Call AppendHeaderType(allWorksheetsTemplateCells, tableNum)
        End If
        
        'Fill xlcg file by collected data
        Print #1, "#" & counterSh & "/" & worksheetName & "/" & _
        usedRangeCol & "/" & usedRangeRow & "/" & stringForTemplate
        Call FillXlcgFile(allWorksheetsTemplateCells, counterSh, usedRangeCol)

    Next counterSh
    
    ActiveWorkbook.Worksheets(Sheets(1).Name).Activate
    Close #1
    Set Restore = New Main_DSC
    Restore.ApplicationRestore
    
End Sub


Private Sub BaseLoop(usedRange, allWsCells, strForTemplate, notesSh)
    '
    'It is base loop over cells to collect data values
    '
    
    Dim baseCell As Integer
    Dim rangeForTemplate As Range
    Dim tempStr As String
    Dim noteStr As String
    Dim xlcgStr As String
    Const notLiveStr = "{}"
    
    Set rangeForTemplate = ActiveSheet.Range(strForTemplate)
    For baseCell = 1 To usedRange
            allWsCells(baseCell, 1) = rangeForTemplate(baseCell).Address(RowAbsolute:=False, _
            ColumnAbsolute:=False)
            allWsCells(baseCell, 2) = VarType(rangeForTemplate(baseCell))
            allWsCells(baseCell, 3) = "n"
            allWsCells(baseCell, 4) = 0
            allWsCells(baseCell, 5) = ""
            If notesSh = True Then
                noteStr = rangeForTemplate(baseCell).NoteText
                tempStr = Right(noteStr, 8)
                If tempStr = "HXLCGHVH" Then
                    allWsCells(baseCell, 4) = 1
                    xlcgStr = rangeForTemplate(baseCell).value
                    xlcgStr = Replace(xlcgStr, "/", "-")
                    allWsCells(baseCell, 5) = "<" + xlcgStr + "]"
                ElseIf tempStr = "DXLCGDVD" Then
                    allWsCells(baseCell, 4) = 2
                    allWsCells(baseCell, 5) = "<" + notLiveStr + "]"
                End If
            End If
            
    Next baseCell
        
End Sub

Private Sub AppendHeaderType(baseList, tCount)
    '
    'It adds Table types to template list if there exist table
    '
    
    Dim tables As New Main_DSC
    
    Dim tablesList() As String
    Dim headerList() As String
    Dim hRangeList As Range, hCount As Range
    Dim hLoopCounter As Long
    Dim hLenList As Integer, nhCount As Integer
    Dim bhCount As Long
    Dim cellList() As String
    Dim cRangeList As Range, cCount As Range
    Dim cLoopCounter As Long
    Dim cLenList As Integer, ncCount As Integer
    Dim bcCount As Long
    Dim bCount As Long
    ReDim tableList(1 To tCount)
    
    'collecting cell (c) type
    Let cLoopCounter = 0
    Let cLenList = 0
    For ncCount = 1 To tCount
        Set cRangeList = ActiveSheet.ListObjects.item(ncCount).DataBodyRange
        cLenList = cLenList + cRangeList.Count
        ReDim Preserve cellList(1 To cLenList)
        For Each cCount In cRangeList
            cLoopCounter = cLoopCounter + 1
            cellList(cLoopCounter) = cCount.Address(RowAbsolute:=False, _
            ColumnAbsolute:=False)
        Next cCount
    Next ncCount
    
    'collecting header (h) type
    Let hLoopCounter = 0
    Let hLenList = 0
    For nhCount = 1 To tCount
        Set hRangeList = ActiveSheet.ListObjects.item(nhCount).HeaderRowRange
        hLenList = hLenList + hRangeList.Count
        ReDim Preserve headerList(1 To hLenList)
        For Each hCount In hRangeList
            hLoopCounter = hLoopCounter + 1
            headerList(hLoopCounter) = hCount.Address(RowAbsolute:=False, _
            ColumnAbsolute:=False)
        Next hCount
    Next nhCount
    
    'append type (c, h, n) to base list
    For bCount = 1 To UBound(baseList)
        For bcCount = 1 To UBound(cellList) 'c type
            If baseList(bCount, 1) = cellList(bcCount) Then
                baseList(bCount, 3) = "c"
            End If
        Next bcCount
        
        For bhCount = 1 To UBound(headerList) 'h type
            If baseList(bCount, 1) = headerList(bhCount) Then
                baseList(bCount, 3) = "h"
            End If
        Next bhCount
        
    Next bCount
    
End Sub


Private Sub FillXlcgFile(xlcgArray, pageCount, rangeCol)
    '
    'It is adds config to template File
    '
    
    Dim cXlcg As Integer
    Dim rowLen As Integer
    Dim rowXlcgData As String
    Dim loopCount As Long
    
    For cXlcg = 1 To UBound(xlcgArray) Step rangeCol
        rowXlcgData = ""
        loopCount = 0
        For rowLen = 1 To rangeCol
            loopCount = (cXlcg + rowLen) - 1
            rowXlcgData = rowXlcgData & xlcgArray(loopCount, 1) & "/" & _
            xlcgArray(loopCount, 2) & "/" & xlcgArray(loopCount, 3) & "/" & _
            xlcgArray(loopCount, 4) & "/" & xlcgArray(loopCount, 5) & "|"
        Next rowLen
        rowXlcgData = Left(rowXlcgData, Len(rowXlcgData) - 1)
        Print #1, rowXlcgData
    Next cXlcg
    
End Sub
