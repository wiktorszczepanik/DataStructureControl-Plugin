Attribute VB_Name = "Manual_Check"
Option Explicit

Sub ManualCheck()
    '
    'Compare current open file to selected template
    '
    
    Dim Suppress As New Main_DSC
    Dim Restore As New Main_DSC
    
    Dim workbookNumber As Workbook
    Dim workbookName, selectedXlcgPath As String
    Dim currentLine, firstCharStr As String
    Dim splitLine, splitRegularLine, strAddrTempWs As String
    Dim splitTopLine() As String
    Dim lenRow, lenCol, workSheetNumber As Integer
    Dim cellCounter, rowCounter As Integer
    Dim tempCell As String
    Dim rowArray() As Variant
    Dim rowRange As Range
    Dim startAddress, endAddress As String
    Dim cellRange As Integer
    Dim cellInfoArray() As Variant
    Dim addToArrStatus, isATableObj, isArrEmpty As Boolean
    Dim addStrToArray, hardValueData As String
    Dim exitArray() As Variant
    Dim sizeOfExitArr As Integer
    Dim tableHeaderArr() As String
    Dim hTempRangeStr, cTempRangeStr As String
    Dim headerRangeArr(), cellRangeArr() As String
    Dim tableCounter, errArrLen, tableNum As Byte
    Dim eCounter As Long
    Dim errorArr As Variant
    Dim outsiteSize As Boolean
    Dim anyOutsiteErr, howManyErrors As Long: anyOutsiteErr = 0
    Dim wsCount, osCounter As Integer
    Dim summaryStr As String
    Dim usedRangeWs As Range
    
    Set Suppress = New Main_DSC
    Suppress.ApplicationSuppress
    
    'Select template file path
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Select template file"
        .ButtonName = "Select template file"
        If .Show = 0 Then Exit Sub
        selectedXlcgPath = .SelectedItems(1)
    End With
    
    Let workbookName = Excel.Application.ActiveWorkbook.FullName
    Set workbookNumber = Workbooks.Open(workbookName)
    
    'Error cells outside of template size
    If MsgBox(Prompt:="Do you want to error cells outside of the template size?", _
    Buttons:=vbYesNo, Title:="Type of process - Template file") = vbYes Then
        Let outsiteSize = True
    Else
        Let outsiteSize = False
    End If
    
    Open selectedXlcgPath For Input As #1
    
    'Main
    Let wsCount = 0
    Do Until EOF(1)
        Input #1, currentLine
        firstCharStr = Left(currentLine, 1)
        rowCounter = 0
        If firstCharStr = "*" Then
        ElseIf firstCharStr = "#" Then
            'Current worksheet info and basic actions
            wsCount = wsCount + 1
            workSheetNumber = Right(Split(currentLine, "/")(0), Len(Split(currentLine, "/")(0)) - 1)
            On Error Resume Next
            workbookNumber.Worksheets(Sheets(workSheetNumber).Name).Activate
            If Err.Number = 9 Then
                MsgBox "Number of workseets is not matching tamplate"
                Sheets.Add After:=Sheets(Sheets.Count)
            End If
            splitLine = Split(currentLine, "/")
            lenRow = splitLine(2)
            lenCol = splitLine(3)
            strAddrTempWs = splitLine(4)
            ReDim rowArray(1 To lenRow, 1 To 3)
            
            'Celar unnecessary notes
            Set usedRangeWs = ActiveSheet.usedRange
            Call clearNotes(usedRangeWs)
            Call clearRgbErrCell(usedRangeWs)
            
            'Checking is table object in worksheet
            Let tableNum = ActiveSheet.ListObjects.Count
            If tableNum > 0 Then
                isATableObj = True
                ReDim headerRangeArr(1 To tableNum)
                ReDim cellRangeArr(1 To tableNum)
            Else
                isATableObj = False
            End If
            
            'Collect headers from tables
            If isATableObj = True Then
                For tableCounter = 1 To tableNum
                    Let cTempRangeStr = ActiveSheet.ListObjects.item(tableCounter) _
                    .DataBodyRange.Address(RowAbsolute:=False, ColumnAbsolute:=False)
                    Let hTempRangeStr = ActiveSheet.ListObjects.item(tableCounter) _
                    .HeaderRowRange.Address(RowAbsolute:=False, ColumnAbsolute:=False)
                    cellRangeArr(tableCounter) = cTempRangeStr
                    headerRangeArr(tableCounter) = hTempRangeStr
                Next tableCounter
            End If
            
            If outsiteSize = True Then
                Call handleOutWs(strAddrTempWs, anyOutsiteErr)
            End If
            
        Else 'Current row actions
            splitTopLine = Split(currentLine, "|")
            startAddress = Split(splitTopLine(0), "/")(0)
            endAddress = Split(splitTopLine(lenRow - 1), "/")(0)
            ReDim cellInfoArray(1 To 5)
            
            'Fill array of current row data
            Set rowRange = Range(startAddress & ":" & endAddress)
            For cellRange = 1 To rowRange.Count
                Call fillWsData(cellRange, isATableObj, rowArray, rowRange, _
                tableNum, headerRangeArr, cellRangeArr)
            Next cellRange
            '
            rowCounter = 0
            For cellCounter = 0 To lenRow - 1
                rowCounter = rowCounter + 1
                
                'Fill array of template data
                tempCell = splitTopLine(cellCounter)
                splitLine = Split(tempCell, "/")
                cellInfoArray(1) = splitLine(0)
                cellInfoArray(2) = CInt(splitLine(1))
                cellInfoArray(3) = splitLine(2)
                cellInfoArray(4) = CInt(splitLine(3))
                cellInfoArray(5) = splitLine(4)
                
                'Looking for errors
                addToArrStatus = False
                addStrToArray = ""
                
                Call ErrorsCheacker(cellCounter, cellInfoArray, rowArray, rowRange, _
                rowCounter, hardValueData, addToArrStatus, addStrToArray)
                
                addStrToArray = Left(addStrToArray, Len(addStrToArray) - 1)
                
                'Adding errors to Array
                If addToArrStatus = True Then
                    sizeOfExitArr = sizeOfExitArr + 1
                    ReDim Preserve exitArray(1 To sizeOfExitArr)
                    exitArray(sizeOfExitArr) = workSheetNumber & "/" & addStrToArray
                End If

            Next cellCounter
            
        End If

    Loop
    
    Close #1
    
Hell:
    'Exit array handling
    isArrEmpty = True
    If sizeOfExitArr > 0 Then isArrEmpty = False
    
    If isArrEmpty = False Then
        For eCounter = 1 To UBound(exitArray)
            Let howManyErrors = howManyErrors + 1
            Call addErrorToWs(exitArray, eCounter, errorArr, workbookNumber)
        Next eCounter
    End If
    
    If isArrEmpty = True And anyOutsiteErr = 0 Then
        MsgBox Prompt:="There is no error in current workbook.", _
        Buttons:=vbOKOnly, Title:="Error raport"
    ElseIf isArrEmpty = False And anyOutsiteErr = 0 Then
        Let summaryStr = "There is " + CStr(howManyErrors) + " errors in current workbook."
        MsgBox Prompt:=summaryStr, _
        Buttons:=vbOKOnly, Title:="Error raport"
    ElseIf isArrEmpty = True And anyOutsiteErr > 0 Then
        Let summaryStr = "There is " + CStr(anyOutsiteErr) + " outsite errors in current workbook."
        MsgBox Prompt:=summaryStr, _
        Buttons:=vbOKOnly, Title:="Error raport"
    Else
        Let howManyErrors = howManyErrors + anyOutsiteErr
        Let summaryStr = "There is " + CStr(howManyErrors) + " errors in current workbook."
        MsgBox Prompt:=summaryStr, _
        Buttons:=vbOKOnly, Title:="Error raport"
    End If
    
    Application.DisplayCommentIndicator = xlCommentIndicatorOnly
    Set Restore = New Main_DSC
    Restore.ApplicationRestore
    
End Sub

Private Sub clearNotes(usedRangeOfCurrWs)
    
    Dim noteCounter As Range
    Dim currentNote As String
    Dim boolNoteInStr As Boolean: boolNoteInStr = False
    Dim midNoteInStr As Boolean: midNoteInStr = False
    Dim splitedNote As String
    Dim midStrLen As String
    
    Set usedRangeOfCurrWs = ActiveSheet.usedRange
    For Each noteCounter In usedRangeOfCurrWs
        If noteCounter.NoteText <> "" Then
            currentNote = noteCounter.NoteText
            boolNoteInStr = InStr(currentNote, "DSC - hint")
            midStrLen = Mid(currentNote, 1, 10)
            If midStrLen = "DSC - hint" Then midNoteInStr = True
            If midNoteInStr = True Then
                noteCounter.Comment.Delete
            ElseIf boolNoteInStr = True Then
                splitedNote = Split(currentNote, "DSC - hint")(0)
                noteCounter.Comment.Text (splitedNote)
            End If
        End If
    Next noteCounter
    
End Sub

Private Sub clearRgbErrCell(usedRangeOfCurrWs)
    
    Dim rgbCounter As Range
    Dim currentColor As Long
    
    For Each rgbCounter In usedRangeOfCurrWs
        currentColor = rgbCounter.Interior.Color
        If currentColor = RGB(255, 146, 145) Then
            rgbCounter.Interior.Color = RGB(255, 255, 255)
            rgbCounter.BorderAround LineStyle:=xlContinuous, _
            Weight:=xlThin, ColorIndex:=15, Color:=0
        End If
    Next rgbCounter
    
End Sub

Private Sub handleOutWs(strAddrTempWs, outsiteError)
    '
    'Handle outside of tamplate
    '
    
    Dim strUsedRangeWs As String
    Dim strUsedRangeTemp() As String
    Dim strUsedRangeCurr() As String
    Dim northAddrStr, southAddrStr, westAddrStr, eastAddrStr As String
    Dim northCheck, southCheck, westCheck, eastCheck As Range
    Dim westTrasform, eastTransform As Range
    Dim westTransCol, eastTransCol As Integer
    Dim westTransCell, eastTransCell As String
    Dim boolNorth As Boolean: boolNorth = False
    Dim boolSouth As Boolean: boolSouth = False
    Dim boolWest As Boolean: boolWest = False
    Dim boolEast As Boolean: boolEast = False
    Dim wrBoolArr() As Variant
    Dim wrCounter As Range
    Dim currentRange As Range
    Dim counterInt As Byte
    Dim isVal As Boolean
    Dim oldComm As String
    Dim commentToAdd As String
    
    Let strUsedRangeWs = ActiveSheet.usedRange.Address(RowAbsolute:=True, ColumnAbsolute:=False)
    
    'Splits for outsite checking
    strUsedRangeTemp = Split(Replace(strAddrTempWs, ":", "$"), "$")
    strUsedRangeCurr = Split(Replace(strUsedRangeWs, ":", "$"), "$")
    
    ReDim wrBoolArr(1 To 4)
    Let wrBoolArr = Array(False, False, False, False)
    If strUsedRangeTemp(1) > 1 Then boolNorth = True: wrBoolArr(0) = True
    If strUsedRangeTemp(3) < 1048576 Then boolSouth = True: wrBoolArr(1) = True
    If strUsedRangeTemp(0) <> "A" Then boolWest = True: wrBoolArr(2) = True
    If strUsedRangeTemp(2) <> "XFD" Then boolEast = True: wrBoolArr(3) = True
    
    'North
    If boolNorth = True Then
        If strUsedRangeTemp(1) = strUsedRangeCurr(1) Then
            northAddrStr = strUsedRangeCurr(0) + CStr(CInt(strUsedRangeTemp(1)) - 1) + ":" _
            + strUsedRangeCurr(2) + CStr(CInt(strUsedRangeCurr(1)) - 1)
        Else
            northAddrStr = strUsedRangeCurr(0) + strUsedRangeCurr(1) + ":" + strUsedRangeCurr(2) _
            + CStr(CInt(strUsedRangeTemp(1)) - 1)
        End If
        
        Set northCheck = Range(northAddrStr)
    End If
    
    'South
    If boolSouth = True Then
        If strUsedRangeTemp(3) = strUsedRangeCurr(3) Then
            southAddrStr = strUsedRangeCurr(0) + CStr(CInt(strUsedRangeTemp(3)) + 1) + ":" _
            + strUsedRangeCurr(2) + CStr(CInt(strUsedRangeCurr(3)) + 1)
        Else
            southAddrStr = strUsedRangeCurr(0) + CStr(CInt(strUsedRangeTemp(3)) + 1) + ":" _
            + strUsedRangeCurr(2) + strUsedRangeCurr(3)
        End If
        
        Set southCheck = Range(southAddrStr)
    End If
    
    'West
    If boolWest = True Then
        Set westTrasform = Range(strUsedRangeTemp(0) + strUsedRangeCurr(3))
        westTransCol = westTrasform.Column - 1
        westTransCell = Cells(strUsedRangeCurr(3), westTransCol).Address(RowAbsolute:=False, _
        ColumnAbsolute:=False)
        If strUsedRangeTemp(0) = strUsedRangeCurr(0) Then
            westAddrStr = westTransCell + ":" + Mid(westTransCell, 1, 1) + strUsedRangeCurr(1)
        Else
            westAddrStr = strUsedRangeCurr(0) + strUsedRangeCurr(1) + ":" + westTransCell
        End If
        
        Set westCheck = Range(westAddrStr)
    End If
    
    'East
    If boolEast = True Then
        Set eastTransform = Range(strUsedRangeTemp(2) + strUsedRangeCurr(1))
        eastTransCol = eastTransform.Column + 1
        eastTransCell = Cells(strUsedRangeCurr(1), eastTransCol).Address(RowAbsolute:=False, _
        ColumnAbsolute:=False)
        If strUsedRangeTemp(2) = strUsedRangeCurr(2) Then
            eastAddrStr = eastTransCell + ":" + Mid(eastTransCell, 1, 1) + strUsedRangeCurr(3)
        Else
            eastAddrStr = eastTransCell + ":" + strUsedRangeCurr(2) + strUsedRangeCurr(3)
        End If
        
        Set eastCheck = Range(eastAddrStr)
    End If
    
    For counterInt = 0 To UBound(wrBoolArr)
        If wrBoolArr(counterInt) = True Then
            If counterInt = 0 Then Set currentRange = northCheck
            If counterInt = 1 Then Set currentRange = southCheck
            If counterInt = 2 Then Set currentRange = westCheck
            If counterInt = 3 Then Set currentRange = eastCheck
            For Each wrCounter In currentRange
                Let isVal = False
                If wrCounter.value <> "" Then isVal = True
                If isVal = True Then
                    wrCounter.Interior.Color = RGB(255, 146, 145)
                    Let commentToAdd = "DSC - hint" + vbNewLine + _
                    "Cell is outsite of template size."
                    If wrCounter.Comment Is Nothing Then
                        oldComm = ""
                    Else
                        oldComm = wrCounter.NoteText
                    End If
                    
                    If oldComm = "" Then
                        wrCounter.AddComment (commentToAdd)
                    ElseIf oldComm = commentToAdd Then
                        'pass
                    Else
                        commentToAdd = oldComm + vbNewLine + vbNewLine + commentToAdd
                        wrCounter.ClearComments
                        wrCounter.AddComment (commentToAdd)
                    End If
                    
                    Let wrCounter.Comment.Visible = True
                    Let outsiteError = outsiteError + 1
                    
                End If
            Next wrCounter
        End If
    Next counterInt
    
    
End Sub

Private Sub fillWsData(cell, isTable, rowArr, rowRan, tNum, hrStr, clStr)
    '
    'Fill array of data from current row
    '
    
    Dim cellRange As Range
    Dim cellCounter As Byte
    
    Dim headerRange As Range
    Dim headerCounter As Byte
    
    rowArr(cell, 1) = rowRan(cell).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    rowArr(cell, 2) = VarType(rowRan(cell))
    
    'check cell and header (c, h, n) type
    Let rowArr(cell, 3) = "n"
    If isTable = True Then
        For cellCounter = 1 To UBound(clStr) 'cell checker (c)
            Set cellRange = Range(clStr(cellCounter))
            If Not Intersect(rowRan(cell), cellRange) Is Nothing Then
                rowArr(cell, 3) = "c"
            End If
        Next cellCounter
        For headerCounter = 1 To UBound(hrStr) 'header checker (h)
            Set headerRange = Range(hrStr(headerCounter))
            If Not Intersect(rowRan(cell), headerRange) Is Nothing Then
                rowArr(cell, 3) = "h"
            End If
        Next headerCounter
    End If
    
End Sub

Private Sub ErrorsCheacker(cCounter, cellInfoArray, rowArray, rRange, rCounter, _
hvData, addStatus, addStr)
    '
    ' Looking for errors
    '
    
    If cellInfoArray(1) <> rowArray(cCounter + 1, 1) Then
        addStatus = True
        addStr = addStr + cellInfoArray(1) + "/"
    Else
        addStr = addStr + cellInfoArray(1) + "/"
    End If
    If cellInfoArray(2) <> rowArray(cCounter + 1, 2) Then
        addStatus = True
        addStr = addStr + CStr(cellInfoArray(2)) + "/"
    Else
        addStr = addStr + "/"
    End If
    If cellInfoArray(3) <> rowArray(cCounter + 1, 3) Then
        addStatus = True
        addStr = addStr + cellInfoArray(3) + "/"
    Else
        addStr = addStr + "/"
    End If
    If cellInfoArray(4) = 1 Then
        hvData = rRange(1, rCounter).value
        hvData = Replace(hvData, "/", "-")
        If cellInfoArray(5) <> "<" + hvData + "]" Then
            addStatus = True
            addStr = addStr + cellInfoArray(5) + "/"
        Else
            addStr = addStr + "/"
        End If
    ElseIf cellInfoArray(4) = 2 Then
        addStatus = False
        addStr = addStr + "/"
    Else
        addStr = addStr + "/"
    End If
    
End Sub

Private Sub addErrorToWs(arrPart, counter, errArr, wbNum)
    '
    'Adding errors to worksheet by comments
    '
    
    Dim oldComment As String
    Dim newComment As String
    Dim currRange As Range
    Dim tableType As String
    Dim valueStr As String
    Dim oldCommState As String
    Dim lenErrArr As Byte
    Dim valueVarType As String
    Dim VarType As New Main_DSC
    Dim vtCollection As New collection
    Dim keyCollection As String


    errArr = Split(arrPart(counter), "/")
    wbNum.Worksheets(Sheets(CInt(errArr(0))).Name).Activate
    Set currRange = Range(errArr(1))
    currRange.Interior.Color = RGB(255, 146, 145)
    
    Let newComment = "DSC - hint"
    
    'Create Dictionary
    If errArr(2) <> "" Then
        Let keyCollection = Trim(CStr(errArr(2)))
        Set VarType = New Main_DSC
        Call VarType.VarTypeCollection(keyCollection, valueVarType, vtCollection)
    End If
    
    'Input variable type
    If errArr(2) <> "" Then
        newComment = newComment + vbNewLine + "Type of variable should be " + valueVarType + "."
    End If
    
    'Input other type
    Let tableType = ""
    If errArr(3) <> "" Then
        If errArr(3) = "n" Then tableType = "outsite the Table"
        If errArr(3) = "c" Then tableType = " Table cell"
        If errArr(3) = "h" Then tableType = " Table header"
        
        newComment = newComment + vbNewLine + "Cell should be " + tableType + "."
    End If
    
    'Input missing text
    If errArr(4) <> "" Then
        valueStr = errArr(4)
        valueStr = Left(valueStr, Len(valueStr) - 1)
        valueStr = Right(valueStr, Len(valueStr) - 1)
        newComment = newComment + vbNewLine + "Value in cell should be exacly '" + valueStr + "'."
    End If
    
    'Add comment
    If currRange.Comment Is Nothing Then
        oldComment = ""
    Else
        oldComment = currRange.NoteText
    End If
    
    If oldComment = "" Then
        currRange.AddComment (newComment)
    ElseIf oldComment = newComment Then
        'pass
    Else
        newComment = oldComment + vbNewLine + vbNewLine + newComment
        currRange.ClearComments
        currRange.AddComment (newComment)
    End If
    
    Let currRange.Comment.Visible = True
    
End Sub
