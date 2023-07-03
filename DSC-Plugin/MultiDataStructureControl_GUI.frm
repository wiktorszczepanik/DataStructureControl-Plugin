VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MultiDSC_GUI 
   Caption         =   "Multi Data Structure Control"
   ClientHeight    =   7455
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9765.001
   OleObjectBlob   =   "MultiDataStructureControl_GUI.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MultiDSC_GUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'-------------------------
'Gorup 1 - Data Management
'-------------------------

'G1, Subgroup 1 - DM
Public Sub mainRun_Click()
    '
    'Checking all selected files in files in
    '
    
    Dim wsPath As String
    Dim splitWsPath() As String
    Dim wsPathInput, wsPathCorrect As String
    Dim splitTemplatePath() As String
    Dim wsTemplatePath As String
    Dim listCountInp, listCountCorr As Integer
    Dim listBoxCountAll As Integer
    Dim fileArrInp() As String
    Dim iI, counterSel, counterSelCheck As Integer
    Dim arrInput() As String
    Dim arrCorrect() As String
    Dim fileManipulation, fileInputDC As String
    Dim errorCellsOut As Boolean
    Dim wbValid As Boolean: wbValid = True
    Dim filesWithErrorCount As Integer
    Dim FSO As Object
    Const unSel = "Unselected"
    Dim actionToAll As VbMsgBoxResult
    Dim pctDone As Single
    Dim pctTotal As Integer
    Dim pctLblWidth As Integer
    Dim pctCounter As Integer
    Dim iSel, cSel As Integer
    Dim tdI, tdC As Integer
    Dim spltTdI, spltTdC  As String
    Dim ehCorr As Integer
    Dim isCorSelEh As Boolean: isCorSelEh = False
    Dim arrCorrBool As Boolean: arrCorrBool = False
    Dim textC As String
    
    'Load error handling
    Let wsPath = workspacePath.Caption
    Let wsTemplatePath = templatePath.Caption
    If wsPath = unSel And wsTemplatePath = unSel Then
        MsgBox "Workspace and Template file are not selected"
        Exit Sub
    ElseIf wsTemplatePath = unSel Then
        MsgBox "Template file is not selected"
        Exit Sub
    ElseIf wsPath = unSel Then
        MsgBox "Workspace was not selected"
        Exit Sub
    End If
    
    splitWsPath = Split(wsPath, " ")
    wsPath = splitWsPath(2)
    wsPathInput = wsPath + "Input_Files"
    wsPathCorrect = wsPath + "Correct_Files"
    listCountInp = listBoxInputFiles.ListCount
    listCountCorr = listBoxCorrectFiles.ListCount
    
    splitTemplatePath = Split(wsTemplatePath, " ")
    wsTemplatePath = splitTemplatePath(2)
    
    'Selection error handling
    For counterSel = 0 To listCountInp - 1
        If listBoxInputFiles.Selected(counterSel) = True Then
            listBoxCountAll = listBoxCountAll + 1
        End If
    Next counterSel
    For ehCorr = 0 To listCountCorr - 1
        If listBoxCorrectFiles.Selected(ehCorr) = True Then isCorSelEh = True
    Next ehCorr
    If isCorSelEh = True Then MsgBox Prompt:= _
    "Macro will not work on items form Correct Files"
    
    If listBoxCountAll > 0 Then ReDim arrInput(1 To listBoxCountAll)
    
    If listBoxCountAll = 0 Then
        actionToAll = MsgBox("None of items form Input files were selected." & vbNewLine & _
        "Do you want to perform an action for all the Files in input directory?", vbYesNo)
        If actionToAll = vbYes Then
            For counterSelCheck = 0 To listCountInp - 1
                listBoxInputFiles.Selected(counterSelCheck) = True
            Next counterSelCheck
        Else
            Exit Sub
        End If
    ElseIf listBoxCountAll = 1 Then
        actionToAll = MsgBox("One file is selected." & vbNewLine & _
        "Do you want to perform an action?", vbYesNo)
        If actionToAll = vbNo Then Exit Sub
    Else
        actionToAll = MsgBox("(" & listBoxCountAll & ") files are selected." & vbNewLine & _
        "Do you want to perform an action?", vbYesNo)
        If actionToAll = vbNo Then Exit Sub
    End If
    
    Let errorCellsOut = checkOutsiteTemplate.value
    pctTotal = listBoxCountAll
    barProgress.BackColor = vbHighlight
    listBoxCountAll = 0
    filesWithErrorCount = 0
    barProgress.Width = 0
    pctLblWidth = 454
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    'run check for input list box
    Let listBoxCountAll = 0
    Let fileManipulation = ""
    Let pctCounter = 0
    For iI = 0 To listCountInp - 1
        On Error Resume Next
        If listBoxInputFiles.Selected(iI) = True Then
            pctCounter = pctCounter + 1
            wbValid = True
            listBoxCountAll = listBoxCountAll + 1
            fileManipulation = Split(listBoxInputFiles.List(iI), ".")(1) _
            + "." + Split(listBoxInputFiles.List(iI), ".")(2)
            fileManipulation = Right(fileManipulation, Len(fileManipulation) - 1)
            fileInputDC = wsPathInput & "\" & fileManipulation
            
            'run clear macro
            Call DataCompare(fileInputDC, wsTemplatePath, errorCellsOut, wbValid)
            If wbValid = False Then
                filesWithErrorCount = filesWithErrorCount + 1
                arrInput(pctCounter) = fileManipulation
            Else
                ReDim Preserve arrCorrect(1 To pctCounter)
                FSO.MoveFile wsPathInput & "\" & fileManipulation, wsPathCorrect & "\" & fileManipulation
                arrCorrect(pctCounter) = fileManipulation
                arrCorrBool = True
            End If
            
            pctDone = (pctCounter) / pctTotal
            barProgress.Width = pctLblWidth * pctDone
            frameProgress.Caption = Format(pctDone, "0%")
            
            DoEvents
            
        End If
    Next iI
    
    frameProgress.Caption = "0%"
    barProgress.BackColor = RGB(255, 255, 255)
    
    Call reloadData_Click
    
    'Selected input files
    On Error GoTo emptyDirHandler
    If UBound(arrInput) > 0 Then
    listCountInp = listBoxInputFiles.ListCount
    For iSel = 0 To listCountInp - 1
        For tdI = 1 To UBound(arrInput)
            spltTdI = Split(listBoxInputFiles.List(iSel), ". ")(1)
            If spltTdI = arrInput(tdI) Then
                listBoxInputFiles.Selected(iSel) = True
            End If
        Next tdI
    Next iSel
    End If
    
    'Selected correct files
    If isCorSelEh = True Then
    listCountCorr = listBoxCorrectFiles.ListCount
    For cSel = 0 To listCountCorr - 1
        For tdC = 1 To UBound(arrCorrect)
            spltTdC = Split(listBoxCorrectFiles.List(cSel), ". ")(1)
            textC = arrCorrect(tdC)
            If spltTdC = textC Then
                listBoxCorrectFiles.Selected(cSel) = True
            End If
        Next tdC
    Next cSel
    End If

Exit Sub
emptyDirHandler:
    MsgBox "Input Files directory is empty"

End Sub

'G1, Subgroup 1 - DM
Private Sub DataCompare(pathToFile, pathToTemplate, doErrorOutsite, valid)
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
    
    Let valid = True
    
    Let workbookName = pathToFile
    Set workbookNumber = Workbooks.Open(workbookName)
    Let outsiteSize = doErrorOutsite
    Let selectedXlcgPath = pathToTemplate
    
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
                Sheets.Add After:=Sheets(Sheets.Count)
                valid = False
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
                    valid = False
                End If

            Next cellCounter
            
        End If

    Loop
    
    Close #1
    
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
        valid = True
    ElseIf isArrEmpty = False And anyOutsiteErr = 0 Then
        valid = False
    ElseIf isArrEmpty = True And anyOutsiteErr > 0 Then
        valid = False
    Else
        valid = False
    End If
    
    Application.DisplayCommentIndicator = xlCommentIndicatorOnly
    workbookNumber.Close SaveChanges:=True
    
    Set Restore = New Main_DSC
    Restore.ApplicationRestore
    
End Sub

'G1, Subgroup 1 - DM
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

'G1, Subgroup 1 - DM
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

'G1, Subgroup 1 - DM
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

'G1, Subgroup 1 - DM
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

'G1, Subgroup 1 - DM
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

'G1, Subgroup 1 - DM
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

'G1, Subgroup 2 - DM
Public Sub selectListedTypeFiles_Click()
    '
    'button select all values from items list
    '
    
    Dim wsPath As String
    Dim splitWsPath() As String
    Dim listBoxItemsInp, listBoxItemsCorr As Integer
    Dim item As Integer
    Dim inpOrCorr As Boolean: inpOrCorr = False
    Dim selectionIC As Byte
    
    On Error GoTo Hell
    Let wsPath = workspacePath.Caption
    splitWsPath = Split(wsPath, " ")
    
    'checking which option button is selected
    If optionInput.value = True Or optionCorrect.value = True _
    Or optionAll.value = True Then
        inpOrCorr = True
        If optionInput.value = True Then
            Let selectionIC = 0
            listBoxItemsInp = listBoxInputFiles.ListCount
        ElseIf optionCorrect.value = True Then
            Let selectionIC = 1
            listBoxItemsCorr = listBoxCorrectFiles.ListCount
        Else
            Let selectionIC = 2
            listBoxItemsInp = listBoxInputFiles.ListCount
            listBoxItemsCorr = listBoxCorrectFiles.ListCount
        End If
    End If
    
    'run selection on input files list
    If selectionIC = 0 Or selectionIC = 2 Then
    For item = 0 To listBoxItemsInp - 1
        If listBoxInputFiles.Selected(item) = False Then
            listBoxInputFiles.Selected(item) = True
        End If
    Next item
    End If
    
    'run selection on correct files list
    If selectionIC = 1 Or selectionIC = 2 Then
    For item = 0 To listBoxItemsCorr - 1
        If listBoxCorrectFiles.Selected(item) = False Then
            listBoxCorrectFiles.Selected(item) = True
        End If
    Next item
    End If
    
Hell: 'error handler
    If wsPath = "Unselected" Then
        MsgBox "Workspace is not selected"
    ElseIf inpOrCorr = False And item = 0 Then
        MsgBox "List of items is not selected"
    ElseIf item = 0 Then
        MsgBox "List does not contain any items"
    End If

End Sub

'G1, Subgroup 2 - DM
Public Sub deselectButton_Click()
    '
    'button deselect all values from items list
    '
    
    Dim wsPath As String
    Dim splitWsPath() As String
    Dim listBoxItemsInp, listBoxItemsCorr As Integer
    Dim item As Integer
    Dim inpOrCorr As Boolean: inpOrCorr = False
    Dim selectionIC As Byte
    
    On Error GoTo Hell
    Let wsPath = workspacePath.Caption
    splitWsPath = Split(wsPath, " ")
    
    'checking which option button is selected
    If optionInput.value = True Or optionCorrect.value = True _
    Or optionAll.value = True Then
        inpOrCorr = True
        If optionInput.value = True Then
            Let selectionIC = 0
            listBoxItemsInp = listBoxInputFiles.ListCount
        ElseIf optionCorrect.value = True Then
            Let selectionIC = 1
            listBoxItemsCorr = listBoxCorrectFiles.ListCount
        Else
            Let selectionIC = 2
            listBoxItemsInp = listBoxInputFiles.ListCount
            listBoxItemsCorr = listBoxCorrectFiles.ListCount
        End If
    End If
    
    'run deselection on input files list
    If selectionIC = 0 Or selectionIC = 2 Then
    For item = 0 To listBoxItemsInp - 1
        If listBoxInputFiles.Selected(item) = True Then
            listBoxInputFiles.Selected(item) = False
        End If
    Next item
    End If
    
    'run deselection on correct files list
    If selectionIC = 1 Or selectionIC = 2 Then
    For item = 0 To listBoxItemsCorr - 1
        If listBoxCorrectFiles.Selected(item) = True Then
            listBoxCorrectFiles.Selected(item) = False
        End If
    Next item
    End If
    
Hell: 'error handler
    If wsPath = "Unselected" Then
        MsgBox "Workspace is not selected"
    ElseIf inpOrCorr = False And item = 0 Then
        MsgBox "List of items is not selected"
    ElseIf item = 0 Then
        MsgBox "List does not contain any items"
    End If
    
End Sub

'G1, Subgroup 3 - DM
Public Sub clearErrors_Click()
    '
    'clear added errors over selected files
    '
    
    Dim wsPath As String
    Dim splitWsPath() As String
    Dim wsPathInput, wsPathCorrect As String
    Dim listCountInp, listCountCorr As Integer
    Dim listBoxCountAll As Integer
    Dim fileArrInp() As String
    Dim fileArrCorr() As String
    Dim iI, iC, counterSelInp, counterSelCorr As Integer
    Dim counterSelCheckInp, counterSelCheckCorr As Integer
    Dim arrInput() As String
    Dim arrCorrect() As String
    Dim fileManipulation As String
    Const unSel = "Unselected"
    Dim actionToAll As VbMsgBoxResult
    Dim inpBoolSel As Boolean: inpBoolSel = False
    Dim corrBoolSel As Boolean: corrBoolSel = False
    Dim listCalc As Integer
    Dim strForMsgInp, strForMsgCorr As String
    Dim pctDone As Single
    Dim pctTotal As Integer
    Dim pctLblWidth As Integer
    
    'Load error handling
    Let wsPath = workspacePath.Caption
    If wsPath = unSel Then
        MsgBox "Workspace file is not selected"
        Exit Sub
    End If
    
    splitWsPath = Split(wsPath, " ")
    wsPath = splitWsPath(2)
    wsPathInput = wsPath + "Input_Files"
    wsPathCorrect = wsPath + "Correct_Files"
    listCountInp = listBoxInputFiles.ListCount
    listCountCorr = listBoxCorrectFiles.ListCount
    
    'Selection error handling
    listBoxCountAll = 0
    listCalc = 0
    
    For counterSelInp = 0 To listCountInp - 1
        If listBoxInputFiles.Selected(counterSelInp) = True Then
            listBoxCountAll = listBoxCountAll + 1
            inpBoolSel = True
            listCalc = listCalc + 1
        End If
    Next counterSelInp
    For counterSelCorr = 0 To listCountCorr - 1
        If listBoxCorrectFiles.Selected(counterSelCorr) = True Then
            listBoxCountAll = listBoxCountAll + 1
            corrBoolSel = True
        End If
    Next counterSelCorr
    
    If listBoxCountAll = 0 Then
        actionToAll = MsgBox("None of items form directories were selected." & vbNewLine & _
        "Do you want to perform an action for all files?", vbYesNo)
        If actionToAll = vbYes Then
            For counterSelCheckInp = 0 To listCountInp - 1
                listBoxInputFiles.Selected(counterSelCheckInp) = True
            Next counterSelCheckInp
            For counterSelCheckCorr = 0 To listCountCorr - 1
                listBoxCorrectFiles.Selected(counterSelCheckCorr) = True
            Next counterSelCheckCorr
        Else
            Exit Sub
        End If
    ElseIf inpBoolSel = True And corrBoolSel = True Then
        If listCalc = 1 Then
            strForMsgInp = "(1) file form Input Directory"
        ElseIf listCalc > 1 Then
            strForMsgInp = "(" & listCalc & ") files from Input Directory"
        End If
        If listBoxCountAll - listCalc = 1 Then
            strForMsgCorr = "(1) file form Correct Directory"
        ElseIf listBoxCountAll - listCalc > 1 Then
            strForMsgCorr = "(" & listBoxCountAll - listCalc & ") files from Correct Directory"
        End If
        
        actionToAll = MsgBox("Selected:" & vbNewLine & strForMsgInp & vbNewLine & strForMsgCorr & _
        vbNewLine & "Do you want to perform an action?", vbYesNo)
        If actionToAll = vbNo Then Exit Sub
        
    ElseIf inpBoolSel = True Then
        If listCalc = 1 Then
            strForMsgInp = "One file was selected form Input Directory"
        ElseIf listCalc > 1 Then
            strForMsgInp = "(" & listCalc & ") files were selected from Input Directory"
        End If
        
        actionToAll = MsgBox(strForMsgInp & vbNewLine & "Do you want to perform an action?", vbYesNo)
        If actionToAll = vbNo Then Exit Sub
        
    Else
        If listBoxCountAll - listCalc = 1 Then
            strForMsgCorr = "One file was selected form Correct Directory"
        ElseIf listBoxCountAll - listCalc > 1 Then
            strForMsgCorr = "(" & listBoxCountAll - listCalc & ") files were selected from Correct Directory"
        End If
        
        actionToAll = MsgBox(strForMsgCorr & vbNewLine & "Do you want to perform an action?", vbYesNo)
        If actionToAll = vbNo Then Exit Sub
        
    End If
    
    barProgress.BackColor = vbHighlight
    barProgress.Width = 0
    pctLblWidth = 454
    pctTotal = listBoxCountAll
    listBoxCountAll = 0
    
    'run clear for input list box
    Let fileManipulation = ""
    For iI = 0 To listCountInp - 1
        On Error Resume Next
        If listBoxInputFiles.Selected(iI) = True Then
            listBoxCountAll = listBoxCountAll + 1
            fileManipulation = Split(listBoxInputFiles.List(iI), ".")(1) _
            + "." + Split(listBoxInputFiles.List(iI), ".")(2)
            fileManipulation = Right(fileManipulation, Len(fileManipulation) - 1)
            fileManipulation = wsPathInput & "\" & fileManipulation
            
            'run clear macro
            Call clearErrorsOverList(fileManipulation)
            
            pctDone = (listBoxCountAll) / pctTotal
            barProgress.Width = pctLblWidth * pctDone
            frameProgress.Caption = Format(pctDone, "0%")
            
            DoEvents
            
        End If
    Next iI
    
    'run clear for correct list box
    Let fileManipulation = ""
    For iC = 0 To listCountCorr - 1
        If listBoxCorrectFiles.Selected(iC) = True Then
            listBoxCountAll = listBoxCountAll + 1
            fileManipulation = Split(listBoxCorrectFiles.List(iC), ".")(1) _
            + "." + Split(listBoxCorrectFiles.List(iC), ".")(2)
            fileManipulation = Right(fileManipulation, Len(fileManipulation) - 1)
            fileManipulation = wsPathCorrect & "\" & fileManipulation
            
            'run clear macro
            Call clearErrorsOverList(fileManipulation)
            
            pctDone = (listBoxCountAll) / pctTotal
            barProgress.Width = pctLblWidth * pctDone
            frameProgress.Caption = Format(pctDone, "0%")
            
            DoEvents
            
        End If
    Next iC
    
    frameProgress.Caption = "0%"
    barProgress.BackColor = RGB(255, 255, 255)
    
End Sub

'G1, Subgroup 3 - DM
Private Sub clearErrorsOverList(path)
    '
    'clears error in selected file
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
    
    Let workBookPath = path
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
    
    workbookAction.Close SaveChanges:=True
    
    Set Restore = New Main_DSC
    Restore.ApplicationRestore
    
End Sub

'G1, Subgroup 4 - DM
Public Sub openFile_Click()
    '
    'button is opening selected files
    '
    
    Dim wsPath As String
    Dim splitWsPath() As String
    Dim wsPathInput, wsPathCorrect As String
    Dim listCountInp, listCountCorr As Integer
    Dim listBoxCountAll As Integer
    Dim fileArrInp() As String
    Dim fileArrCorr() As String
    Dim iI, iC, counterSel, counterSelCheck As Integer
    Dim arrInput() As String
    Dim arrCorrect() As String
    Dim fileManipulation As String
    Dim actionToAll As VbMsgBoxResult
    Const unSel = "Unselected"
    
    Let wsPath = workspacePath.Caption
    
    'Load error handling
    Let wsPath = workspacePath.Caption
    If wsPath = unSel Then
        MsgBox "Workspace is not selected"
        Exit Sub
    End If
    
    splitWsPath = Split(wsPath, " ")
    wsPath = splitWsPath(2)
    wsPathInput = wsPath + "Input_Files"
    wsPathCorrect = wsPath + "Correct_Files"
    listCountInp = listBoxInputFiles.ListCount
    listCountCorr = listBoxCorrectFiles.ListCount
    
    'Selection error handling
    For counterSel = 0 To listCountInp - 1
        If listBoxInputFiles.Selected(counterSel) = True Then
            listBoxCountAll = listBoxCountAll + 1
        End If
    Next counterSel
    counterSel = 0
    For counterSel = 0 To listCountCorr - 1
        If listBoxCorrectFiles.Selected(counterSel) = True Then
            listBoxCountAll = listBoxCountAll + 1
        End If
    Next counterSel
    
    If listBoxCountAll = 0 Then
        actionToAll = MsgBox("None of items form both directories were selected")
        Exit Sub
    End If
    
    listBoxCountAll = 0
    
    'Open form input list box
    Let fileManipulation = ""
    For iI = 0 To listCountInp
        On Error Resume Next
        If listBoxInputFiles.Selected(iI) = True Then
            listBoxCountAll = listBoxCountAll + 1
            fileManipulation = Split(listBoxInputFiles.List(iI), ".")(1) _
            + "." + Split(listBoxInputFiles.List(iI), ".")(2)
            fileManipulation = Right(fileManipulation, Len(fileManipulation) - 1)
            fileManipulation = wsPathInput & "\" & fileManipulation
            Workbooks.Open fileManipulation
        End If
    Next iI
    
    'Open form correct list box
    Let fileManipulation = ""
    For iC = 0 To listCountCorr
        If listBoxCorrectFiles.Selected(iC) = True Then
            listBoxCountAll = listBoxCountAll + 1
            fileManipulation = Split(listBoxCorrectFiles.List(iC), ".")(1) _
            + "." + Split(listBoxCorrectFiles.List(iC), ".")(2)
            fileManipulation = Right(fileManipulation, Len(fileManipulation) - 1)
            fileManipulation = wsPathCorrect & "\" & fileManipulation
            Workbooks.Open fileManipulation
        End If
    Next iC
    
End Sub

'---------------------------
'Gorup 2 - Load Dependencies
'---------------------------

'G2, Subgroup 1 - LD
Public Sub loadWorkspace_Click()
    '
    'load workspace to window
    '
    
    Dim filePath As String
    Dim captionText As String
    Dim arrInp() As String
    Dim arrCorr() As String
    Dim dirErr As Boolean: dirErr = False
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select workspace folder"
        .ButtonName = "Select Folder"
        If .Show = 0 Then Exit Sub
        filePath = .SelectedItems(1) & "\"
    End With
    
    'change status of workspace
    Let captionText = "Selected"
    If filePath <> "" Then
        workspacePath.Caption = captionText & " - " & filePath
        Call inputFilesPopulate(filePath, arrInp, dirErr)
        If dirErr = True Then
            MsgBox "Selected workspace was not created by DSC macro" + vbNewLine + _
            "(Remember to select root directory)"
            workspacePath.Caption = "Unselected"
            listBoxInputFiles.Clear
            listBoxCorrectFiles.Clear
            Exit Sub
        End If
        Call correctFilesPopulate(filePath, arrCorr)
    End If
    
End Sub

'G2, Subgroup 2 - LD
Public Sub selectTemplate_Click()
    '
    'button is selecting already created workspace
    '
    
    Dim xlcgFile As String
    Dim captionText As String
    
    With Application.FileDialog(msoFileDialogOpen)
        .Title = "Select Template File"
        .ButtonName = "Select File"
        .Filters.Add "Templete Files", "*.xlcg?", 1
        .AllowMultiSelect = False
        If .Show = 0 Then Exit Sub
        xlcgFile = .SelectedItems(1)
    End With
    
    'add path to window
    Let captionText = "Selected"
    If xlcgFile <> "" Then templatePath.Caption = captionText & " - " & xlcgFile
    
End Sub

'--------------------
'Gorup 3 - Move Files
'--------------------

'G3, Subgroup 1 - MF
Private Sub reloadData_Click()
    '
    'button is reload data in input folder and correct folder
    '
    
    Dim wsPath As String
    Dim splitWsPath() As String
    Dim fileArrInp() As String
    Dim fileArrCorr() As String
    Const unSel = "Unselected"
    Dim dirErr As Boolean: dirErr = False
    
    'Error handling
    Let wsPath = workspacePath.Caption
    If wsPath = unSel Then
        MsgBox "Workspace is not selected"
        Exit Sub
    End If
    
    splitWsPath = Split(wsPath, " ")
    wsPath = splitWsPath(2)
    
    'Reload in input files and correct files
    If wsPath <> "" Then
        Call inputFilesPopulate(wsPath, fileArrInp, dirErr)
        Call correctFilesPopulate(wsPath, fileArrCorr)
    End If
    
End Sub

'G3, Subgroup 2 - MF
Private Sub filerMoveRight_Click()
    '
    'move items from input files folder to correct files folder
    '
    
    Dim wsPath, destPath As String
    Dim splitWsPath() As String
    Dim fileName As String
    Dim listBoxSelectedCount As Integer
    Dim listBoxSelectedArr() As String
    Dim listBoxCount, destListCount As Integer
    Dim fileManipulation As String
    Dim FSO As Object
    Dim i, d, c As Integer
    Dim strNameItem As String
    Dim strNameItemSplit() As String
    Dim listText As String
    Const unSel = "Unselected"
    
    'Error handling
    Let wsPath = workspacePath.Caption
    If wsPath = unSel Then
        MsgBox "Workspace is not selected"
        Exit Sub
    End If
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    splitWsPath = Split(wsPath, " ")
    wsPath = splitWsPath(2) + "Input_Files"
    destPath = splitWsPath(2) + "Correct_Files"
    listBoxSelectedCount = listBoxInputFiles.ListCount
    destListCount = listBoxCorrectFiles.ListCount
    
    'changing location action (I to C)
    listBoxCount = 0
    For i = 0 To listBoxSelectedCount
        On Error GoTo outOfLoop
        If listBoxInputFiles.Selected(i) = True Then
            listBoxCount = listBoxCount + 1
            ReDim Preserve listBoxSelectedArr(1 To listBoxCount)
            fileManipulation = Split(listBoxInputFiles.List(i), ".")(1) _
            + "." + Split(listBoxInputFiles.List(i), ".")(2)
            fileManipulation = Right(fileManipulation, Len(fileManipulation) - 1)
            
            FSO.MoveFile wsPath & "\" & fileManipulation, destPath & "\" & fileManipulation
            
            listBoxSelectedArr(listBoxCount) = fileManipulation
        End If
    Next i
    
outOfLoop:
    Call reloadData_Click
    
    If listBoxCount = 0 Then Exit Sub
    For d = 0 To destListCount + listBoxCount - 1
        strNameItem = listBoxCorrectFiles.List(d)
        strNameItemSplit = Split(strNameItem, ". ")
        strNameItem = strNameItemSplit(1)
        For c = 0 To UBound(listBoxSelectedArr) - 1
            On Error GoTo outOfSecLoop
            listText = listBoxSelectedArr(c + 1)
            If strNameItem = listText Then
                listBoxCorrectFiles.Selected(d) = True
            End If
        Next c
    Next d
    
outOfSecLoop:
    
End Sub

'G3, Subgroup 3 - MF
Private Sub filesMoveLeft_Click()
    '
    'move items from correct files folder to input files folder
    '
    
    Dim wsPath, destPath As String
    Dim splitWsPath() As String
    Dim fileName As String
    Dim listBoxSelectedCount As Integer
    Dim listBoxSelectedArr() As String
    Dim listBoxCount, destListCount As Integer
    Dim fileManipulation As String
    Dim FSO As Object
    Dim i, d, c As Integer
    Dim strNameItem As String
    Dim strNameItemSplit() As String
    Dim listText As String
    Const unSel = "Unselected"
    
    'Error handling
    Let wsPath = workspacePath.Caption
    If wsPath = unSel Then
        MsgBox "Workspace is not selected"
        Exit Sub
    End If
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    splitWsPath = Split(wsPath, " ")
    wsPath = splitWsPath(2) + "Correct_Files"
    destPath = splitWsPath(2) + "Input_Files"
    listBoxSelectedCount = listBoxCorrectFiles.ListCount
    destListCount = listBoxInputFiles.ListCount
    
    'changing location action (C to I)
    listBoxCount = 0
    For i = 0 To listBoxSelectedCount
        On Error GoTo outOfLoop
        If listBoxCorrectFiles.Selected(i) = True Then
            listBoxCount = listBoxCount + 1
            ReDim Preserve listBoxSelectedArr(1 To listBoxCount)
            fileManipulation = Split(listBoxCorrectFiles.List(i), ".")(1) _
            + "." + Split(listBoxCorrectFiles.List(i), ".")(2)
            fileManipulation = Right(fileManipulation, Len(fileManipulation) - 1)
            
            FSO.MoveFile wsPath & "\" & fileManipulation, destPath & "\" & fileManipulation
            
            listBoxSelectedArr(listBoxCount) = fileManipulation
        End If
    Next i
    
outOfLoop:
    Call reloadData_Click
    
    If listBoxCount = 0 Then Exit Sub
    For d = 0 To destListCount + listBoxCount - 1
        strNameItem = listBoxInputFiles.List(d)
        strNameItemSplit = Split(strNameItem, ". ")
        strNameItem = strNameItemSplit(1)
        For c = 0 To UBound(listBoxSelectedArr) - 1
            On Error GoTo outOfSecLoop
            listText = listBoxSelectedArr(c + 1)
            If strNameItem = listText Then
                listBoxInputFiles.Selected(d) = True
            End If
        Next c
    Next d
    
outOfSecLoop:
    
End Sub

'G3, Subgroup 4 - MF
Private Sub moveFilestToBin_Click()
    '
    'button is moveing selected files to bin folder
    '
    
    Dim wsPath As String
    Dim wsPathInp, wsPathCorr As String
    Dim destPath As String
    Dim splitWsPath() As String
    Dim fileName As String
    Dim lbSelectedInp, lbSelectedCorr As Integer
    Dim listBoxSelectedArr() As String
    Dim listBoxCount As Integer
    Dim fileManipulation As String
    Dim FSO As Object
    Dim iI, iC As Integer
    Const unSel = "Unselected"
    
    'Error handling
    Let wsPath = workspacePath.Caption
    If wsPath = unSel Then
        MsgBox "Workspace is not selected"
        Exit Sub
    End If
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    splitWsPath = Split(wsPath, " ")
    wsPathInp = splitWsPath(2) + "Input_Files"
    wsPathCorr = splitWsPath(2) + "Correct_Files"
    destPath = splitWsPath(2) + "Deleted_Files"
    lbSelectedInp = listBoxInputFiles.ListCount
    lbSelectedCorr = listBoxCorrectFiles.ListCount
    
    'Move input files to bin
    If lbSelectedInp > 0 Then
        Let fileManipulation = ""
        For iI = 0 To lbSelectedInp
            On Error Resume Next
            If listBoxInputFiles.Selected(iI) = True Then
                fileManipulation = Split(listBoxInputFiles.List(iI), ".")(1) _
                + "." + Split(listBoxInputFiles.List(iI), ".")(2)
                fileManipulation = Right(fileManipulation, Len(fileManipulation) - 1)
                FSO.MoveFile wsPathInp & "\" & fileManipulation, destPath & "\" & fileManipulation
            End If
        Next iI
    End If
    
    'Move correct files to bin
    If lbSelectedCorr > 0 Then
        Let fileManipulation = ""
        For iC = 0 To lbSelectedCorr
            On Error GoTo outOfLoop
            If listBoxCorrectFiles.Selected(iC) = True Then
                fileManipulation = Split(listBoxCorrectFiles.List(iC), ".")(1) _
                + "." + Split(listBoxCorrectFiles.List(iC), ".")(2)
                fileManipulation = Right(fileManipulation, Len(fileManipulation) - 1)
                FSO.MoveFile wsPathCorr & "\" & fileManipulation, destPath & "\" & fileManipulation
            End If
        Next iC
    End If
    
outOfLoop:
    
    Call reloadData_Click
    
End Sub

'----------------------
'Gorup 4 - Process Info
'----------------------

'G4, Subgroup 1 - PI
Public Sub emptyDirWarningCurr(label, boolFile, quantity)
    '
    'showing in labels files status
    '
    
    Const fileExist = "in directory"
    Const filesNotExist = "- directory is empty"
    
    'Add text to input Status
    If label = 0 Then
        If boolFile = True Then
            statusInpLabel.Caption = "- " + CStr(quantity) + " " + fileExist
        Else
            statusInpLabel.Caption = filesNotExist
        End If
    End If
    
    'Add text to output Status
    If label = 1 Then
        If boolFile = True Then
            statusCorrLabel.Caption = "- " + CStr(quantity) + " " + fileExist
        Else
            statusCorrLabel.Caption = filesNotExist
        End If
    End If

End Sub

'G4, Subgroup 2 - PI
Private Sub processBarActivity()
    
    Dim iStart As Integer
    Dim filesCount As Integer
    Dim pctDone As Single
    Dim lblWith As Integer
    
    
End Sub

'----------------------
'Gorup 5 - Service Code
'----------------------

'G5, Subgroup 1 - SC
Private Sub inputFilesPopulate(wsPath, fileArrInp, wrongDir)
    '
    'populates list of input files
    '
    
    Dim fileObj As Object
    Dim folderPickerObj As Object
    Dim filePicker As Object
    Dim pathArrayInp() As String
    Dim fileCounter As Integer
    Const folderInp = "Input_Files"
    Const typeLabel = 0
    Dim fileFill As Boolean: fileFill = False
    
    On Error GoTo wrongDirHandler
    Set folderPickerObj = CreateObject("scripting.filesystemobject")
    Set filePicker = folderPickerObj.getfolder(wsPath & folderInp).Files
    If filePicker.Count > 0 Then fileFill = True
    Let fileCounter = 0
    
    'collect input files to array
    If fileFill = True Then
    For Each fileObj In filePicker
        ReDim Preserve fileArrInp(fileCounter)
        ReDim Preserve pathArrayInp(fileCounter)
        fileArrInp(fileCounter) = CStr(fileCounter + 1) & ". " & fileObj.Name
        pathArrayInp(fileCounter) = fileObj.path
        fileCounter = fileCounter + 1
    Next fileObj
    End If
    
    Call emptyDirWarningCurr(typeLabel, fileFill, fileCounter) 'populates label
    
    'adding input array to list
    If fileFill = True Then
        listBoxInputFiles.List = fileArrInp
    Else
        listBoxInputFiles.Clear
    End If

Exit Sub
wrongDirHandler:
    If fileCounter = 0 Then wrongDir = True

End Sub

'G5, Subgroup 1 - SC
Private Sub correctFilesPopulate(wsPath, fileArrCorr)
    '
    'populates list of correct files
    '
    
    Dim fileObj As Object
    Dim folderPickerObj As Object
    Dim filePicker As Object
    Dim pathArrCorr() As String
    Dim fileCounter As Integer
    Const folderCorr = "Correct_Files"
    Const typeLabel = 1
    Dim fileFill As Boolean: fileFill = False
    
    Set folderPickerObj = CreateObject("scripting.filesystemobject")
    Set filePicker = folderPickerObj.getfolder(wsPath & folderCorr).Files
    If filePicker.Count > 0 Then fileFill = True
    Let fileCounter = 0
    
    'collect correct files to array
    If fileFill = True Then
    For Each fileObj In filePicker
        ReDim Preserve fileArrCorr(fileCounter)
        ReDim Preserve pathArrCorr(fileCounter)
        fileArrCorr(fileCounter) = CStr(fileCounter + 1) & ". " & fileObj.Name
        pathArrCorr(fileCounter) = fileObj.path
        fileCounter = fileCounter + 1
    Next fileObj
    End If
    
    Call emptyDirWarningCurr(typeLabel, fileFill, fileCounter) 'populates label
    
    'adding correct array to list
    If fileFill = True Then
        listBoxCorrectFiles.List = fileArrCorr
    Else
        listBoxCorrectFiles.Clear
    End If
    
End Sub

'G5, Subgroup 2 - SC
Private Sub listOfSelections_MouseDown(ByVal Button As Integer, _
ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
End Sub
