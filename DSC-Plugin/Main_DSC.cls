VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Main_DSC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub ApplicationSuppress()
    'freeze excel screen
    With Application
        .StatusBar = "Wait..."
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .DisplayAlerts = False
    End With
    
End Sub

Public Sub ApplicationRestore()
    'refresh excel screen
    With Application
        .StatusBar = "Done"
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
        .DisplayAlerts = True
        .CutCopyMode = False
        .StatusBar = ""
    End With
    
End Sub


Public Sub VarTypeCollection(key As String, value As String, vtColl As collection)
    '
    'Dict - Values collection
    '

    vtColl.Add "Empty", "0"
    vtColl.Add "Null", "1"
    vtColl.Add "Integer", "2"
    vtColl.Add "Long integer", "3"
    vtColl.Add "Single-precision floating-point number", "4"
    vtColl.Add "Double-precision floating-point number", "5"
    vtColl.Add "Currency value", "6"
    vtColl.Add "Date value", "7"
    vtColl.Add "String", "8"
    vtColl.Add "Object", "9"
    vtColl.Add "Error value", "10"
    vtColl.Add "Boolean value", "11"
    vtColl.Add "Variant", "12"
    vtColl.Add "A data access object", "13"
    vtColl.Add "Decimal value", "14"
    vtColl.Add "Byte value", "17"
    vtColl.Add "LongLong integer (valid on 64-bit platforms only)", "20"
    vtColl.Add "Variants that contain user-defined types", "36"
    vtColl.Add "Array", "8192"
    
    value = vtColl(key)
    
End Sub
