# Formats-editor_Print
Set print range and scan for [cash] word in column



Sub GEMReport()

Dim FileRange As Range
Dim Count As Integer 'Count variable
Dim StartRow As Integer 'Starting row number to place data
Dim r As Integer
Dim c As Integer
Dim SourceFile As String
Dim CurrentReportingDate As String
Dim path As Variant
    

Application.StatusBar = "Please be patient - clearing the old data"
Application.ScreenUpdating = False
Application.Calculation = xlAutomatic
      
    ThisWorkbook.Activate
    Worksheets("Print").Activate
    path = Cells(8, 5).Value & "\"
        
    Worksheets("Report_GEM").Activate
    Worksheets("Report_GEM").Copy
    ActiveWorkbook.SaveAs FileName:= _
        path & "Return Page GEM.xls", FileFormat:= _
        xlNormal, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False _
        , CreateBackup:=False
      
    Workbooks("Return Page GEM.xls").Activate
    Worksheets("Report_GEM").Activate
    Range("D9", Range("D9").Offset(32, 16)).Copy
    Cells(9, 4).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
           Application.CutCopyMode = False
           
    ThisWorkbook.Activate
    
    MsgBox "Copy Macro Done!"

      
Application.StatusBar = False
Application.ScreenUpdating = True
Application.Calculation = xlAutomatic
        
End Sub

