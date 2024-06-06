Attribute VB_Name = "Module5"
Sub CopySpecificColumns()
    'Declare Variables
    Dim PackageCol As Range, CompCol As Range
    Dim BVWorkbook As Workbook, EEIWorkbook As Workbook, targetWorkbook As Workbook
    Dim BVval As Range, EEIval As Range, Compval As Range
    Dim Discip() As Variant, SrcColNum() As Variant
    Dim i As Integer
    
    Application.ScreenUpdating = False
    
    ' Declare Arrays
    Discip() = Array("Mechanical", "Electrical", "Instrument")
    SrcColNum() = Array("V:Y", "Z:AC", "AD:AG")
    
    ' Set Target workbook for comparison
    Set targetWorkbook = ThisWorkbook
    
    ' Open the BV workbook
    Set BVWorkbook = Workbooks.Open(targetWorkbook.Sheets("Comparison Generator").Cells(4, 3))
 
    ' Open the EEI workbook
    Set EEIWorkbook = Workbooks.Open(targetWorkbook.Sheets("Comparison Generator").Cells(6, 3))
    
    ' Generate Data
    i = 0
    Do While (i < 3)
    
        'Clear Previous Content
        targetWorkbook.Sheets(Discip(i)).AutoFilter.ShowAllData
        targetWorkbook.Sheets(Discip(i)).Cells.Clear
        
        ' Copy BV Packages to Comparison Sheet
        Set PackageCol = BVWorkbook.Worksheets("TOP Log").Columns("A:E")
        Set CompCol = targetWorkbook.Worksheets(Discip(i)).Columns("A:E")
        PackageCol.Copy Destination:=CompCol
        PackageCol.Cells.EntireColumn.AutoFit
        
        ' Copy BV values to Comparison Sheet
        Set BVval = BVWorkbook.Worksheets("TOP Log").Range(SrcColNum(i))
        Set Compval = targetWorkbook.Worksheets(Discip(i)).Range("F:I")
        BVval.Copy Destination:=Compval
        BVval.Cells.EntireColumn.AutoFit
        BVval.Range("A3:C389").Copy
        Compval.Range("A3:C389").PasteSpecial xlPasteValuesAndNumberFormats
        
        ' Copy EEI values to Comparison Sheet
        Set EEIval = EEIWorkbook.Worksheets("TOP Log").Range(SrcColNum(i))
        Set Compval = targetWorkbook.Worksheets(Discip(i)).Range("J:M")
        EEIval.Copy Destination:=Compval
        EEIval.Cells.EntireColumn.AutoFit
        EEIval.Range("A3:C389").Copy
        Compval.Range("A3:C389").PasteSpecial xlPasteValuesAndNumberFormats
        Compval.Range("A3:C389").Interior.ColorIndex = xlNone
        
        Dim condFormat As FormatCondition
        
        targetWorkbook.Sheets(Discip(i)).Range("J1:M2").FormatConditions.Delete
        Set condFormat = targetWorkbook.Sheets(Discip(i)).Range("J1:M2").FormatConditions.Add(Type:=xlCellValue, Operator:=xlNotEqual, Formula1:="=L3")
        With condFormat
            .Interior.Color = RGB(0, 176, 240)
        End With
        
        
        ' Sort which are below 100% sa BV
        targetWorkbook.Sheets(Discip(i)).Activate
        targetWorkbook.Sheets(Discip(i)).Range("A2:I389").AutoFilter Field:=9, Criteria1:="<1"
        
        ' Highlight differences in remaining items
        targetWorkbook.Sheets(Discip(i)).Range("L3:L389").FormatConditions.Delete
        Set condFormat = targetWorkbook.Sheets(Discip(i)).Range("L3:L389").FormatConditions.Add(Type:=xlCellValue, Operator:=xlNotEqual, Formula1:="=H3")
        With condFormat
            .Interior.Color = RGB(202, 237, 251)
        End With
        
        targetWorkbook.Sheets(Discip(i)).Range("H3:H389").FormatConditions.Delete
        Set condFormat = targetWorkbook.Sheets(Discip(i)).Range("H3:H389").FormatConditions.Add(Type:=xlCellValue, Operator:=xlNotEqual, Formula1:="=L3")
        With condFormat
            .Interior.Color = RGB(251, 226, 213)
        End With
        
        ' Next Discipline
        i = i + 1
    Loop

MsgBox "Process completed!"
 
End Sub
