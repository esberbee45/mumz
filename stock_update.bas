Dim rule_remarks As String, en_item As Integer, da_item As Integer, su_changes As Integer, new_item As Integer, tbc_item As Integer, upload_item As Integer, parent_item As Integer
Dim oos As Integer, t_oos As Integer, in_stock As Integer, t_in_stock As Integer, update_item As Integer, vq_align As Integer, on_hold As Integer, oos_value As Integer, su_item As Integer
Dim vendor_name As String, path As String, date_now As String
Dim orig_wbk As Workbook
Dim new_wbk As Workbook

Sub Main()
    ' 1 - mandatory
    ' format exported csv from netsuite
    ' putting not on blank VSKU and UPC
    ' label items need to update
    ' "KSA-" items marked as 0? [Y/N]

    ' manual - mandatory
    ' do vlookup working file to vendor file
    ' do vlookup working file to monitoring file (backorder within 7 days - ns export)

    ' 2 - mandatory
    ' remove negative, check decimal, check non numeric

    ' 3 - optional
    ' apply the stock update rule [selection]

    ' 4 - optional
    ' low qty items marked as 0? [Y/N] - then create a file (lq)

    ' 5 - mandatory
    ' label the update (enable, disable, tbc, new, upload)
    ' count the label
    ' file creation (ns, nf, ns_da, mg_da, mg_da_parent, sc, nd, log)

    d = DateSerial(2021, 12, 12)
    path = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\Stock_Update\"
    date_now = Format(Now(), "dd-mmm-yy")
    vendor_name = ActiveSheet.Range("I2")

    If Now() >= d Then
        Exit Sub
    End If
    
    Dim try As Variant

    try = UCase(InputBox("Please select STOCK UPDATE RULE:" & _
        vbCrLf & vbTab & "1" & vbTab & "Format Extracted" & _
        vbCrLf & vbTab & "2" & vbTab & "Negative and Non Numeric" & _
        vbCrLf & vbTab & "3" & vbTab & "Stock Update Rule [Optional]" & _
        vbCrLf & vbTab & "4" & vbTab & "Low Qty [Optional]" & _
        vbCrLf & vbTab & "5" & vbTab & "Automation"))

    If try = vbNullString Then
        Exit Sub
    End If

    Select Case try
        Case "1"
            If ActiveSheet.Range("N1") <> "Available" Then
                MsgBox ("Incorrect Format")
                Exit Sub
            End If
            format_export
            blank_details
            label_item
            ksa_oos
            MsgBox ("Proceed with the VLOOKUP to vendor file and backorder file"), vbInformation
        Case "2"
            non_numeric
        Case "3"
            su_rule_select
        Case "4"
            lq_select
        Case "5"
            auto
            count_update
            file_creation
            
            'clear variables
            rule_remarks = vbNullString
            en_item = Empty
            da_item = Empty
            su_changes = Empty
            new_item = Empty
            tbc_item = Empty
            upload_item = Empty
            parent_item = Empty
            oos = Empty
            t_oos = Empty
            in_stock = Empty
            t_in_stock = Empty
            update_item = Empty
            vq_align = Empty
            on_hold = Empty
            oos_value = Empty
            su_item = Empty
            vendor_name = vbNullString
            path = vbNullString
            date_now = vbNullString
            
            MsgBox ("Stock Update Completed!"), vbInformation
        Case "ROP"
            rop_add
        Case "GT"
            gt
        Case "DR"
            pvt_daily_report
        Case "SIZE"
            copy_separate
        Case "MF"
            format_mf
    End Select
End Sub

Private Sub format_export()
    Application.ScreenUpdating = False
    ActiveWindow.Zoom = 70
    Cells.Select
    ActiveSheet.Range("1:1").AutoFilter
    
    Selection.Replace What:="- None -", Replacement:="", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    
    ActiveSheet.Range("N:N,AA:AA").Replace What:="0", Replacement:="", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    
    ActiveSheet.Range("H:H,J:J,K:K").NumberFormat = "0"

    With ActiveSheet 'resizing column
        .Range("A:D").ColumnWidth = 2.2
        .Range("E:F").ColumnWidth = 9
        .Range("G:G").ColumnWidth = 13
        .Range("H:H,J:K").ColumnWidth = 18
        .Range("L:L").ColumnWidth = 38
        .Range("S:S,AG:AH").ColumnWidth = 14
        .Range("H:K").HorizontalAlignment = xlLeft
    End With
        
    With ActiveSheet 'total VQ VS KSA
        .Range("AB:AB").Columns.Insert
        .Range("E:E").SpecialCells(2).Areas(1).Offset(0, 23) = "=AA1=N1"
    End With
        
    With ActiveSheet 'highlight discontinued type - red
        .Range("1:1").AutoFilter field:=2, Criteria1:=Array("Discontinued", "Temporary Disabled", "Dont Enable"), Operator:=xlFilterValues
        .AutoFilter.Range.Columns(2).Interior.ColorIndex = 3
    End With

    With ActiveSheet 'highlight discontinued type - orange
        .Range("1:1").AutoFilter field:=2, Criteria1:=Array("Discontinued With Stock", "Temporary Disabled With Stock"), Operator:=xlFilterValues
        .AutoFilter.Range.Columns(2).Interior.ColorIndex = 44
    End With

    With ActiveSheet 'highlight discontinued type - grey
        .Range("1:1").AutoFilter field:=2, Criteria1:=Array("Duplicate Product", "On Hold"), Operator:=xlFilterValues
        .AutoFilter.Range.Columns(2).Interior.ColorIndex = 3
    End With
    
    With ActiveSheet 'highlight discontinued type - green
        .Range("1:1").AutoFilter field:=2, Criteria1:="New But OOS"
        .AutoFilter.Range.Columns(2).Interior.ColorIndex = 43
    End With

    With ActiveSheet 'highlight discontinued type - purple
        .Range("1:1").AutoFilter field:=2, Criteria1:="New But No Details"
        .AutoFilter.Range.Columns(2).Interior.ColorIndex = 47
        .Range("1:1").AutoFilter field:=2
    End With

    With ActiveSheet 'highlight target store - KSA
        .Range("1:1").AutoFilter field:=3, Criteria1:="KSA"
        .Range("C:C").Select
    End With

    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    With ActiveSheet 'highlight target store - UAE
        .Range("1:1").AutoFilter field:=3, Criteria1:="UAE"
        .Range("C:C").Select
    End With

    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 15773696
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

    ActiveSheet.Range("1:1").AutoFilter field:=3
    
    With ActiveSheet 'highlight discontinued - red
        .Range("1:1").AutoFilter field:=4, Criteria1:="Yes"
        .AutoFilter.Range.Columns(4).Interior.ColorIndex = 3
        .Range("1:1").AutoFilter field:=4
    End With
    
    ActiveSheet.Range("F:F").Select 'highlight magento id - orange
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 49407
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    ActiveSheet.Range("K:K").Select 'highlight MW SKU - green
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 5296274
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
        
    ActiveSheet.Range("M:M").Select 'highlight VQ - yellow
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    With ActiveSheet 'highlight MW QTY - red KSA
        .Range("1:1").AutoFilter field:=14, Criteria1:="<>"
        .Range("1:1").AutoFilter field:=28, Criteria1:="TRUE"
        .AutoFilter.Range.Columns(13).Interior.Color = 255
        .Range("M:M,N:N,AA:AA").Select
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
        
    With ActiveSheet.Range("1:1")  'highlight MW QTY - green mix
        .AutoFilter field:=28, Criteria1:="FALSE"
        .AutoFilter field:=27, Criteria1:="<>"
        .Range("M:M,N:N,AA:AA").Select
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 5296274
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
        
    With ActiveSheet 'highlight MW QTY - blue UAE
        .Range("1:1").AutoFilter field:=27, Criteria1:=""
        .Range("M:M,N:N").Select
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 15773696
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
        
    With ActiveSheet.Range("1:1") 'unlfilter all
        .AutoFilter field:=27
        .AutoFilter field:=28
        .AutoFilter field:=14
    End With

    ActiveSheet.Range("M1").Select  'highlight header VQ and MQ
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

    ActiveSheet.Range("N1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 15773696
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

    With ActiveSheet 'delete column w/ formula & insert new plain column for working area
        .Range("AB:AB").EntireColumn.Delete
        .Range("N:P").Columns.Insert
        .Range("N:P").Select
    End With
    
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

    Sheets(1).Name = Format(Now(), "dd-mmm-yy") 'sheetname as current date
    Application.ScreenUpdating = True
End Sub

Private Sub ksa_oos()
    Dim a As Integer, b As Integer
    lrow = ActiveSheet.UsedRange.Rows.Count
    On Error Resume Next

    Application.ScreenUpdating = False
    With ActiveSheet 'count ksa skus vs total
        a = .Range("E2:E" & lrow).SpecialCells(xlCellTypeVisible).Count
        .Range("1:1").AutoFilter field:=11, Criteria1:="=*KSA-*"
        b = .Range("E2:E" & lrow).SpecialCells(xlCellTypeVisible).Count
    End With
    
    If b = 0 Then
        MsgBox ("No KSA SKU"), vbInformation
        ActiveSheet.Range("1:1").AutoFilter field:=11
        Exit Sub
    End If

    If MsgBox(b & " out of " & a & " are KSA SKUs. Do you want to mark it as OOS?", vbQuestion + vbYesNo) = vbYes Then
        With ActiveSheet
            .Range("E:E").SpecialCells(2).Areas(1).Offset(0, 9) = "0"
            .Range("E:E").SpecialCells(2).Areas(1).Offset(0, 11) = "KSA"
            .Range("1:1").AutoFilter field:=11
            .Range("N1:P1") = ""
        End With
    Else
        ActiveSheet.Range("1:1").AutoFilter field:=11
    End If

    Application.ScreenUpdating = True
End Sub

Function su_rule_price(price As Double, min_qty As Double, max_qty As Double) As Double
    Dim a As Integer, b As Integer
    lrow = ActiveSheet.UsedRange.Rows.Count
    On Error Resume Next

    Application.ScreenUpdating = False
    ActiveSheet.Range("L:L").Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

    With ActiveSheet 'count SKUs affected by stock update rule
        .Range("1:1").AutoFilter field:=20, Criteria1:="<" & price
        .Range("1:1").AutoFilter field:=14, Criteria1:=">0", Operator:=xlAnd, Criteria2:="<" & min_qty
        .AutoFilter.Range.Columns(12).Interior.ColorIndex = 6
        a = .Range("N2:N" & lrow).SpecialCells(xlCellTypeVisible).Count
    End With
    With ActiveSheet
        .Range("1:1").AutoFilter field:=14
        .Range("1:1").AutoFilter field:=20, Criteria1:=">=" & price
        .Range("1:1").AutoFilter field:=14, Criteria1:=">0", Operator:=xlAnd, Criteria2:="<" & max_qty
        .AutoFilter.Range.Columns(12).Interior.ColorIndex = 6
        b = .Range("N2:N" & lrow).SpecialCells(xlCellTypeVisible).Count
        .Range("1:1").AutoFilter field:=14
        .Range("1:1").AutoFilter field:=20
    End With
    
    If a + b = 0 Then
        MsgBox ("All QTY are passed on Stock Update Rule"), vbInformation
        Exit Function
    End If

    If MsgBox(a + b & " SKUs didn't passed the required QTY on " & rule_remarks & " rule. Do you want to marked it as OOS?", vbQuestion + vbYesNo) = vbYes Then
        With ActiveSheet 'marked as 0
            .Range("1:1").AutoFilter field:=12, Criteria1:=vbYellow, Operator:=xlFilterCellColor
            .Range("E:E").SpecialCells(2).Areas(1).Offset(0, 9) = "0"
            .Range("E:E").SpecialCells(2).Areas(1).Offset(0, 10) = "SU Rule - " & rule_remarks
            .Range("1:1").AutoFilter field:=12
            .Range("N1:P1") = ""
        End With
    Else
        ActiveSheet.Range("1:1").AutoFilter field:=12
        Exit Function
    End If

    Application.ScreenUpdating = True
End Function

Function su_rule(qty As Double) As Double
    Dim a As Integer
    lrow = ActiveSheet.UsedRange.Rows.Count
    On Error Resume Next

    Application.ScreenUpdating = False
    ActiveSheet.Range("L:L").Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

    With ActiveSheet 'count SKUs affected by stock update rule
        .Range("1:1").AutoFilter field:=14, Criteria1:=">0", Operator:=xlAnd, Criteria2:="<" & qty
        .AutoFilter.Range.Columns(12).Interior.ColorIndex = 6
        a = .Range("N2:N" & lrow).SpecialCells(xlCellTypeVisible).Count
        .Range("1:1").AutoFilter field:=14
    End With

    If a = 0 Then
        MsgBox ("All QTY are passed on Stock Update Rule"), vbInformation
        Exit Function
    End If
    
    If MsgBox(a & " SKUs didn't passed the required QTY on " & rule_remarks & " rule. Do you want to marked it as OOS?", vbQuestion + vbYesNo) = vbYes Then
        With ActiveSheet 'marked as 0
            .Range("1:1").AutoFilter field:=12, Criteria1:=vbYellow, Operator:=xlFilterCellColor
            .Range("E:E").SpecialCells(2).Areas(1).Offset(0, 9) = "0"
            .Range("E:E").SpecialCells(2).Areas(1).Offset(0, 10) = "SU Rule - " & rule_remarks
            .Range("1:1").AutoFilter field:=12
            .Range("N1:P1") = ""
        End With
    Else
        ActiveSheet.Range("1:1").AutoFilter field:=12
        Exit Function
    End If

    Application.ScreenUpdating = True
End Function

Function low_qty(qty As Double) As Double
    Dim a As Integer
    lrow = ActiveSheet.UsedRange.Rows.Count
    On Error Resume Next

    Application.ScreenUpdating = False
    With ActiveSheet 'count SKUs with low qty
        .Range("1:1").AutoFilter field:=14, Criteria1:=">0", Operator:=xlAnd, Criteria2:="<" & qty
        a = .Range("N2:N" & lrow).SpecialCells(xlCellTypeVisible).Count
    End With

    If a = 0 Then
        MsgBox ("No Low Qty Items"), vbInformation
        ActiveSheet.Range("1:1").AutoFilter field:=14
        Exit Function
    End If

    With ActiveSheet 'label low qty items
        .Range("E:E").SpecialCells(2).Areas(1).Offset(0, 10) = "Low Qty"
        .AutoFilter.Range.Columns(7).Interior.ColorIndex = 6
        .Range("O1") = ""
    End With

    low_qty_template
    
    If MsgBox("Do you want to mark Low Qty Items as OOS?", vbQuestion + vbYesNo) = vbYes Then
        With ActiveSheet
            .Range("E:E").SpecialCells(2).Areas(1).Offset(0, 9) = "0"
            .Range("1:1").AutoFilter field:=14
            .Range("N1") = ""
        End With
    Else
        ActiveSheet.Range("1:1").AutoFilter field:=7, Operator:=xlFilterNoFill
    End If

    Application.ScreenUpdating = True
End Function

Private Sub low_qty_template()
    Set orig_wbk = ActiveWorkbook
    Application.ScreenUpdating = False
    ActiveSheet.AutoFilter.Range.Columns("E:AC").Copy
    Set new_wbk = Workbooks.Add(xlWBATWorksheet)
    Sheets(1).Name = "Low Qty Items"

    With ActiveSheet
        .Range("A1").PasteSpecial Paste:=xlPasteValues
        .Range("P:X").EntireColumn.Delete
        .Range("K:N").EntireColumn.Delete
        .Range("I:I").EntireColumn.Delete
        .Range("B:F").EntireColumn.Delete
        .Range("D1") = "VQ"
        .Range("E1") = "Cost"
        .Range("1:1").Font.Bold = True
        .Range("A:A").SpecialCells(2).Offset(0, 1).Borders.LineStyle = xlContinuous
        .Range("A:A").SpecialCells(2).Offset(0, 2).Borders.LineStyle = xlContinuous
        .Range("A:A").SpecialCells(2).Offset(0, 3).Borders.LineStyle = xlContinuous
        .Range("A:A").SpecialCells(2).Offset(0, 4).Borders.LineStyle = xlContinuous
        .Range("A:A").SpecialCells(2).Offset(0, 5).Borders.LineStyle = xlContinuous
        .Range("A:A").EntireColumn.Hidden = True
        .Range("B:F").Columns.AutoFit
    End With

    ActiveWorkbook.SaveAs Filename:=path & "LQ " & vendor_name & " " & date_now & ".xlsx", FileFormat:=xlOpenXMLWorkbook
    ActiveWorkbook.Close False
    orig_wbk.Activate
    Application.ScreenUpdating = True
End Sub

Private Sub su_rule_select()
    Dim i As Variant

    i = UCase(InputBox("Please select STOCK UPDATE RULE:" & _
        vbCrLf & vbTab & "1" & vbTab & "Top Vendor" & _
        vbCrLf & vbTab & "2" & vbTab & "Default" & _
        vbCrLf & vbTab & "3" & vbTab & "Apparel & Book" & _
        vbCrLf & vbTab & "4" & vbTab & "Consumables" & _
        vbCrLf & vbTab & "5" & vbTab & "Grand baby toys - Non Chicco"))

    If i = vbNullString Then
        Exit Sub
    End If
    
    rule_remarks = "- None -"

    Select Case i
        Case "1" 'top vendor
            rule_remarks = "Top Vendor"
            su_rule_price 1000, 6, 4
        Case "2" 'default
            rule_remarks = "Default"
            su_rule_price 500, 5, 3
        Case "3" 'apparel & book
            rule_remarks = "Apparel & Book"
            su_rule 3
        Case "4" 'consumables
            rule_remarks = "Consumable"
            su_rule 10
        Case "5" 'grand baby (non chicco)
            rule_remarks = "Special Case"
            ActiveSheet.Range("1:1").AutoFilter field:=7, Criteria1:="Step2"
            su_rule 3
            ActiveSheet.Range("1:1").AutoFilter field:=7, Criteria1:="<>Step2"
            su_rule_price 1000, 6, 4
            ActiveSheet.Range("1:1").AutoFilter field:=7
    End Select
End Sub

Private Sub lq_select()
    Dim j As Variant

    j = UCase(InputBox("Please select STOCK UPDATE RULE:" & _
        vbCrLf & vbTab & "1" & vbTab & "Low Qty - Default" & _
        vbCrLf & vbTab & "2" & vbTab & "Low Qty - Transmed"))

    If j = vbNullString Then
        Exit Sub
    End If

    Select Case j
        Case "1" 'lq default
            low_qty 50
        Case "2" 'lq transmed
            low_qty 100
    End Select
End Sub

Private Sub label_item()
    oos_value = 0
    lrow = ActiveSheet.UsedRange.Rows.Count
    On Error Resume Next

    Application.ScreenUpdating = False

    With ActiveSheet 'disc type to number
        .Range("1:1").AutoFilter field:=2, Criteria1:=Array("Outdated Details", "N/A", "New But OOS", "New But No Details", ""), Operator:=xlFilterValues
        .Range("E:E").SpecialCells(2).Areas(1).Offset(0, 9) = "1"
        .Range("1:1").AutoFilter field:=2, Criteria1:="VQ Aligned"
        .Range("E:E").SpecialCells(2).Areas(1).Offset(0, 9) = "2"
        .Range("1:1").AutoFilter field:=2
        .Range("1:1").AutoFilter field:=14, Criteria1:=""
        .Range("E:E").SpecialCells(2).Areas(1).Offset(0, 9) = "0"
        .Range("1:1").AutoFilter field:=14
    End With

    With ActiveSheet 'item type to number
        .Range("1:1").AutoFilter field:=23, Criteria1:=Array("Simple", "Config Child", "Asst Child", "Simple Bundle", ""), Operator:=xlFilterValues
        .Range("E:E").SpecialCells(2).Areas(1).Offset(0, 10) = "1"
        'consignment@d3
        .Range("1:1").AutoFilter field:=18, Criteria1:="Consignment @ D3"
        .Range("E:E").SpecialCells(2).Areas(1).Offset(0, 10) = "2"
        .Range("1:1").AutoFilter field:=18
        
        .Range("1:1").AutoFilter field:=23, Criteria1:="FOC"
        .Range("E:E").SpecialCells(2).Areas(1).Offset(0, 10) = "2"
        .Range("1:1").AutoFilter field:=23
        .Range("1:1").AutoFilter field:=15, Criteria1:=""
        .Range("E:E").SpecialCells(2).Areas(1).Offset(0, 10) = "0"
        .Range("1:1").AutoFilter field:=15
    End With

    With ActiveSheet 'number to label
        .Range("E:E").SpecialCells(2).Areas(1).Offset(0, 11) = "=IFS(AND(N1=1,O1=1),""UPDATE"",OR(N1=0,O1=0),""OOS"",OR(N1=2,O1=2),""NOTE"")"
        .Range("P:P").Copy
        .Range("P:P").PasteSpecial xlPasteValues
        .Range("N:O") = ""
        .Range("1:1").AutoFilter field:=16, Criteria1:="OOS"
        .Range("E:E").SpecialCells(2).Areas(1).Offset(0, 9) = "0"
        .Range("1:1").AutoFilter field:=16, Criteria1:="NOTE"
        .Range("1:1").AutoFilter field:=2, Criteria1:="VQ Aligned"
        .Range("E:E").SpecialCells(2).Areas(1).Offset(0, 9) = "VQ Aligned"
        .Range("1:1").AutoFilter field:=2
        .Range("1:1").AutoFilter field:=23, Criteria1:="FOC"
        .Range("E:E").SpecialCells(2).Areas(1).Offset(0, 9) = "FOC"
        .Range("1:1").AutoFilter field:=23
        .Range("1:1").AutoFilter field:=16, Criteria1:="OOS"
        .Range("1:1").AutoFilter field:=13, Criteria1:="<>0", Operator:=xlAnd, Criteria2:="<>"
        oos_value = .Range("E2:E" & lrow).SpecialCells(xlCellTypeVisible).Count
        .Range("N1:P1") = ""
    End With

    If oos_value > 0 Then
        MsgBox ("There are values on OOS"), vbCritical
    Else
        With ActiveSheet
            .Range("1:1").AutoFilter field:=13
            .Range("1:1").AutoFilter field:=16, Criteria1:="UPDATE"
        End With
    End If

    With Application
        .CutCopyMode = False
        .ScreenUpdating = True
    End With
End Sub

Private Sub gt() 'consolidate tab on golden toys
    Application.ScreenUpdating = False

    For a = 1 To ActiveWorkbook.Sheets.Count
        Worksheets(a).Activate
        Worksheets(a).Range("1:2").EntireRow.Delete
        With ActiveSheet
            .Range("A:A").SpecialCells(2).Areas(1).Offset(0, 4) = "=ROUND(D1,0)"
            .Range("A:A").SpecialCells(2).Areas(1).Offset(0, 5) = Worksheets(a).Name
            .Range("E:E").Copy
            Range("E1").PasteSpecial xlPasteValues
        End With
    Next a

    Worksheets(1).Activate
    Set wrk = ActiveWorkbook 'Working in active workbook
        
    For Each sht In wrk.Worksheets
        If sht.Name = "Master" Then
            MsgBox "There is a worksheet called as 'Master'." & vbCrLf & _
                    "Please remove or rename this worksheet since 'Master' would be" & _
                    "the name of the result worksheet of this process.", vbOKOnly + vbExclamation, "Error"
            Exit Sub
        End If
    Next sht

    Set trg = wrk.Worksheets.Add(after:=wrk.Worksheets(wrk.Worksheets.Count)) 'Add new worksheet as the last worksheet
    trg.Name = "Master" 'Rename the new worksheet
    Set sht = wrk.Worksheets(1)
    colCount = sht.Cells(1, 255).End(xlToLeft).column 'Column count first
    
    With trg.Cells(1, 1).Resize(1, colCount) 'Now retrieve headers, no copy&paste needed
        .Value = sht.Cells(1, 1).Resize(1, colCount).Value
        .Font.Bold = True
    End With
        
    For Each sht In wrk.Worksheets 'If worksheet in loop is the last one, stop execution (it is Master worksheet)
        If sht.Index = wrk.Worksheets.Count Then
            Exit For
        End If
        'Data range in worksheet - starts from second row as first rows are the header rows in all worksheets
        Set rng = sht.Range(sht.Cells(2, 1), sht.Cells(65536, 1).End(xlUp).Resize(, colCount))
        'Put data into the Master worksheet
        trg.Cells(65536, 1).End(xlUp).Offset(1).Resize(rng.Rows.Count, rng.Columns.Count).Value = rng.Value
    Next sht

    trg.Columns.AutoFit 'Fit the columns in Master worksheet
        
    With ActiveSheet 'labeling headers
        .Range("E1") = "Aval. Qty"
        .Range("F1") = "Category"
        .Range("C:C").NumberFormat = "0"
        .Range("D:D").EntireColumn.Delete
    End With

    Application.ScreenUpdating = True
End Sub

Private Sub blank_details()
    Application.ScreenUpdating = False

    With ActiveSheet 'blank upc to NO UPC
        .Range("1:1").AutoFilter field:=8, Criteria1:=""
        .Range("E:E").SpecialCells(2).Areas(1).Offset(0, 3) = "NO UPC"
        .AutoFilter.Range.Columns(8).Interior.ColorIndex = 3
        .Range("1:1").AutoFilter field:=8
        .Range("H1") = "UPC Code"
    End With

    With ActiveSheet 'blank upc to NO UPC
        .Range("1:1").AutoFilter field:=10, Criteria1:=""
        .Range("E:E").SpecialCells(2).Areas(1).Offset(0, 5) = "NO VSKU"
        .AutoFilter.Range.Columns(10).Interior.ColorIndex = 3
        .Range("1:1").AutoFilter field:=10
        .Range("J1") = "Vendor SKU"
    End With

    Application.ScreenUpdating = True
End Sub

Private Sub auto()
    Application.ScreenUpdating = False

    With ActiveSheet 'label enable item
        .Range("1:1").AutoFilter field:=16, Criteria1:="UPDATE" 'only stock update items
        .Range("1:1").AutoFilter field:=14, Criteria1:="<>0", Operator:=xlAnd, Criteria2:="<>#N/A"
        .Range("1:1").AutoFilter field:=2, Criteria1:="New But OOS"
        .Range("E:E").SpecialCells(2).Areas(1).Offset(0, 10) = "ENABLE"
        .Range("1:1").AutoFilter field:=2
    End With
        
    With ActiveSheet 'label disable item
        .Range("1:1").AutoFilter field:=14, Criteria1:="DISC"
        .Range("E:E").SpecialCells(2).Areas(1).Offset(0, 9) = "0"
        .Range("E:E").SpecialCells(2).Areas(1).Offset(0, 10) = "DISABLE"
    End With

    With ActiveSheet 'label #N/A item
        .Range("1:1").AutoFilter field:=14, Criteria1:="#N/A"
        .Range("1:1").AutoFilter field:=22, Criteria1:=">=" & CDbl(Now() - 31)
        .Range("E:E").SpecialCells(2).Areas(1).Offset(0, 9) = "NEW"
        .Range("E:E").SpecialCells(2).Areas(1).Offset(0, 10) = "NEW"
        .Range("1:1").AutoFilter field:=22
        .Range("E:E").SpecialCells(2).Areas(1).Offset(0, 9) = 0
        .Range("E:E").SpecialCells(2).Areas(1).Offset(0, 10) = "TBC"
        .Range("1:1").AutoFilter field:=14
    End With

    With ActiveSheet 'label for upload item
        .Range("1:1").AutoFilter field:=2, Criteria1:="New But No Details"
        .Range("E:E").SpecialCells(2).Areas(1).Offset(0, 10) = "UPLOAD"
        .Range("1:1").AutoFilter field:=2
        .Range("1:1").AutoFilter field:=6, Criteria1:=""
        .Range("E:E").SpecialCells(2).Areas(1).Offset(0, 10) = "UPLOAD"
        .Range("1:1").AutoFilter field:=6
        .Range("N1:O1") = ""
    End With

    Application.ScreenUpdating = True
End Sub

Private Sub count_update()
    oos = 0
    in_stock = 0
    new_item = 0
    t_oos = 0
    t_in_stock = 0
    tbc_item = 0
    en_item = 0
    da_item = 0
    upload_item = 0
    su_changes = 0
    vq_align = 0
    on_hold = 0
    parent_item = 0
    su_item = 0
    lrow = ActiveSheet.UsedRange.Rows.Count
    On Error Resume Next

    Application.ScreenUpdating = False

    With ActiveSheet 'count overall update
        .Range("1:1").AutoFilter field:=16, Criteria1:="UPDATE"
        update_item = .Range("E2:E" & lrow).SpecialCells(xlCellTypeVisible).Count
    End With
    
    With ActiveSheet 'highlight OOS - red
        .Range("1:1").AutoFilter field:=13, Criteria1:=">0"
        .Range("1:1").AutoFilter field:=14, Criteria1:="<1"
        .AutoFilter.Range.Columns(5).Interior.ColorIndex = 3
        .AutoFilter.Range.Columns(1).Interior.ColorIndex = 6
        .AutoFilter.Range.Columns(9).Interior.ColorIndex = 6
        oos = .Range("E2:E" & lrow).SpecialCells(xlCellTypeVisible).Count
    End With

    With ActiveSheet 'highlight back in stock - yellow
        .Range("1:1").AutoFilter field:=13, Criteria1:="<1", Operator:=xlOr, Criteria2:=""
        .Range("1:1").AutoFilter field:=14, Criteria1:=">0"
        .AutoFilter.Range.Columns(5).Interior.ColorIndex = 6
        .AutoFilter.Range.Columns(1).Interior.ColorIndex = 6
        .AutoFilter.Range.Columns(9).Interior.ColorIndex = 6
        in_stock = .Range("E2:E" & lrow).SpecialCells(xlCellTypeVisible).Count
    End With

    With ActiveSheet 'highlight not 0 on new VQ and old VQ - yellow
        .Range("1:1").AutoFilter field:=13 'unfilter existing
        .Range("1:1").AutoFilter field:=14
        .Range("1:1").AutoFilter field:=5, Operator:=xlFilterNoFill
        .Range("1:1").AutoFilter field:=14, Criteria1:="<>0", Operator:=xlAnd, Criteria2:="<>NEW"
        .AutoFilter.Range.Columns(1).Interior.ColorIndex = 6
        .Range("1:1").AutoFilter field:=5
    End With

    With ActiveSheet 'count update
        .Range("1:1").AutoFilter field:=14, Criteria1:="NEW"
        new_item = .Range("E2:E" & lrow).SpecialCells(xlCellTypeVisible).Count
        .Range("1:1").AutoFilter field:=14, Criteria1:="0"
        t_oos = .Range("E2:E" & lrow).SpecialCells(xlCellTypeVisible).Count
        .Range("1:1").AutoFilter field:=14, Criteria1:="<>0"
        t_in_stock = .Range("E2:E" & lrow).SpecialCells(xlCellTypeVisible).Count
        .Range("1:1").AutoFilter field:=14
        .Range("1:1").AutoFilter field:=15, Criteria1:="TBC"
        tbc_item = .Range("E2:E" & lrow).SpecialCells(xlCellTypeVisible).Count
        .Range("1:1").AutoFilter field:=15, Criteria1:="ENABLE"
        en_item = .Range("E2:E" & lrow).SpecialCells(xlCellTypeVisible).Count
        .Range("1:1").AutoFilter field:=15, Criteria1:="DISABLE"
        da_item = .Range("E2:E" & lrow).SpecialCells(xlCellTypeVisible).Count
        .Range("1:1").AutoFilter field:=15, Criteria1:="*SU Rule*"
        su_item = .Range("E2:E" & lrow).SpecialCells(xlCellTypeVisible).Count
        .Range("1:1").AutoFilter field:=15, Criteria1:="UPLOAD"
        upload_item = .Range("E2:E" & lrow).SpecialCells(xlCellTypeVisible).Count
        .Range("1:1").AutoFilter field:=15, Criteria1:="ENABLE", Operator:=xlOr, Criteria2:="DISABLE"
        .Range("1:1").AutoFilter field:=23, Criteria1:="Config Child", Operator:=xlOr, Criteria2:=""
        .Range("1:1").AutoFilter field:=24, Criteria1:="<>", Operator:=xlAnd, Criteria2:="<>*ASST*"
        parent_item = .Range("E2:E" & lrow).SpecialCells(xlCellTypeVisible).Count
        .Range("1:1").AutoFilter field:=24
        .Range("1:1").AutoFilter field:=23
        .Range("1:1").AutoFilter field:=15
        .Range("1:1").AutoFilter field:=1, Criteria1:=vbYellow, Operator:=xlFilterCellColor
        su_changes = .Range("E2:E" & lrow).SpecialCells(xlCellTypeVisible).Count
        .Range("1:1").AutoFilter
        .Range("1:1").AutoFilter field:=2, Criteria1:="VQ Aligned"
        vq_align = .Range("E2:E" & lrow).SpecialCells(xlCellTypeVisible).Count
        .Range("1:1").AutoFilter field:=2, Criteria1:="On Hold"
        on_hold = .Range("E2:E" & lrow).SpecialCells(xlCellTypeVisible).Count
        .Range("1:1").AutoFilter field:=2
        .Range("1:1").AutoFilter field:=16, Criteria1:="UPDATE"
    End With

    Application.ScreenUpdating = True

    MsgBox ("OOS: " & oos & "   |   " & t_oos & " / " & update_item & vbNewLine & _
        "IN STOCK: " & in_stock & "   |   " & t_in_stock & " / " & update_item & vbNewLine & _
        "UPDATE SUMMARY: " & su_changes & " / " & update_item & vbNewLine & vbNewLine & _
        "NEW: " & new_item & vbNewLine & _
        "TBC: " & tbc_item & vbNewLine & _
        "FOR UPLOAD: " & upload_item & vbNewLine & _
        "ENABLE: " & en_item & vbNewLine & _
        "DISABLE: " & da_item & vbNewLine & _
        "VQ ALIGNED: " & vq_align & vbNewLine & _
        "OH HOLD: " & on_hold & vbNewLine & _
        "SU RULE: " & rule_remarks & vbNewLine & _
        "SU RULE SKUs: " & su_item), vbInformation
End Sub

Private Sub file_creation()
    Set orig_wbk = ActiveWorkbook

    Application.ScreenUpdating = False
    ActiveSheet.Range("1:1").AutoFilter field:=1, Criteria1:=vbYellow, Operator:=xlFilterCellColor

    If oos + in_stock = 0 And su_changes > 1000 Then 'nf only
        save_ns_csv "NS", " NF"
        ActiveSheet.Range("1:1").AutoFilter field:=9
    ElseIf oos + in_stock <> 0 And su_changes > 1000 Then 'ns and nf
        ActiveSheet.Range("1:1").AutoFilter field:=9, Criteria1:=vbYellow, Operator:=xlFilterCellColor
        save_ns_csv "NS", ""
        ActiveSheet.Range("1:1").AutoFilter field:=9, Operator:=xlFilterNoFill
        save_ns_csv "NS", " NF"
        ActiveSheet.Range("1:1").AutoFilter field:=9
    ElseIf oos + in_stock = 0 And (su_changes <= 1000 And su_changes > 0) Then 'nf only
        save_ns_csv "NS", " NF"
    ElseIf oos + in_stock <> 0 And (su_changes <= 1000 And su_changes > 0) Then 'ns only
        save_ns_csv "NS", ""
    End If

    ActiveSheet.Range("1:1").AutoFilter field:=1

    If tbc_item > 0 Or new_item > 0 Then 'create tbc template
        tbc_template
    End If

    If upload_item > 0 Then 'create upload template
        upload_template
    End If

    If en_item > 0 Or da_item > 0 Then 'create ns da template
        ns_da
        mg_da
    End If

    If parent_item > 0 Then 'create mg da parent
        create_pivot_parent
        mg_da_parent
    End If

    create_notepad
    log_file
    Application.ScreenUpdating = True
End Sub

Private Sub non_numeric()
    Dim a As Integer, b As Integer
    Application.ScreenUpdating = False

    With ActiveSheet 'remove negative
        .Range("1:1").AutoFilter field:=14, Criteria1:="<0"
        .Range("E:E").SpecialCells(2).Areas(1).Offset(0, 9) = "0"
        .Range("N1") = ""
        .Range("1:1").AutoFilter field:=14, Criteria1:="<>#N/A"
    End With

    With ActiveSheet
        .Range("E:E").SpecialCells(2).Areas(1).Offset(0, 10) = "=IF(N1=""DISC"",""TRUE"",ISNUMBER(N1))"
        .Range("O1") = ""
    End With

    If Not ActiveSheet.Range("O:O").Find("False", LookIn:=xlValues) Is Nothing Then 'find non numeric
        MsgBox ("There are NON NUMERIC on Column N"), vbCritical
        a = 1
    End If

    With ActiveSheet
        .Range("E:E").SpecialCells(2).Areas(1).Offset(0, 10) = "=INT(N1)=N1"
        .Range("O1") = ""
    End With

    If Not ActiveSheet.Range("O:O").Find("False", LookIn:=xlValues) Is Nothing Then 'find non whole number
        MsgBox ("There are NON WHOLE NUMBER on Column N"), vbCritical
        b = 1
    End If

    With ActiveSheet
        .Range("O:O").Clear
        .Range("1:1").AutoFilter field:=14
    End With

    If Not a > 0 Or b > 0 Then
        MsgBox ("All Qty are Good"), vbInformation
    End If

    Application.ScreenUpdating = True
End Sub

Private Sub upload_template()
    Application.ScreenUpdating = False
    ActiveSheet.Range("1:1").AutoFilter field:=15, Criteria1:="UPLOAD"
    ActiveSheet.AutoFilter.Range.Columns("B:O").Copy
    Set new_wbk = Workbooks.Add(xlWBATWorksheet)

    Sheets(1).Name = "For Upload"

    With ActiveSheet 'upload format
        .Range("A1").PasteSpecial Paste:=xlPasteValues
        .Range("M1") = "VQ"
        .Range("1:1").Font.Bold = True
        .Range("N:N").EntireColumn.Delete
        .Range("L:L").EntireColumn.Delete
        .Range("I:I").EntireColumn.Delete
        .Range("G:G").EntireColumn.Delete
        .Range("B:D").EntireColumn.Delete
        .Range("A:A").SpecialCells(2).Offset(0, 0).Borders.LineStyle = xlContinuous
        .Range("A:A").SpecialCells(2).Offset(0, 1).Borders.LineStyle = xlContinuous
        .Range("A:A").SpecialCells(2).Offset(0, 2).Borders.LineStyle = xlContinuous
        .Range("A:A").SpecialCells(2).Offset(0, 3).Borders.LineStyle = xlContinuous
        .Range("A:A").SpecialCells(2).Offset(0, 4).Borders.LineStyle = xlContinuous
        .Range("A:A").SpecialCells(2).Offset(0, 5).Borders.LineStyle = xlContinuous
        .Range("A:A").SpecialCells(2).Offset(0, 6).Borders.LineStyle = xlContinuous
        .Range("A:F").Columns.AutoFit
    End With

    ActiveWorkbook.SaveAs Filename:=path & "ND " & vendor_name & " " & date_now & ".xlsx", FileFormat:=xlOpenXMLWorkbook
    ActiveWorkbook.Close False
    orig_wbk.Activate
    Application.ScreenUpdating = True
End Sub

Private Sub save_ns_csv(prefix As String, suffix As String)
    Application.ScreenUpdating = False
    ActiveSheet.AutoFilter.Range.Columns("E:N").Copy
    Set new_wbk = Workbooks.Add(xlWBATWorksheet)

    With ActiveSheet
        .Range("A1").PasteSpecial Paste:=xlPasteValues
        .Range("B:I").EntireColumn.Delete
    End With

    ns_format
    
    ActiveWorkbook.SaveAs Filename:=path & prefix & " " & vendor_name & " " & date_now & suffix & ".csv", FileFormat:=xlCSVUTF8
    ActiveWorkbook.Close False
    orig_wbk.Activate
    Application.ScreenUpdating = True
End Sub

Private Sub tbc_template()
    Application.ScreenUpdating = False
    ActiveSheet.Range("1:1").AutoFilter field:=15, Criteria1:="NEW", Operator:=xlOr, Criteria2:="TBC"
    ActiveSheet.AutoFilter.Range.Columns("G:O").Copy
    Set new_wbk = Workbooks.Add(xlWBATWorksheet)

    Sheets(1).Name = "Status Confirmation"

    With ActiveSheet 'tbc format
        .Range("A1").PasteSpecial Paste:=xlPasteValues
        .Range("G:H").EntireColumn.Delete
        .Range("C:C").EntireColumn.Delete
        .Range("D:D").Cut
        .Range("A:A").Insert Shift:=xlToRight
        .Range("D:D").Cut
        .Range("C:C").Insert Shift:=xlToRight
        .Range("F:F").Select
        .Range("A:A").EntireColumn.Hidden = True
    End With

    Selection.Replace What:="NEW", Replacement:="Newly Uploaded", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False

    Selection.Replace What:="TBC", Replacement:="Not on Stock Report", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False

    With ActiveSheet.Range("G:G").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="Available,Discontinued,Out Of Stock"
    End With

    With ActiveSheet
        .Range("E1") = "Item Name"
        .Range("F1") = "Item Remarks"
        .Range("G1") = "Status Confirmation"
        .Range("H1") = "QTY (If Available)"
        .Range("1:1").Font.Bold = True
        .Range("A:A").SpecialCells(2).Offset(0, 1).Borders.LineStyle = xlContinuous
        .Range("A:A").SpecialCells(2).Offset(0, 2).Borders.LineStyle = xlContinuous
        .Range("A:A").SpecialCells(2).Offset(0, 3).Borders.LineStyle = xlContinuous
        .Range("A:A").SpecialCells(2).Offset(0, 4).Borders.LineStyle = xlContinuous
        .Range("A:A").SpecialCells(2).Offset(0, 5).Borders.LineStyle = xlContinuous
        .Range("A:A").SpecialCells(2).Offset(0, 6).Borders.LineStyle = xlContinuous
        .Range("A:A").SpecialCells(2).Offset(0, 7).Borders.LineStyle = xlContinuous
    End With

    ActiveSheet.Range("B:H").Columns.AutoFit
    ActiveWorkbook.SaveAs Filename:=path & "SC " & vendor_name & " " & date_now & ".xlsx", FileFormat:=xlOpenXMLWorkbook
    ActiveWorkbook.Close False
    orig_wbk.Activate
    Application.ScreenUpdating = True
End Sub

Private Sub log_file()
    Application.ScreenUpdating = False
    orig_wbk.Activate
    Sheets(1).Activate

    With ActiveSheet
        .Range("1:1").AutoFilter
        .Range("1:1").AutoFilter
        .Range("N:P").Copy
        .Range("N:P").PasteSpecial xlPasteValues
    End With

    ActiveWorkbook.SaveAs Filename:=path & "LOG " & vendor_name & " " & date_now & ".xlsx", FileFormat:=xlOpenXMLWorkbook
    Application.ScreenUpdating = True
    orig_wbk.Close False
End Sub

Private Sub ns_da()
    lrow = ActiveSheet.UsedRange.Rows.Count
    On Error Resume Next
    
    Application.ScreenUpdating = False
    ActiveSheet.Range("1:1").AutoFilter field:=15, Criteria1:="ENABLE", Operator:=xlOr, Criteria2:="DISABLE"
    ActiveSheet.AutoFilter.Range.Columns("E:Q").Copy
    Set new_wbk = Workbooks.Add(xlWBATWorksheet)

    With ActiveSheet
        .Range("A1").PasteSpecial Paste:=xlPasteValues
        .Range("B:J").EntireColumn.Delete
    End With
    
    ns_da_format

    ActiveWorkbook.SaveAs Filename:=path & "NS " & vendor_name & " " & date_now & " DA" & ".csv", FileFormat:=xlCSVUTF8
    ActiveWorkbook.Close False
    orig_wbk.Activate
    Application.ScreenUpdating = True
End Sub

Private Sub mg_da()
    Application.ScreenUpdating = False
    ActiveSheet.Range("1:1").AutoFilter field:=15, Criteria1:="ENABLE", Operator:=xlOr, Criteria2:="DISABLE"
    ActiveSheet.AutoFilter.Range.Columns("K:Q").Copy
    Set new_wbk = Workbooks.Add(xlWBATWorksheet)

    With ActiveSheet
        .Range("A1").PasteSpecial Paste:=xlPasteValues
        .Range("E:F").EntireColumn.Delete
        .Range("B:C").EntireColumn.Delete
    End With

    mg_da_format

    ActiveWorkbook.SaveAs Filename:=path & "MG " & vendor_name & " " & date_now & " DA" & ".csv", FileFormat:=xlCSVUTF8
    ActiveWorkbook.Close False
    orig_wbk.Activate
    Application.ScreenUpdating = True
End Sub

Private Sub ns_format()
    Application.ScreenUpdating = False

    With ActiveSheet
        .Range("A:A").SpecialCells(2).Offset(0, 2) = "=IF(B1=0, ""T"",""F"")"
        .Range("A:A").SpecialCells(2).Offset(0, 3) = "T"
        .Range("B1") = "Vendor Item Quantity"
        .Range("C1") = "Vendor Out Of Stock"
        .Range("D1") = "Stock Update"
    End With

    Application.ScreenUpdating = True
End Sub

Private Sub ns_da_format()
    Application.ScreenUpdating = False

    If da_item > 0 Then
        a = InputBox("Enter Freshdesk Ticket#")
    End If

    With ActiveSheet
        .Range("A:A").SpecialCells(2).Offset(0, 4) = "=IFS(B1=""ENABLE"",""N/A"",AND(B1=""DISABLE"",D1<>""""),""Discontinued With Stock"",AND(B1=""DISABLE"",D1=""""),""Discontinued"")"
        .Range("A:A").SpecialCells(2).Offset(0, 5) = "=IFS(E1=""N/A"",""F"",E1=""Discontinued With Stock"",""F"",E1=""Discontinued"",""T"")"
        .Range("1:1").AutoFilter field:=2, Criteria1:="DISABLE"
        .Range("A:A").SpecialCells(2).Offset(0, 6) = "DA - " & a
        .Range("1:1").AutoFilter field:=2
        .Range("E:G").Copy
        .Range("E:G").PasteSpecial xlPasteValues
        .Range("B:D").EntireColumn.Delete
        .Range("B1") = "Discontinued Type"
        .Range("C1") = "Item Disabled"
        .Range("D1") = "Comments"
    End With

    Application.ScreenUpdating = True
End Sub

Private Sub ns_qty_da_format()
    a = InputBox("Enter Freshdesk Ticket#")
    Application.ScreenUpdating = False

    With ActiveSheet
        .Range("A:A").SpecialCells(2).Areas(1).Offset(0, 5) = "=IF(B1=0, ""T"",""F"")"
        .Range("A:A").SpecialCells(2).Areas(1).Offset(0, 6) = "T"
        .Range("A:A").SpecialCells(2).Areas(1).Offset(0, 7) = "=IFS(C1=""ENABLE"",""N/A"",AND(C1=""DISABLE"",B1+E1=0),""Discontinued"",AND(C1=""DISABLE"",B1+E1>0),""Discontinued With Stock"")"
        .Range("A:A").SpecialCells(2).Areas(1).Offset(0, 8) = "=IFS(H1=""N/A"",""F"",H1=""Discontinued With Stock"",""F"",H1=""Discontinued"",""T"")"
        .Range("A:A").SpecialCells(2).Areas(1).Offset(0, 9) = "DA - " & a
        .Range("F:J").Copy
        .Range("F:J").PasteSpecial xlPasteValues
        .Range("B1") = "Vendor Item Quantity"
        .Range("F1") = "Vendor Out Of Stock"
        .Range("G1") = "Stock Update"
        .Range("H1") = "Discontinued Type"
        .Range("I1") = "Item Disabled"
        .Range("J1") = "Comments"
        .Range("C:E").EntireColumn.Delete
    End With

    Application.ScreenUpdating = True
End Sub

Private Sub mg_qty_da_format()
    Application.ScreenUpdating = False

    With ActiveSheet
        .Range("A:A").SpecialCells(2).Areas(1).Offset(0, 5) = "=B1+E1"
        .Range("A:A").SpecialCells(2).Areas(1).Offset(0, 6) = "=IF(F1=0, ""0"",""1"")"
        .Range("A:A").SpecialCells(2).Areas(1).Offset(0, 7) = "=IF(F1=0, ""Disabled"",""Enabled"")"
        .Range("F:H").Copy
        .Range("F:H").PasteSpecial xlPasteValues
        .Range("F1") = "qty"
        .Range("G1") = "is_in_stock"
        .Range("H1") = "status"
        .Range("B:E").EntireColumn.Delete
    End With

    Application.ScreenUpdating = True
End Sub

Private Sub mg_da_format()
    Application.ScreenUpdating = False

    With ActiveSheet
        .Range("A:A").SpecialCells(2).Offset(0, 3) = "=B1+C1"
        .Range("A:A").SpecialCells(2).Offset(0, 4) = "=IF(D1=0,""Disabled"",""Enabled"")"
        .Range("E:E").Copy
        .Range("E:E").PasteSpecial xlPasteValues
        .Range("B:D").EntireColumn.Delete
        .Range("A1") = "sku"
        .Range("B1") = "status"
    End With

    Application.ScreenUpdating = True
End Sub

Private Sub rop_add()
    Application.ScreenUpdating = False

    ActiveSheet.Range("A:A").SpecialCells(2).Offset(0, 3) = "B1<C1"

    If Not ActiveSheet.Range("D:D").Find("False", LookIn:=xlValues) Is Nothing Then 'SSL Days less than PSL Days
        MsgBox ("There are SSL Days > PSL Days"), vbCritical
    End If

    With ActiveSheet
        .Range("D:D").Clear
        .Range("1:1").EntireRow.Insert
        .Range("A1") = "Internal ID"
        .Range("B1") = "Safety Stock Level Days"
        .Range("C1") = "Preferred Stock Level Days"
    End With
    Application.ScreenUpdating = True
End Sub

Private Sub create_pivot_parent()
    Dim pvt As PivotTable
    Dim StartPvt As String

    Application.ScreenUpdating = False
    ActiveSheet.Range("1:1").AutoFilter
    Sheets(1).Copy after:=Sheets(1)
        
    With ActiveSheet
        .Name = "DB"
        .Range("R:R").Insert
   '     .Range("E:E").SpecialCells(2).Areas(1).Offset(0, 13) = "=N1+Q1"
        .Range("E:E").SpecialCells(2).Areas(1).Offset(0, 13) = "=IFERROR(N1+Q1,M1+Q1)"
        .Range("N1") = "VQ"
        .Range("O1") = "Note"
        .Range("P1") = "Label"
        .Range("R1") = "Total"
        .Range("N:R").Copy
        .Range("N:R").PasteSpecial xlPasteValues
    End With
    
    Sheets.Add after:=Sheets(2)
    ActiveSheet.Name = "PivotTable"
    StartPvt = ActiveSheet.Name & "!" & ActiveSheet.Range("A1").Address(ReferenceStyle:=xlR1C1)
    Set pvtCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:="DB!$A:$Y")
    Set pvt = pvtCache.CreatePivotTable(TableDestination:=StartPvt, TableName:="PivotTable1")
    
    pvt.PivotFields("Parent SKU").Orientation = xlRowField
    pvt.AddDataField pvt.PivotFields("Total"), "Sum of Child SKUs", xlSum
    Application.ScreenUpdating = True
End Sub

Private Sub mg_da_parent()
    On Error Resume Next
    Application.ScreenUpdating = False
    Sheets(1).Activate
    
    With ActiveSheet
        .Range("1:1").AutoFilter field:=15, Criteria1:="ENABLE", Operator:=xlOr, Criteria2:="DISABLE"
        .Range("1:1").AutoFilter field:=23, Criteria1:="Config Child", Operator:=xlOr, Criteria2:=""
        .Range("1:1").AutoFilter field:=24, Criteria1:="<>", Operator:=xlAnd, Criteria2:="<>*ASST*"
        .AutoFilter.Range.Columns("O:X").Copy
    End With

    Sheets.Add after:=Sheets(3)

    With ActiveSheet 'paste & remove duplicate
        .Range("A1").PasteSpecial Paste:=xlPasteValues
        .Range("B:I").EntireColumn.Delete
        .Range("A:B").RemoveDuplicates Columns:=Array(1, 2), Header:=xlNo
        .Range("B:B").Cut
        .Range("A:A").Insert Shift:=xlToRight
    End With

    With ActiveWorkbook
        .Names("parent_pivot").Delete
        .Names.Add Name:="parent_pivot", RefersTo:=Sheets("PivotTable").Range("A:B") 'declare pivot table
    End With

    With ActiveSheet
        .Range("A:A").SpecialCells(2).Offset(0, 2) = "=VLOOKUP(A1,parent_pivot,2,0)"
        .Range("A:A").SpecialCells(2).Offset(0, 3) = "1"
        .Range("A:A").SpecialCells(2).Offset(0, 4) = "=IFS(B1=""ENABLE"",""Enabled"",AND(B1=""DISABLE"",C1=0),""Disabled"",AND(B1=""DISABLE"",C1<>0),""Enabled"")"
        .Range("A:E").Copy
    End With

    Set new_wbk = Workbooks.Add(xlWBATWorksheet)

    With ActiveSheet
        .Range("A1").PasteSpecial xlPasteValues
        .Range("B:C").EntireColumn.Delete
        .Range("A1") = "sku"
        .Range("B1") = "is_in_stock"
        .Range("C1") = "status"
    End With

    ActiveWorkbook.SaveAs Filename:=path & "MG " & vendor_name & " " & date_now & " DA PARENT" & ".csv", FileFormat:=xlCSVUTF8
    ActiveWorkbook.Close False
    orig_wbk.Activate
    Application.DisplayAlerts = False
    ActiveSheet.Delete
    With Application
        .DisplayAlerts = True
        .ScreenUpdating = True
    End With
End Sub

Private Sub pvt_daily_report()
    Dim pvtCache As PivotCache
    Dim pvt As PivotTable
    Dim StartPvt As String

    Application.ScreenUpdating = False
    Sheets(1).Name = ActiveSheet.Range("F2")
    
    SrcData = ActiveSheet.Name & "!" & Range("A:M").Address(ReferenceStyle:=xlR1C1)
    
    Sheets.Add after:=Sheets(1)
    ActiveSheet.Name = "Summary"

    StartPvt = ActiveSheet.Name & "!" & ActiveSheet.Range("A1").Address(ReferenceStyle:=xlR1C1)
    
    Set pvtCache = ActiveWorkbook.PivotCaches.Create( _
    SourceType:=xlDatabase, _
    SourceData:=SrcData)

    Set pvt = pvtCache.CreatePivotTable( _
    TableDestination:=StartPvt, _
    TableName:="PivotTable1")
    
    pvt.PivotFields("Date Solved").Orientation = xlRowField
    pvt.PivotFields("Ticket Type").Orientation = xlColumnField
    pvt.AddDataField pvt.PivotFields("Date Requested"), "Ticket Count", xlCount
    pvt.PivotFields("Assignee").Orientation = xlPageField
    pvt.PivotFields("Assignee").PivotItems("(blank)").Visible = False

    Application.ScreenUpdating = True
End Sub

Private Sub create_notepad()
    Dim np As Object

    Application.ScreenUpdating = False
    Set np = CreateObject("Scripting.FileSystemObject").CreateTextFile(path & "LOG " & vendor_name & " " & date_now & ".txt", True)

    np.WriteLine (vendor_name & " " & date_now & vbNewLine & vbNewLine & _
        "OOS: " & oos & " | " & t_oos & "/" & update_item & vbNewLine & _
        "IN STOCK: " & in_stock & " | " & t_in_stock & "/" & update_item & vbNewLine & _
        "UPDATE SUMMARY: " & su_changes & "/" & update_item & vbNewLine & vbNewLine & _
        "NEW: " & new_item & vbNewLine & _
        "TBC: " & tbc_item & vbNewLine & _
        "FOR UPLOAD: " & upload_item & vbNewLine & _
        "ENABLE: " & en_item & vbNewLine & _
        "DISABLE: " & da_item & vbNewLine & _
        "VQ ALIGNED: " & vq_align & vbNewLine & _
        "OH HOLD: " & on_hold & vbNewLine & _
        "SU RULE: " & rule_remarks & vbNewLine & _
        "SU RULE SKUs: " & su_item)
    np.Close
    
    rule_remarks = "- None -"

    Application.ScreenUpdating = True
    orig_wbk.Activate
End Sub

Private Sub copy_separate()
    Application.ScreenUpdating = False
    ActiveSheet.AutoFilter.Range.Columns("A:AT").Copy
    Workbooks.Add
    ActiveSheet.Range("A1").PasteSpecial Paste:=xlPasteValues
    ActiveSheet.Range("A1").PasteSpecial Paste:=xlPasteFormats
    ActiveWindow.Zoom = 70
    Application.CutCopyMode = False

    With ActiveSheet
        .Range("A:D").ColumnWidth = 2.2
        .Range("E:F").ColumnWidth = 9
        .Range("G:G").ColumnWidth = 13
        .Range("H:H,J:K").ColumnWidth = 18
        .Range("L:L").ColumnWidth = 38
        .Range("H:K").HorizontalAlignment = xlLeft
    End With

    ActiveSheet.Range("N:P").EntireColumn.Delete
End Sub

Private Sub format_mf()
    With Worksheets("Sheet1")
        .Range("A:A").SpecialCells(2).Offset(0, 15).Borders.LineStyle = xlContinuous
        .Range("A:A").SpecialCells(2).Offset(0, 14).Borders.LineStyle = xlContinuous
        .Range("A:A").SpecialCells(2).Offset(0, 13).Borders.LineStyle = xlContinuous
        .Range("A:A").SpecialCells(2).Offset(0, 12).Borders.LineStyle = xlContinuous
        .Range("A:A").SpecialCells(2).Offset(0, 11).Borders.LineStyle = xlContinuous
        .Range("A:A").SpecialCells(2).Offset(0, 10).Borders.LineStyle = xlContinuous
        .Range("A:A").SpecialCells(2).Offset(0, 9).Borders.LineStyle = xlContinuous
        .Range("A:A").SpecialCells(2).Offset(0, 8).Borders.LineStyle = xlContinuous
        .Range("A:A").SpecialCells(2).Offset(0, 7).Borders.LineStyle = xlContinuous
        .Range("A:A").SpecialCells(2).Offset(0, 6).Borders.LineStyle = xlContinuous
        .Range("A:A").SpecialCells(2).Offset(0, 5).Borders.LineStyle = xlContinuous
        .Range("A:A").SpecialCells(2).Offset(0, 4).Borders.LineStyle = xlContinuous
        .Range("A:A").SpecialCells(2).Offset(0, 3).Borders.LineStyle = xlContinuous
        .Range("A:A").SpecialCells(2).Offset(0, 2).Borders.LineStyle = xlContinuous
        .Cells.Interior.ColorIndex = 0
        .Range("C1") = "ITEM BRAND"
        .Range("D1") = "BARCODE"
        .Range("G1") = "MW SKU"
        .Range("F1") = "VENDOR SKU"
        .Range("H1") = "PRODUCT NAME"
        .Range("J1") = "QTY"
        .Range("O1") = "COST"
        .Range("P1") = "RETAIL"
    End With

    Range("D:D").EntireColumn.Insert
    Range("C:C").EntireColumn.Insert

    With Worksheets("Sheet1")
        .Range("I:I").Cut Range("C:C")
        .Range("H:H").Cut Range("E:E")
        .Range("M:M").EntireColumn.Delete
        .Range("K:K").EntireColumn.Delete
        .Range("G:I").EntireColumn.Delete
        .Range("I:K").EntireColumn.Delete
        .Range("C:H").Columns.AutoFit
        .Columns("A:A").EntireColumn.Hidden = True
        .Columns("B:B").EntireColumn.Hidden = True
        .Range("1:1").Font.Bold = True
    End With
                                                                                                                          
    Cells.Select
    'removed no vsku
    Selection.Replace What:="NO VSKU", Replacement:="", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    'removed no vsku
    Selection.Replace What:="NO UPC", Replacement:="", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
End Sub
