Sub UpdateProductOptionFinishes()
    Dim folderPath As String
    Dim wbTarget As Workbook, wbSource As Workbook
    Dim wsTarget As Worksheet, wsSource As Worksheet
    Dim dictG As Object, dictI As Object
    Dim lastRowSource As Long, lastRowTarget As Long
    Dim lastCol As Long
    Dim i As Long
    Dim key As String
    
    ' Folder where both files are located
    folderPath = ThisWorkbook.Path & "\"
    
    ' Open both files
    Set wbTarget = Workbooks.Open(folderPath & "UD03_ProductOptionFinishes.xlsm")
    Set wbSource = Workbooks.Open(folderPath & "UD02_ProductOptions.xlsx")
    
    Set wsTarget = wbTarget.Sheets(1)
    Set wsSource = wbSource.Sheets(1)
    
    ' Create dictionaries for lookup
    Set dictG = CreateObject("Scripting.Dictionary")
    Set dictI = CreateObject("Scripting.Dictionary")
    
    ' Load source data
    lastRowSource = wsSource.Cells(wsSource.Rows.Count, "C").End(xlUp).Row
    
    For i = 2 To lastRowSource
        key = Trim(CStr(wsSource.Cells(i, "C").Value))
        If Len(key) > 0 Then
            dictG(key) = wsSource.Cells(i, "G").Value
            dictI(key) = wsSource.Cells(i, "I").Value
        End If
    Next i
    
    ' Prepare target range
    lastRowTarget = wsTarget.Cells(wsTarget.Rows.Count, "C").End(xlUp).Row
    lastCol = wsTarget.Cells(1, wsTarget.Columns.Count).End(xlToLeft).Column
    
    ' Add headers for new columns
    wsTarget.Cells(1, "K").Value = "Copied from UD02 Col G"
    wsTarget.Cells(1, "L").Value = "Copied from UD02 Col I"
    
    ' Fill new columns
    For i = 2 To lastRowTarget
        key = Trim(CStr(wsTarget.Cells(i, "C").Value))
        If dictG.Exists(key) Then
            wsTarget.Cells(i, "K").Value = dictG(key)
            wsTarget.Cells(i, "L").Value = dictI(key)
        End If
    Next i
    
    ' Highlight K and L
    With wsTarget.Range("K1:L" & lastRowTarget).Interior
        .Color = RGB(255, 199, 206)
    End With
    With wsTarget.Range("K1:L1").Font
        .Bold = True
        .Color = RGB(156, 0, 6)
    End With
    
' -----------------------------------------------------------
'  WORKING MULTI-COLUMN SORT WITH CORRECT PRIORITIES
' -----------------------------------------------------------
    Dim sortRange As Range
    Dim lastUsedRow As Long, lastUsedCol As Long
    
    With wsTarget.UsedRange
        lastUsedRow = .Rows(.Rows.Count).Row
        lastUsedCol = .Columns(.Columns.Count).Column
    End With
    
    Set sortRange = wsTarget.Range(wsTarget.Cells(1, 1), wsTarget.Cells(lastUsedRow, lastUsedCol))
    
    With wsTarget.Sort
        .SortFields.Clear
    
        ' Excel applies the LAST added sort field as the PRIMARY KEY
        ' So we add them in REVERSE priority order.
    
        ' 4th priority (lowest) – K
        .SortFields.Add key:=wsTarget.Range("K2:K" & lastUsedRow), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    
        ' 3rd priority – L
        .SortFields.Add key:=wsTarget.Range("L2:L" & lastUsedRow), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    
        ' 2nd priority – K again (tie-breaker)
        .SortFields.Add key:=wsTarget.Range("K2:K" & lastUsedRow), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    
        ' 1st priority (highest) – H
        .SortFields.Add key:=wsTarget.Range("H2:H" & lastUsedRow), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    
        .SetRange sortRange
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
    
        .Apply
    End With
' -----------------------------------------------------------


    
    ' Save and close
    wbTarget.Save
    wbSource.Close SaveChanges:=False
    
    MsgBox "? Data update complete and sorted.", vbInformation

End Sub


