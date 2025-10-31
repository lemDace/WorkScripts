Sub UpdateProductOptionFinishes()
    Dim folderPath As String
    Dim wbTarget As Workbook, wbSource As Workbook
    Dim wsTarget As Worksheet, wsSource As Worksheet
    Dim dictG As Object, dictI As Object
    Dim lastRowSource As Long, lastRowTarget As Long
    Dim i As Long
    Dim key As String
    
    ' Folder where both files are located
    folderPath = ThisWorkbook.Path & "\"
    
    ' Open both files
    Set wbTarget = Workbooks.Open(folderPath & "UD03_ProductOptionFinishes.xlsm")
    Set wbSource = Workbooks.Open(folderPath & "UD02_ProductOptions.xlsx")
    
    ' Set the first sheets (edit if needed)
    Set wsTarget = wbTarget.Sheets(1)
    Set wsSource = wbSource.Sheets(1)
    
    ' Create dictionaries for fast lookups
    Set dictG = CreateObject("Scripting.Dictionary")
    Set dictI = CreateObject("Scripting.Dictionary")
    
    ' Find last row in source sheet
    lastRowSource = wsSource.Cells(wsSource.Rows.Count, "C").End(xlUp).Row
    
    ' Load all key/value pairs from source (File2)
    For i = 2 To lastRowSource
        key = Trim(CStr(wsSource.Cells(i, "C").Value))
        If Len(key) > 0 Then
            dictG(key) = wsSource.Cells(i, "G").Value
            dictI(key) = wsSource.Cells(i, "I").Value
        End If
    Next i
    
    ' Find last row in target sheet
    lastRowTarget = wsTarget.Cells(wsTarget.Rows.Count, "C").End(xlUp).Row
    
    ' Add headers for new columns
    wsTarget.Cells(1, "K").Value = "Copied from UD02 Col G"
    wsTarget.Cells(1, "L").Value = "Copied from UD02 Col I"
    
    ' Loop through target and fill columns K and L where match found
    For i = 2 To lastRowTarget
        key = Trim(CStr(wsTarget.Cells(i, "C").Value))
        If dictG.exists(key) Then
            wsTarget.Cells(i, "K").Value = dictG(key)
            wsTarget.Cells(i, "L").Value = dictI(key)
        End If
    Next i
    
    ' 🔴 Highlight new columns in red (background)
    With wsTarget.Range("K1:L" & lastRowTarget).Interior
        .Color = RGB(255, 199, 206) ' light red fill
    End With
    
    ' Optional: make header text bold and red
    With wsTarget.Range("K1:L1").Font
        .Bold = True
        .Color = RGB(156, 0, 6)
    End With
    
    ' Save and close
    wbTarget.Save
    wbSource.Close SaveChanges:=False
    
    MsgBox "✅ Data update complete! Columns K and L highlighted in red.", vbInformation
End Sub
