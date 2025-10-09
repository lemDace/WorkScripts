'Attribute VB_Name = "Module1"
Sub BatchPrintAllRecords()
    Dim wsData As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim templateFolder As String
    Dim templatePath As String
    Dim modelName As String
    Dim wbTemplate As Workbook
    Dim wsTemplate As Worksheet
    Dim logFilePath As String
    Dim logFileNum As Integer
    Dim printSuccess As Boolean
    Dim timeStamp As String
    
    ' ===== SETTINGS =====
    templateFolder = "I:\Monash\Framemaking\" ' This is pointing to where all the empty frames templates are
    Set wsData = ThisWorkbook.Sheets("Sheet1") ' Data sheet in your opened workbook
    
    ' Create timestamped log filename
    timeStamp = Format(Now, "yyyy-mm-dd_hh-nn-ss")
    logFilePath = ThisWorkbook.Path & "\TemplatePrintLog_" & timeStamp & ".txt"
    
    ' Open log file for output (overwrite each run)
    logFileNum = FreeFile
    Open logFilePath For Output As #logFileNum
    Print #logFileNum, "----- Run started: " & Now & " -----"
    Print #logFileNum, "Row | ModelName     | Status           | Details"
    Print #logFileNum, String(65, "-")
    
    lastRow = wsData.Cells(wsData.Rows.Count, "V").End(xlUp).Row
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' ?? Show print dialog ONCE at the start
    ' If user presses Cancel, exit gracefully
    If Application.Dialogs(xlDialogPrint).Show = False Then   'Check dialog result
        MsgBox "Process cancelled. No printer selected.", vbExclamation
        Print #logFileNum, "Process cancelled by user: " & Now() 'Log cancellation
        Close #logFileNum
        Exit Sub                                             'Stop macro
    End If
    
    ' Store the chosen printer name for later use
    chosenPrinter = Application.ActivePrinter
    
    Print #logFileNum, "Printer selected: " & chosenPrinter   'Print to log the printer being used
    
    For r = 2 To lastRow
        
        modelName = wsData.Cells(r, "V").Value ' Adjust if ModelName is in another column
        templatePath = GetTemplatePathByModel(templateFolder, modelName)
        printSuccess = False
        
        If templatePath <> "" Then
        
            ' Open template
            Set wbTemplate = Workbooks.Open(templatePath, ReadOnly:=True) 'opens a template up as read only so it cannot be saved over
            Set wsTemplate = wbTemplate.Sheets(1) ' Adjust if needed
            
            ' Fill your template cells here (adjust as needed)
            wsTemplate.Range("B5").Value = wsData.Cells(r, "Y").Value 'order#
            wsTemplate.Range("D5").Value = wsData.Cells(r, "AB").Value 'Date
            wsTemplate.Range("B6").Value = wsData.Cells(r, "AD").Value 'QTY
            
            'Check if this is row is a critical order, if so the change fill color to red
            If wsData.Cells(r, "AG").Value = "X" Then
                wsTemplate.Range("A3", "G6").Interior.Color = RGB(255, 0, 0) 'red,green,blue
            End If
            
            'Check if the cell in column W contains the word "CUSTOM" (case-insensitive)
            If InStr(1, LCase(wsData.Cells(r, "W").Value), "custom") > 0 Then
                wsTemplate.Range("A3", "G6").Interior.Color = RGB(255, 255, 0) 'yellow
            End If
                                    
            wsTemplate.Columns("D").EntireColumn.AutoFit 'date column was not big enough so this stret
                                    
            'setup the worksheet for printing here
            With wsTemplate.PageSetup
            
                'DEBUG LINE: Print #logFileNum, "setting up page for printing"
                .FitToPagesWide = 1 'scale page so that all columns fit on the same page
                .Zoom = False
                '.Orientation = xlLandscape
                            
            End With
                                                
            ' Print with error handling
            On Error Resume Next
            Application.ActivePrinter = chosenPrinter
            wsTemplate.PrintOut Copies:=1, Collate:=True, IgnorePrintAreas:=False
            If Err.Number = 0 Then
                printSuccess = True
            End If
            On Error GoTo 0
            
            wbTemplate.Close SaveChanges:=False
            
            If printSuccess Then
                Print #logFileNum, r & "   | " & modelName & " | Printed OK       | " & templatePath
            Else
                Print #logFileNum, r & "   | " & modelName & " | Print Failed     | " & templatePath
            End If
        Else
            Print #logFileNum, r & "   | " & modelName & " | Template Missing | No file found matching '" & modelName & "'"
        End If
    Next r
    
    Print #logFileNum, String(65, "-")
    Print #logFileNum, "----- Run ended: " & Now & " -----"
    Close #logFileNum
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    MsgBox "Processing complete. Log saved to:" & vbCrLf & logFilePath
End Sub

' Helper function: returns full path of first template file containing modelName in filename
Function GetTemplatePathByModel(templateFolder As String, modelName As String) As String
    Dim fileName As String
    Dim foundTemplate As String
    
    foundTemplate = ""
    fileName = Dir(templateFolder & "*" & modelName & "*.xlsx") ' Wildcards around modelName
    
    If fileName <> "" Then
        foundTemplate = templateFolder & fileName
    End If
    
    GetTemplatePathByModel = foundTemplate
End Function



