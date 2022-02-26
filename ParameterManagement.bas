Attribute VB_Name = "ParameterManagement"
' ========================================
' Last Revision Date = 02/01/2022, 06:52 PM
' Module: BatchFileManagement
' Developed by: Mark Kranz
' ========================================
Sub GetGlobalParameter(parameter As String, value As Variant, errCode As Integer, errMsg As String)
    Dim globalParamSheet As Worksheet
    Dim globalParamLastRow As Long
    Dim result As String
        
    errCode = 0
    errMsg = ""
    Set globalParamSheet = ThisWorkbook.Worksheets("GlobalParameters")
    globalParamLastRow = globalParamSheet.Cells(globalParamSheet.Rows.Count, 1).End(xlUp).row
    
    If Not (parameter = "") Then
        ' Search for Part Number in BOMs Index.
        result = ""
        On Error Resume Next
        result = Application.WorksheetFunction.VLookup(parameter, globalParamSheet.Range(Cells(1, 1).Address, Cells(globalParamLastRow, 1).Address), 1, False)
        If (result = parameter) Then
            value = Application.WorksheetFunction.VLookup(parameter, globalParamSheet.Range(Cells(1, 1).Address, Cells(globalParamLastRow, 2).Address), 2, False)
        Else
            value = ""
        End If
        On Error GoTo 0
        
        ' If Part Number not was found
        If (value = "") Then
            errCode = 99
            errMsg = "GetGlobalParameter: parameter (" & parameter & ") not found"
        End If
    Else
        errCode = 1
        errMsg = "GetGlobalParameter: parameter was not specified"
    End If
End Sub

Sub GetGlobalParameterRow(parameter As String, rowNumber As Long, errCode As Integer, errMsg As String)
    Dim globalParamSheet As Worksheet
    Dim globalParamLastRow As Long
        
    errCode = 0
    errMsg = ""
    Set globalParamSheet = ThisWorkbook.Worksheets("GlobalParameters")
    globalParamLastRow = globalParamSheet.Cells(globalParamSheet.Rows.Count, 1).End(xlUp).row
    
    If Not (parameter = "") Then
        ' Search for Part Number in BOMs Index.
        rowNumber = 0
        On Error Resume Next
        rowNumber = Application.Match(parameter, globalParamSheet.Range(Cells(1, 1).Address, Cells(globalParamLastRow, 1).Address), 0)
        On Error GoTo 0
        
        ' If Part Number not was found
        If (rowNumber = 0) Then
            errCode = 99
            errMsg = "GetGlobalParameterRow: parameter (" & parameter & ") not found"
        End If
    Else
        errCode = 1
        errMsg = "GetGlobalParameterRow: parameter was not specified"
    End If
End Sub

Sub GetGlobalNameValuePair(rowNumber As Long, name As String, value As Variant, errCode As Integer, errMsg As String)
    Dim globalParamSheet As Worksheet
    Dim globalParamLastRow As Long
        
    errCode = 0
    errMsg = ""
    Set globalParamSheet = ThisWorkbook.Worksheets("GlobalParameters")
    globalParamLastRow = globalParamSheet.Cells(globalParamSheet.Rows.Count, 1).End(xlUp).row
    
    If (rowNumber <= globalParamLastRow And rowNumber > 0) Then
        ' Get name and value.
        name = globalParamSheet.Cells(rowNumber, 1).value
        value = globalParamSheet.Cells(rowNumber, 2).value
    Else
        errCode = 1
        errMsg = "GetGlobalNameValuePair: rowNumber (" & rowNumber & ") is outside of valid range."
    End If
End Sub

Sub GetApplicationParameter(parameter As String, value As Variant, errCode As Integer, errMsg As String)
    Dim applicationParamSheet As Worksheet
    Dim applicationParamLastRow As Long
    Dim result As String
        
    errCode = 0
    errMsg = ""
    Set applicationParamSheet = ThisWorkbook.Worksheets("ApplicationParameters")
    applicationParamLastRow = applicationParamSheet.Cells(applicationParamSheet.Rows.Count, 1).End(xlUp).row
    
    If Not (parameter = "") Then
        ' Search for Part Number in BOMs Index.
        result = ""
        On Error Resume Next
        result = Application.WorksheetFunction.VLookup(parameter, applicationParamSheet.Range(Cells(1, 1).Address, Cells(applicationParamLastRow, 1).Address), 1, False)
        If (result = parameter) Then
            value = Application.WorksheetFunction.VLookup(parameter, applicationParamSheet.Range(Cells(1, 1).Address, Cells(applicationParamLastRow, 2).Address), 2, False)
        Else
            value = ""
        End If
        On Error GoTo 0
        
        ' If Part Number not was found
        If (value = "") Then
            errCode = 99
            errMsg = "GetApplicationParameter: parameter (" & parameter & ") not found"
        End If
    Else
        errCode = 1
        errMsg = "GetApplicationParameter: parameter not specified"
    End If
End Sub

Sub ExportGlobalParameters(errCode As Integer, errMsg As String)
    Dim globalParamSheet As Worksheet
    Dim M3ReportsPath As String
    
    Dim lastCol, lastRow As Long
    Dim textLine, delimiter As String
    Dim fileNumber As Integer
    Dim i, j As Integer
    
    ' Tab delimiter
    delimiter = Chr(9)
    
    errCode = 0
    errMsg = ""
    Set globalParamSheet = ThisWorkbook.Worksheets("GlobalParameters")
    
    Call GetApplicationParameter("M3ReportsPath", M3ReportsPath, errCode, errMsg)
    
    fileNumber = FreeFile
    
    lastRow = globalParamSheet.Cells.SpecialCells(xlCellTypeLastCell).row
    lastCol = globalParamSheet.Cells.SpecialCells(xlCellTypeLastCell).Column
    
    Open M3ReportsPath & "\GlobalParameters.tsv" For Output As fileNumber
    
    For i = 1 To lastRow
        For j = 1 To lastCol
            If (j = lastCol) Then
                textLine = textLine & globalParamSheet.Cells(i, j).value
            Else
                textLine = textLine & globalParamSheet.Cells(i, j).value & delimiter
            End If
        Next j
        
        Print #fileNumber, textLine
        textLine = ""
    Next i
    
    Close #fileNumber
    
    MsgBox ("Generated: " & M3ReportsPath & "\GlobalParameters.tsv")
End Sub

Sub ImportGlobalParameters(errCode As Integer, errMsg As String)
    Dim i As Integer
    Dim currentRow As Long
    Dim nameString As String
    Dim fieldName1 As String
    Dim fieldCol As Integer
    Dim fieldHeading1 As String
    Dim fieldHeading2 As String
    Dim fieldColumn As String
    Dim fieldChecked As String
    Dim fieldHeading1Length As Integer
    Dim fieldHeading2Length As Integer
    Dim fieldColumnLength As Integer
    Dim fieldCheckedLength As Integer
    Dim tempString As String
    Dim blanks As String

    Dim formatArray As Variant

    Dim directory As String
    Dim fileName As String
    Dim sheetName As String
    Dim noOfHeadings As Integer
    
    Dim fieldsStartRow As Long
    Dim fieldsEndRow As Long
    Dim noOfFields As Integer

    Dim sheet As Worksheet
    Dim oldSheet As Worksheet
    
    Dim total As Integer
    Dim sheetExists As Boolean
    
    Dim fileDate As Date
    Dim postProcessFileDate As Date
    
    blanks = "                              "
    errCode = 0
    errMsg = ""
    
    'Initialize format array.
    formatArray = Array(xlGeneralFormat, xlGeneralFormat, xlGeneralFormat, _
                        xlGeneralFormat, xlGeneralFormat, xlGeneralFormat)
    
     
    'Get file path to M3 Reports Folder
    Call GetApplicationParameter("M3ReportsPath", directory, errCode, errMsg)

    If errCode <> 0 Then
        Exit Sub
    End If
    
     
    'Set file name and sheet name.
    fileName = "GlobalParameters.tsv"
    sheetName = "GlobalParameters"
         
    'Turn off screen updating and displaying alerts.
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    'Delete older version of sheet first, if it exists.
    On Error Resume Next
    Set oldSheet = ThisWorkbook.Sheets(sheetName)
    sheetExists = Not oldSheet Is Nothing
    If (sheetExists) Then
        ThisWorkbook.Sheets(sheetName).Delete
    End If
    On Error GoTo 0

    ' If file exists
    If fileExists(directory, fileName) Then
        ' Import latest M3 Global Parameters File with appropriate column format.
        Set sheet = ThisWorkbook.Sheets.Add()
        With sheet.QueryTables.Add(Connection:="TEXT;" & directory & "\" & fileName, Destination:=Range("A1"))
            .name = "mytest"
            .FieldNames = False
            .AdjustColumnWidth = True
            .TextFileStartRow = 1
            .TextFileParseType = xlDelimited
            .TextFileTextQualifier = xlTextQualifierDoubleQuote
            .TextFileConsecutiveDelimiter = False
            .TextFileTabDelimiter = True
            .TextFileSemicolonDelimiter = False
            .TextFileCommaDelimiter = False
            .TextFileSpaceDelimiter = False
            .TextFileColumnDataTypes = formatArray
            .Refresh BackgroundQuery:=False
        End With
        sheet.name = sheetName
        sheet.Move After:=ThisWorkbook.Worksheets("ApplicationParameters")
        sheet.Columns("A:B").HorizontalAlignment = xlLeft
        sheet.Columns("A:B").AutoFit

    Else
        errCode = 1
        errMsg = "ImportGlobalParameters: " & fileName & " not found!" & Chr(10) & _
                   "File Path: " & directory & "\" & fileName & Chr(10) & _
                   "Contact Mark Kranz"
        Exit Sub
    End If

    'Turn on screen updating and displaying alerts again
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub


