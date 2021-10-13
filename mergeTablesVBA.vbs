Option Explicit
main

'Main Program
Sub main()
    Dim inputExcelFile, outputFolder, idFieldName
    Dim objFso, objExcelApp, objShXL, objXLBook

    'Command line Variables
    inputExcelFile = "E:\Downloads\TEST Datei mit div. Produkten.xlsm"
    idFieldName = "Producto"
    outputFolder = "E:\Downloads\result"

    Set objFso = CreateObject("Scripting.FileSystemObject")

    Set objExcelApp = CreateObject("Excel.Application")
    objExcelApp.Visible = False
    objExcelApp.DisplayAlerts = False
    objExcelApp.ScreenUpdating = False
    objExcelApp.DisplayStatusBar = False
    


    'Open Work Excel File
    WriteLine "OpenFile " & Time() 
    Set objXLBook = objExcelApp.Workbooks.Open(inputExcelFile)
    Dim arraySheets
    Dim region
    Dim resultWrk, tempSheet, rangeWithData, columnsRange
    Dim Sheet
    Dim i
    Dim regExp, sName, Sformula, squery, columnName
    Dim headersDictionary, primaryDictionary
    Dim colPosGlobal, rowPosGlobal, primaryPosGlobal
    WriteLine "File was Open " & Time() 
    Set headersDictionary = CreateObject("Scripting.Dictionary")
    Set primaryDictionary = CreateObject("Scripting.Dictionary")
    
    WriteLine "Create new File " & Time()
    Set resultWrk = objExcelApp.Workbooks.Add

    WriteLine "new File Open " & Time()
    'Get sheets of the original Workbook
    Set arraySheets = objXLBook.Worksheets
    i = 0
    colPosGlobal = 0
    rowPosGlobal = 1
    Set regExp = CreateObject("VBScript.RegExp")
    regExp.IgnoreCase = True
    regExp.Global = True
    regExp.Pattern = "^[0]+" 'Pattern for name of sheets
    'Loop for each sheet that its name starts with 0
    'and copy that information in new temp workbook
    
    WriteLine "Start merging " & Time()

    For Each Sheet In arraySheets
        If regExp.Test(Sheet.Name) Then
            Set rangeWithData = Sheet.Range("A10").CurrentRegion
            Set columnsRange = rangeWithData.Columns
            
            'Combine Local headers
            Dim colPosLocal, primaryPosSelection, rowPosLocal            
            Dim headerValue, rowValue
            'rowValue = rangeWithData.Cells(rowPosLocal, colPosLocal).value
            primaryPosSelection = 1
            
            For rowPosLocal = 1 To rangeWithData.Rows.Count
                Dim existRow, rowWork
                If rowPosLocal > 1 Then
                    existRow = primaryDictionary.Exists(rangeWithData.Cells(rowPosLocal, primaryPosSelection).value)
                    If Not existRow Then
                        rowPosGlobal = rowPosGlobal + 1
                        rowWork = rowPosGlobal
                    Else
                        rowWork = primaryDictionary(rangeWithData.Cells(rowPosLocal, primaryPosSelection).value)
                    End If
                End If
                For colPosLocal = 1 To rangeWithData.Columns.Count
                    headerValue = rangeWithData.Cells(1, colPosLocal).value
                    If Not headersDictionary.Exists(headerValue) Then
                        colPosGlobal = colPosGlobal + 1
                        headersDictionary.Add headerValue, colPosGlobal
                        addFieldToNewTable resultWrk.Worksheets(1), 1, colPosGlobal, headerValue
                    ElseIf rowPosLocal > 1 And Not existRow Then
                        addFieldToNewTable resultWrk.Worksheets(1), rowWork, headersDictionary(headerValue), rangeWithData.Cells(rowPosLocal, colPosLocal).value
                    End If
                    
                    If rowPosLocal = 1 And headerValue = idFieldName Then
                        primaryPosSelection = colPosLocal
                    End If
                    
                    If rowPosLocal > 1 And colPosLocal = primaryPosSelection And Not existRow Then
                        primaryDictionary.Add rangeWithData.Cells(rowPosLocal, colPosLocal).value, rowPosGlobal
                    End If
                    
                Next 
            Next
        End If
    Next   
    WriteLine "End merging " & Time()

    'Save Workbook and Close
    resultWrk.SaveAs outputFolder & "\File_2- End_result.xlsx"
    objXLBook.Close False
    resultWrk.Close
    objExcelApp.Quit
    WriteLine "Process Completed " & Time()
End Sub

'Subrutine to write String in Console
Sub WriteLine(strLine)
    'WScript.Stdout.WriteLine strLine
    WScript.Echo strLine
End Sub


Sub addFieldToNewTable(TmpSheet, rowPos, colPos, value)
    TmpSheet.Cells(rowPos, colPos).value = value
    
End Sub