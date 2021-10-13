Option Explicit
main

'Main Program
Sub main()
    Dim inputExcelFile, outputFolder, idFieldName
    Dim objFso, objExcelApp, objShXL, objXLBook, regexpress

    'Command line Variables
    inputExcelFile = "E:\Downloads\TEST Datei mit div. Produkten.xlsm"
    idFieldName = "Producto"
    outputFolder = "E:\Downloads\result"
    regexpress = " \([^\)]+\)"

    Set objFso = CreateObject("Scripting.FileSystemObject")

    Set objExcelApp = CreateObject("Excel.Application")
    objExcelApp.Visible = False
    objExcelApp.DisplayAlerts = False
    objExcelApp.ScreenUpdating = False
    objExcelApp.DisplayStatusBar = False


    'Open Work Excel File
    'WriteLine "OpenFile " & Time() 
    Set objXLBook = objExcelApp.Workbooks.Open(inputExcelFile, true)
    Dim arraySheets
    Dim region
    Dim resultWrk, tempSheet, rangeWithData, columnsRange
    Dim Sheet, SheetTmp
    Dim i
    Dim regExp, regExpCell, sName, Sformula, squery, columnName
    Dim headersDictionary, primaryDictionary, listY, listX
    Dim colPosGlobal, rowPosGlobal, primaryPosGlobal, tmpArray

    'WriteLine "File was Open " & Time() 
    Set headersDictionary = CreateObject("Scripting.Dictionary")
    Set primaryDictionary = CreateObject("Scripting.Dictionary")
    
    'Get sheets of the original Workbook
    Set arraySheets = objXLBook.Worksheets
    i = 0
    colPosGlobal = 0
    rowPosGlobal = 1

    'Pattern to identify worksheets that starts with 0
    Set regExp = CreateObject("VBScript.RegExp")
    regExp.IgnoreCase = True
    regExp.Global = True
    regExp.Pattern = "^[0]+" 'Pattern for name of sheets

    'Pattern to find _(AnyText) and replace
    Set regExpCell = CreateObject("VBScript.RegExp")
    regExpCell.IgnoreCase = True
    regExpCell.Global = True
    regExpCell.Pattern = regexpress 

    'WriteLine "Start merging " & Time()
    
    'Only for Calculations - Array length
    Dim xCal, ycal
    xCal=0
    ycal=0
    For Each SheetTmp In arraySheets
        If regExp.Test(SheetTmp.Name) Then
            Dim rangetmp
            Set rangetmp = SheetTmp.Range("A10").CurrentRegion
            xCal = xCal + rangetmp.Rows.Count
            yCal = yCal + rangetmp.Columns.Count
        End If
    Next
    
    'Initialize Array with max rows and columns
    'WriteLine "xcal " & xCal & " ycal" & ycal
    Dim arryTmp()
    reDim arryTmp(xCal, yCal)

    'Loop for each sheet that its name starts with 0
    'and copy that information in new temp workbook
    For Each Sheet In arraySheets
        If regExp.Test(Sheet.Name) Then
            Set rangeWithData = Sheet.Range("A10").CurrentRegion
            Set columnsRange = rangeWithData.Columns
            
            'Combine Local headers
            Dim colPosLocal, primaryPosSelection, rowPosLocal            
            Dim headerValue, rowValue, rowsCount, columnsCount
            'rowValue = rangeWithData.Cells(rowPosLocal, colPosLocal).value
            primaryPosSelection = 1
            tmpArray =  rangeWithData.value2
            rowsCount = UBound(tmpArray, 1) - LBound(tmpArray, 1) + 1
            columnsCount = UBound(tmpArray, 2) - LBound(tmpArray, 2) + 1
            
            'Loop Rows
            For rowPosLocal = 1 To rowsCount
                'Set listY = CreateObject("System.Collections.Sortedlist")
                
                Dim existRow, rowWork
                If rowPosLocal > 1 Then
                    'check if primary key already exist 
                    existRow = primaryDictionary.Exists(tmpArray(rowPosLocal, primaryPosSelection))
                    If Not existRow Then
                        'Increase one more row
                        rowPosGlobal = rowPosGlobal + 1
                        rowWork = rowPosGlobal
                        'listX.Add rowWork, listY
                    Else
                        'Use the same row
                        rowWork = primaryDictionary(tmpArray(rowPosLocal, primaryPosSelection))
                    End If
                End If

                'Loop cols
                For colPosLocal = 1 To columnsCount
                    headerValue = tmpArray(1, colPosLocal)
                    'Check if column header exist else create
                    If Not headersDictionary.Exists(headerValue) Then
                        colPosGlobal = colPosGlobal + 1
                        headersDictionary.Add headerValue, colPosGlobal
                        arryTmp(0, colPosGlobal-1) = regExpCell.Replace(headerValue, "")
                        'addFieldToNewTable resultWrk.Worksheets(1), 1, colPosGlobal, headerValue
                    ElseIf rowPosLocal > 1 And Not existRow Then 'add new row
                        arryTmp(rowWork-1, headersDictionary(headerValue)-1) = regExpCell.Replace(tmpArray(rowPosLocal, colPosLocal), "")
                        'addFieldToNewTable resultWrk.Worksheets(1), rowWork, headersDictionary(headerValue), tmpArray(rowPosLocal, colPosLocal)
                    End If
                    
                    If rowPosLocal = 1 And headerValue = idFieldName Then
                        primaryPosSelection = colPosLocal
                    End If
                    
                    'If not exist PK, Add primary key to map.
                    If rowPosLocal > 1 And colPosLocal = primaryPosSelection And Not existRow Then
                        primaryDictionary.Add tmpArray(rowPosLocal, colPosLocal), rowPosGlobal
                    End If                    
                Next
            Next
            
        End If
    Next   

    WriteLine "End merging " & Time()
    objXLBook.Close False

    WriteLine "Create new File " & Time()
    Set resultWrk = objExcelApp.Workbooks.Add

    WriteLine "new File Open " & Time()

    WriteLine "Copy Data in new File " & Time()
    Dim Destination 
    Set Destination = resultWrk.Worksheets(1).Range("A1")
    Destination.Resize(rowPosGlobal, colPosGlobal).Value2 = arryTmp
    WriteLine "End Copy Data in new File " & Time()
    resultWrk.SaveAs outputFolder & "\File_2- End_result.xlsx"
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
    TmpSheet.Cells(rowPos, colPos).value2 = value
    
    
End Sub