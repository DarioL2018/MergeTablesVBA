Option Explicit

main

'Main Program
Sub main()
    Dim inputExcelFile, outputFolder, idFieldName
    Dim objFso, objExcelApp, objShXL, objXLBook

    If WScript.Arguments.Count <> 3 Then
		WriteLine "You need to specify input and output files."
		WScript.Quit
	End If

    'Command line Variables 
    inputExcelFile = WScript.Arguments(0)
    idFieldName = WScript.Arguments(1)
    outputFolder = WScript.Arguments(2)

    set objFso = CreateObject("Scripting.FileSystemObject")

    If Not objFso.FileExists( inputExcelFile ) Then
		WriteLine "Unable to find your input file " & inputExcelFile
		WScript.Quit
	End If

    set objExcelApp = CreateObject("Excel.Application")
    objExcelApp.visible = False
    objExcelApp.DisplayAlerts = False 
    'Open Work Excel File 
    Set objXLBook = objExcelApp.workbooks.open(inputExcelFile)
    Dim arraySheets
    Dim region
    Dim resultWrk
    Dim Sheet
    Dim i
    dim regExp, sName, Sformula, squery

    Set arraySheets = objXLBook.Worksheets  
    i = 0    
    set regExp=CreateObject("VBScript.RegExp")
    regExp.IgnoreCase = true
    regExp.Global = true
    regExp.Pattern = "0*" 'Pattern for name of sheets
    squery = "Table.Distinct(Table.Combine({"
    For Each Sheet In arraySheets
        If regExp.Test(Sheet.Name) Then
            sName = "Table" & i
            if(i>0) Then
                squery = squery & ", "
            end if
            Set region = Sheet.Range("A1").CurrentRegion
            'Create Query
            Sheet.ListObjects.Add(1, region, , 1).Name = sName
            Sformula = "Excel.CurrentWorkbook() {[Name=""" & sName & """]}[Content]"
            objXLBook.Queries.Add sName, ("let" & Chr(13) & "" & Chr(10) & " Source =" & Sformula _
            & "" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & " Source")
            'Create Connection
            objXLBook.Connections.Add2 "Query – " & sName, _
            "Connection to the '" & sName & "' query in the workbook.", _
            "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" _
            & sName & ";Extended Properties=""""", _
            "SELECT * FROM [" & sName & "]", _
            2, _
            False, _
            False
            'Add Tables name to String
            squery = squery & sName 
            i = i + 1
        End If
    Next
    squery = squery & "})"

    'Call Method to Create Merge Table
    generateResultTable objXLBook, squery, idFieldName
    
    'Copy Values from Merge Table
    objXLBook.Sheets("Merge Table").ListObjects(1).Range.Copy
    
    'Create new Workbook
    set resultWrk = objExcelApp.workbooks.Add
    
    'Paste Values in new WorkBook
    resultWrk.Worksheets(1).Range("A1").PasteSpecial -4163 'xlPasteValues
    
    'Save Workbook and Close
    resultWrk.SaveAs outputFolder & "\File_2- End_result.xlsx"
    objXLBook.Close False
    resultWrk.Close
    objExcelApp.Quit
    WriteLine "Process Completed"
End Sub

Sub generateResultTable(objXLBook, squery, idFieldName)
    Dim qry
    Dim currentSheet
    squery = squery & ", {""" & idFieldName & """})"
    'Create new Query to Append Tables
    objXLBook.Queries.Add "Query - Merge", _
    ("let" & Chr(13) & "" & Chr(10) & " Source =" & squery & "" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & " Source")

    Set qry = objXLBook.Queries("Query - Merge")
    'Create new Sheet at end of workbook
    Set currentSheet = objXLBook.Sheets.Add(,objXLBook.Sheets(objXLBook.Sheets.Count))
    currentSheet.Name = "Merge Table"

    'Load Query on Sheet
    LoadToWorksheetOnly qry, currentSheet    
End Sub

Sub LoadToWorksheetOnly(query, currentSheet)
     Dim qryTab
     'Create Table with Append Query 
     set qryTab = currentSheet.ListObjects.Add(3, _
        ("OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & query.Name), _
         , 1, currentSheet.Range("$A$1")).QueryTable
         With qryTab
            .CommandType = 4 'xlCmdDefault
            .CommandText = Array("SELECT * FROM [" & query.Name & "]")
            .RowNumbers = False
            .FillAdjacentFormulas = False
            .PreserveFormatting = True
            .RefreshOnFileOpen = False
            .BackgroundQuery = True
            .RefreshStyle = 1 'xlInsertDeleteCells
            .SavePassword = False
            .SaveData = True
            .AdjustColumnWidth = True
            .RefreshPeriod = 0
            .PreserveColumnInfo = False
            .Refresh False
        End With
     
End Sub

'Subrutine to write String in Console
Sub WriteLine ( strLine )
    'WScript.Stdout.WriteLine strLine
	WScript.Echo strLine
End Sub