Attribute VB_Name = "Main"
' --------------------------------------------
' Procedure: GetWordBankData
' Author: Aleksei Ogarkov
' Date: 2025-05-28
' Description:
'   High-level orchestrator that executes all steps in the World Bank data pipeline:
'   1. Loads indicators via API
'   2. Converts raw data into a Power Query table
'   3. Refreshes all data models and visualizations
' --------------------------------------------
Sub GetWordBankData()
    Application.StatusBar = "Step 1 of 3: Loading World Bank Indicators..."
    Call LoadWorldBankIndicators

    Application.StatusBar = "Step 2 of 3: Converting and loading query..."
    Call ConvertAndLoadQuery

    Application.StatusBar = "Step 3 of 3: Refreshing workbook data..."
    Call RefreshWorkbookData

    Application.StatusBar = "? All steps completed successfully."
    Application.Wait Now + TimeValue("0:00:01")
    Application.StatusBar = False

MsgBox "All data has been successfully updated.", vbInformation, "Update Complete"
    
End Sub

' --------------------------------------------
' Procedure: LoadWorldBankIndicators
' Author: Aleksei Ogarkov
' Date: 2025-05-28
' Description:
'   Retrieves macroeconomic data from the World Bank API for a given set of
'   countries and indicators over a specified time range. Results are stored in
'   the 'API_Import' sheet and activity is logged in the 'Log' sheet.
' --------------------------------------------
Sub LoadWorldBankIndicators()
    On Error GoTo ErrorHandler

    ' Optimize performance during execution
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Initializing..."

    ' Set references to relevant sheets
    Dim wsDataset As Worksheet, wsParams As Worksheet, wsLog As Worksheet
    Set wsDataset = ThisWorkbook.Sheets("API_Import")
    Set wsParams = ThisWorkbook.Sheets("Parameters")

    ' Prepare the Log sheet: create or clear
    On Error Resume Next
    Set wsLog = ThisWorkbook.Sheets("Log")
    If wsLog Is Nothing Then
        Set wsLog = ThisWorkbook.Sheets.Add(After:=wsParams)
        wsLog.Name = "Log"
    Else
        wsLog.Cells.Clear
    End If
    On Error GoTo ErrorHandler

    ' Clear output area and write header
    wsDataset.Cells.Clear
    wsDataset.Range("A1:D1").Value = Array("Country", "Indicator", "Year", "Value")

    ' Load named ranges: countries, indicators, year range
    Dim countries As Variant, indicators As Variant
    Dim startYear As Integer, endYear As Integer

    On Error GoTo ParamError
    countries = Range("Countries").Value
    indicators = Range("Indicators").Value
    startYear = Range("StartYear").Value
    endYear = Range("EndYear").Value

    ' Validate parameter input
    If IsEmpty(countries) Or IsEmpty(indicators) Or startYear = 0 Or endYear = 0 Then GoTo CleanExit
    If endYear < startYear Then GoTo CleanExit
    GoTo ContinueExecution

ParamError:
    Application.StatusBar = "? Parameter read error. Check named ranges."
    GoTo CleanExit

ContinueExecution:
    Dim outputData As Collection
    Set outputData = New Collection

    Dim json As String, parsed As Object, item As Object
    Dim i As Long, j As Long
    Dim rowLog As Long: rowLog = 1
    Dim countryCode As Variant, indicatorCode As Variant

    ' Loop through countries and indicators
    For i = 1 To UBound(countries, 1)
        countryCode = countries(i, 1)
        If Trim(countryCode) = "" Then GoTo NextCountry
        For j = 1 To UBound(indicators, 1)
            indicatorCode = indicators(j, 1)
            If Trim(indicatorCode) = "" Then GoTo NextIndicator

            Application.StatusBar = "Loading: " & countryCode & " / " & indicatorCode
            DoEvents

            ' API call
            json = GetWorldBankData(CStr(countryCode), CStr(indicatorCode), startYear, endYear)
            If json <> "" Then
                Set parsed = JsonConverter.ParseJson(json)
                If parsed.Count >= 2 Then
                    For Each item In parsed(2)
                        If Not IsNull(item("value")) Then
                            outputData.Add Array(item("country")("id"), item("indicator")("id"), item("date"), item("value"))
                        End If
                    Next item
                End If
                ' Log success
                rowLog = rowLog + 1
                wsLog.Cells(rowLog, 1).Value = Now
                wsLog.Cells(rowLog, 2).Value = countryCode
                wsLog.Cells(rowLog, 3).Value = indicatorCode
                wsLog.Cells(rowLog, 4).Value = "OK"
            Else
                ' Log failure
                rowLog = rowLog + 1
                wsLog.Cells(rowLog, 1).Value = Now
                wsLog.Cells(rowLog, 2).Value = countryCode
                wsLog.Cells(rowLog, 3).Value = indicatorCode
                wsLog.Cells(rowLog, 4).Value = "Failed"
            End If
NextIndicator:
        Next j
NextCountry:
    Next i

    ' Handle case of empty result
    If outputData.Count = 0 Then
        Application.StatusBar = "?? No data returned from API."
        GoTo CleanExit
    End If

    ' Output to worksheet
    Dim k As Long
    For k = 1 To outputData.Count
        wsDataset.Cells(k + 1, 1).Resize(1, 4).Value = outputData(k)
    Next k

    ' Format table
    With wsDataset
        .Rows(1).Font.Bold = True
        .Columns("A:D").AutoFit
        .Range("A1:D1").AutoFilter
    End With

    Application.StatusBar = "? Data loaded into API_Import."

CleanExit:
    ' Restore Excel settings
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub

ErrorHandler:
    Application.StatusBar = "? Error: " & Err.Description
    Resume CleanExit
End Sub

' --------------------------------------------
' Procedure: ConvertAndLoadQuery
' Author: Aleksei Ogarkov
' Date: 2025-05-28
' Description:
'   Converts the imported data range into an Excel Table named "RawData",
'   and creates/updates a Power Query connection pointing to it.
' --------------------------------------------
Sub ConvertAndLoadQuery()
    Dim ws As Worksheet
    Dim rng As Range
    Dim tbl As ListObject
    Dim wb As Workbook
    Dim lastRow As Long
    Dim pqName As String

    Application.StatusBar = "Converting to table..."
    Set wb = ThisWorkbook
    Set ws = wb.Sheets("API_Import")
    pqName = "RawDataQuery_vVBA"

    ' Identify used range
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then
        Application.StatusBar = "?? No data to convert."
        Exit Sub
    End If

    Set rng = ws.Range("A1:D" & lastRow)

    ' Recreate table
    On Error Resume Next
    Set tbl = ws.ListObjects("RawData")
    If Not tbl Is Nothing Then tbl.Unlist
    Set tbl = Nothing
    On Error GoTo 0

    Set tbl = ws.ListObjects.Add(xlSrcRange, rng, , xlYes)
    tbl.Name = "RawData"

    ' Replace Power Query
    Application.StatusBar = "Registering Power Query..."
    On Error Resume Next
    wb.Queries(pqName).Delete
    On Error GoTo 0

    wb.Queries.Add Name:=pqName, Formula:="let Source = Excel.CurrentWorkbook(){[Name=""RawData""]}[Content] in Source"

    Application.StatusBar = "? Table and Query registered."
End Sub

' --------------------------------------------
' Procedure: RefreshWorkbookData
' Author: Aleksei Ogarkov
' Date: 2025-05-28
' Description:
'   Triggers refresh of all queries, the data model, and PivotTables.
'   Assumes all PivotTables are based on the data model.
' --------------------------------------------
Sub RefreshWorkbookData()
    Application.StatusBar = "Refreshing all Power Query connections and data model..."
    ThisWorkbook.RefreshAll
    DoEvents

    ' Wait briefly for background queries to complete
    Application.Wait Now + TimeValue("0:00:02")

    Application.StatusBar = "? All model-based PivotTables refreshed."
    

End Sub

