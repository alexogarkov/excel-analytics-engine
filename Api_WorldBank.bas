Attribute VB_Name = "Api_WorldBank"
' =============================================
' Module: Api_WorldBank.bas
' Author: Aleksei Ogarkov
' Date: 2025-04-28
'
' Purpose:
'   Provides HTTP-based data access to the World Bank Open API.
'   Retrieves indicator data for a specific country and year range in JSON format.
'
' Function:
'   GetWorldBankData(countryCode, indicatorCode, startYear, endYear) As String
'
' Parameters:
'   - countryCode   : ISO Alpha-3 country code (e.g., "USA", "FRA")
'   - indicatorCode : World Bank indicator code (e.g., "NY.GDP.PCAP.CD")
'   - startYear     : Integer, beginning year of range
'   - endYear       : Integer, ending year of range
'
' Returns:
'   - String (JSON response from API if successful; empty string otherwise)
'
' Dependencies:
'   - Microsoft XML v6.0
'   - Microsoft Scripting Runtime
' =============================================
' Attribute VB_Name = "Api_WorldBank"

Option Explicit

' Returns a JSON response from the World Bank API for the given country, indicator, and year range.
Public Function GetWorldBankData(countryCode As String, indicatorCode As String, startYear As Integer, endYear As Integer) As String
    On Error GoTo ErrorHandler ' Enable structured error handling

    Dim http As Object
    ' Instantiate the HTTP request object using late binding (no reference needed)
    Set http = CreateObject("MSXML2.XMLHTTP")

    ' Build the World Bank API request URL
    Dim url As String
    url = "https://api.worldbank.org/v2/country/" & countryCode & _
          "/indicator/" & indicatorCode & _
          "?format=json&date=" & startYear & ":" & endYear & "&per_page=1000"

    ' Perform the HTTP GET request
    With http
        .Open "GET", url, False ' Synchronous GET request
        .setRequestHeader "Content-Type", "application/json" ' Set request type to JSON
        .Send ' Execute request

        If .Status = 200 Then
            ' If request was successful, return the response text
            GetWorldBankData = .responseText
        Else
            ' If the response failed, log the HTTP status to the Immediate window
            Debug.Print "HTTP Error: " & .Status & " - " & .statusText
            GetWorldBankData = "" ' Return an empty string on error
        End If
    End With

    Exit Function ' Skip the error handler if execution is successful

ErrorHandler:
    ' Log the error details to the Immediate window for debugging
    Debug.Print "Request failed for Country: " & countryCode & _
                " | Indicator: " & indicatorCode & _
                " | Error: " & Err.Description
    GetWorldBankData = "" ' Return an empty string on failure
End Function

