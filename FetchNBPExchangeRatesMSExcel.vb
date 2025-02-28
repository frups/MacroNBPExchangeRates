Rem  *****  BASIC  *****

' Function to extract the "bid" value from the "rates" array
Function GetBidValue(JSON As String) As String
    Dim ratesStart As Integer
    Dim ratesEnd As Integer
    Dim ratesJson As String
    Dim bidValue As String

    ' Find the starting position of the "rates" array
    ratesStart = InStr(JSON, """rates"":[{") + 9
    If ratesStart = 9 Then
        GetBidValue = "No rates section"
        Exit Function
    End If

    ' Find the end of the first object in the "rates" array
    ratesEnd = InStr(ratesStart, JSON, "}]")
    If ratesEnd = 0 Then
        GetBidValue = "Invalid JSON structure"
        Exit Function
    End If

    ' Extract the first object from the "rates" array
    ratesJson = Mid(JSON, ratesStart, ratesEnd - ratesStart)

    ' Get the "bid" value from the extracted object
    bidValue = GetValueFromJson(ratesJson, "bid")

    GetBidValue = bidValue
End Function

' Function to extract a value for a given key in JSON (for simple structures)
Function GetValueFromJson(JSON As String, key As String) As String
    Dim startPos As Integer
    Dim endPos As Integer
    Dim result As String

    ' Find the position of the key in the JSON string
    startPos = InStr(JSON, """" & key & """") + Len(key) + 3
    If startPos = 0 Then
        GetValueFromJson = "Key not found"
        Exit Function
    End If

    ' Find the end of the value (comma or closing brace)
    endPos = InStr(startPos, JSON, ",")
    If endPos = 0 Then
        endPos = InStr(startPos, JSON, "}")  ' If it's the last value in the JSON
    End If

    ' Extract the value and return it
    result = Mid(JSON, startPos, endPos - startPos)
    GetValueFromJson = Trim(Replace(result, """", ""))
End Function

Sub FetchNBPExchangeRates(cell As Range)
    Dim DateValue As String
    DateValue = cell.Value ' Assuming the date is in the cell in the format yyyy-mm-dd
    
    Dim URL As String
    URL = "https://api.nbp.pl/api/exchangerates/rates/c/eur/" & DateValue & "?format=json"
    
    Dim Http As Object
    Set Http = CreateObject("MSXML2.XMLHTTP")
    Http.Open "GET", URL, False
    Http.Send
    
    Dim JSON As String
    JSON = Http.responseText
    
    cell.Offset(0, 1).Value = GetBidValue(JSON)
End Sub

Sub ApplyMacroToRange()
    Dim cell As Range
    For Each cell In Range("M1:M35") ' Adjust the range as needed
        FetchNBPExchangeRates cell
    Next cell
End Sub
