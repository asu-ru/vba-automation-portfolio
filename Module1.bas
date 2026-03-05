Attribute VB_Name = "Module1"
Option Explicit


Sub WriteLog(phase As String, level As String, msg As String)
    With Sheets("Log")
        Dim r As Long
        r = .Cells(.Rows.Count, 1).End(xlUp).Row + 1
        .Cells(r, 1).Value = Now
        .Cells(r, 2).Value = phase
        .Cells(r, 3).Value = level
        .Cells(r, 4).Value = msg
    End With
End Sub

' （差分）SetRequestHeader の行を変更し、apiKey を Trim してから送る
Sub Main()

    Dim apiKey As String
    apiKey = Sheets("Config").Range("B2").Value

    Call GetDailyPrice_V2("7203", "2025-01-10", apiKey)

End Sub

Function GetDailyPrice_V2(code As String, targetDate As String, apiKey As String) As Boolean
    Dim http As Object
    Dim url As String
    Dim errNo As Long
    Dim keyTrim As String

    keyTrim = Trim(apiKey) ' ←余分な空白/改行を削る

    Call WriteLog("PRICE", "START", code & " " & targetDate)
    Call WriteLog("PRICE", "DEBUG", "apiKey length=" & Len(keyTrim) & " preview=" & Left$(keyTrim, IIf(Len(keyTrim) >= 6, 6, Len(keyTrim))))

    url = "https://api.jquants.com/v2/equities/bars/daily" & _
          "?code=" & code & _
          "&from=" & targetDate & _
          "&to=" & targetDate

    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.SetTimeouts 5000, 5000, 5000, 5000

    On Error Resume Next
    http.Open "GET", url, False

    ' ← ここを変更：V2 は x-api-key ヘッダーを使う
    http.SetRequestHeader "x-api-key", keyTrim

    http.Send
    errNo = Err.Number
    On Error GoTo 0

    If errNo <> 0 Then
        Call WriteLog("PRICE", "ERROR", "通信エラー Err=" & errNo)
        GetDailyPrice_V2 = False
        Exit Function
    End If

    If http.Status = 200 Then
        Call WriteLog("PRICE", "OK", "HTTP 200")
        GetDailyPrice_V2 = True
    Else
        Call WriteLog("PRICE", "ERROR", "HTTP " & http.Status & " " & http.ResponseText)
        GetDailyPrice_V2 = False
    End If
End Function

Sub Run_PQ_Update()

    Call WriteLog("PQ", "START", "Power Query refresh start")

    ThisWorkbook.RefreshAll

    Application.CalculateUntilAsyncQueriesDone

    Call WriteLog("PQ", "END", "Power Query refresh end")

End Sub

