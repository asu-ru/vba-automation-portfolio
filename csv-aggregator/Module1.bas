Attribute VB_Name = "Module1"
Option Explicit


Public Sub AggregateCSV()

    Dim folder As String
    Dim file As String
    Dim ws As Worksheet
    Dim row As Long

    On Error GoTo ErrHandler

    folder = ThisWorkbook.Path & "\csv"
    If Dir(folder, vbDirectory) = "" Then
        MsgBox "csvフォルダがありません"
        Exit Sub
    End If

    Set ws = ThisWorkbook.Sheets.Add
    row = 1

    file = Dir(folder & "\*.csv")
    Do While file <> ""

        With ws.QueryTables.Add(Connection:="TEXT;" & folder & "\" & file, Destination:=ws.Cells(row, 1))
            .TextFileCommaDelimiter = True
            .Refresh
        End With

        row = ws.UsedRange.Rows.Count + 1
        file = Dir()

    Loop

    MsgBox "完了"

    Exit Sub

ErrHandler:
    MsgBox "エラー：" & Err.Description

End Sub
