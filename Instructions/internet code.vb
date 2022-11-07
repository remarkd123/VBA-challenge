'i like this because i know when i have
'incorrectly named a variable // otherwise
'VBA will just create the incorrectly-named variable
'and set its type to Variant
Option Explicit

'i like to explicitly state whether a sub or function
'will be able to be called from outsite the module (public)
'or if i want it to only be called from within the module (private)
Public Sub Stocks()
    'i prefer to keep all my 'Dim' statements in 1 block so they're
    'easier to find later when i need to change something
    Dim Total_Stock_Volume As Double
    Dim Summary_Table_Row As Integer
    Dim lastRow As Long
    Dim ws As Worksheet
    Dim j As Long
    Dim i As Long
    Dim Ticker As String

    'i prefer to clump my like assignments together in a block
    Summary_Table_Row = 2
    Total_Stock_Volume = 0

    'iterate through the collection of worksheets in your workbook
    For j = 1 To Excel.Application.ThisWorkbook.Worksheets.Count
        'set a reference to a worksheet // this will go through
        'the different worksheets in the workbook as the loop
        'progresses
        Set ws = Excel.Application.ThisWorkbook.Worksheets(j)

        With ws
            'this is a better way to get the last column in a worksheet
            lastRow = .Range("A" & .Rows.Count).End(xlUp).Row
        End With

        For i = 2 To lastRow
            'i prefer to explicitly cast anything i get from a cell to the type
            'i intend to use because .Value returns a Variant type by default
            Total_Stock_Volume = Total_Stock_Volume + CDbl(ws.Cells(i, 7).Value)

            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                'me explicitly casting the value of the cell to a string
                Ticker = CStr(ws.Cells(i, 1).Value)

                'i like with statements because it looks nicer to me.
                'i'm sure there's a better reason to use these, but that's
                'my reason!
                With ws
                    .Range("J" & Summary_Table_Row).Value = Ticker
                    .Range("K" & Summary_Table_Row).Value = Total_Stock_Volume
                End With
                Summary_Table_Row = Summary_Table_Row + 1

                'i'm unsure about the intention with this, so i'll leave it alone
                Total_Stock_Volume = 0
            End If
        Next i

    Next
End Sub
