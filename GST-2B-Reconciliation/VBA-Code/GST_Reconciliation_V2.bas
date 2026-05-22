Attribute VB_Name = "Module1"
Sub GST_Reconciliation_V2()

    Dim wsBooks As Worksheet
    Dim ws2B As Worksheet

    Dim lastRowBooks As Long
    Dim lastRow2B As Long

    Dim i As Long
    Dim j As Long

    Dim bookGSTIN As String
    Dim bookInvoice As String
    Dim bookValue As Double

    Dim foundMatch As Boolean

    Set wsBooks = Sheets("Books_Data")
    Set ws2B = Sheets("GSTR2B_Data")

    lastRowBooks = wsBooks.Cells(wsBooks.Rows.Count, 1).End(xlUp).Row
    lastRow2B = ws2B.Cells(ws2B.Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastRowBooks

        foundMatch = False

        bookGSTIN = wsBooks.Cells(i, 1).Value
        bookInvoice = wsBooks.Cells(i, 2).Value
        bookValue = wsBooks.Cells(i, 3).Value

        For j = 2 To lastRow2B

            If ws2B.Cells(j, 1).Value = bookGSTIN _
            And ws2B.Cells(j, 2).Value = bookInvoice Then

                foundMatch = True

                If ws2B.Cells(j, 3).Value = bookValue Then

                    wsBooks.Cells(i, 4).Value = "MATCH"
                    wsBooks.Cells(i, 4).Interior.Color = RGB(0, 255, 0)

                Else

                    wsBooks.Cells(i, 4).Value = "VALUE MISMATCH"
                    wsBooks.Cells(i, 4).Interior.Color = RGB(255, 255, 0)

                End If

                Exit For

            End If

        Next j

        If foundMatch = False Then

            wsBooks.Cells(i, 4).Value = "NOT FOUND IN 2B"
            wsBooks.Cells(i, 4).Interior.Color = RGB(255, 0, 0)

        End If

    Next i

    MsgBox "GST Reconciliation Completed!"

End Sub

