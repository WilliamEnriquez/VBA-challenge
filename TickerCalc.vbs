Attribute VB_Name = "TickerCalc"
Option Explicit

Sub Reference_Worksheet()
' This sections sets the headers on each worksheet
    Dim Sh As Worksheet
    Set Sh = ActiveSheet
    Dim SheetCount As Integer
    Dim LastRow As Long, i As Long
    Dim SPrice(1 To 4) As Double, SVol(1 To 2) As Double
    Dim myRange As Range
    Dim myValue As Double
    Dim Summary_Table_Row As Long
    Dim LastSheet As Byte
    Dim StkSymbl As String
    Dim Space As String
    Dim OpenRow As Long, rwc As Long
    Space = " "
    Let LastSheet = Sheets.Count
    
    SheetCount = 1
   For SheetCount = 1 To LastSheet
   
   Summary_Table_Row = 2
   
        Worksheets(SheetCount).Select
        Columns("K:K").Select
        Selection.NumberFormat = "0.00%"
        Worksheets(SheetCount).Range("H1").Value = "<vol/K>"
        Worksheets(SheetCount).Range("I1").Value = "Ticker"
        Worksheets(SheetCount).Range("J1").Value = "Yearly Change"
        Worksheets(SheetCount).Range("K1").Value = "Percent Change"
        Worksheets(SheetCount).Range("L1").Value = "Stock Volume"
        Worksheets(SheetCount).Range("P1").Value = "Ticker"
        Worksheets(SheetCount).Range("Q1").Value = "Value"
        Worksheets(SheetCount).Range("O2").Value = "Greatest % Increase"
        Range("O2:Q2").Interior.Color = VBA.vbGreen
        Worksheets(SheetCount).Range("O3").Value = "Greatest % Decrease"
        Range("O3:Q3").Interior.Color = VBA.vbRed
        Worksheets(SheetCount).Range("O4").Value = "Greatest Total Volume/K"
        
        ' Set myRange = Range("A2", "H" & Cells(Rows.Count, 1).End(xlUp).Row)
        Let LastRow = ActiveSheet.UsedRange.Cells(ActiveSheet.UsedRange.Rows.Count, 1).Row
        ' LastRow = Worksheets(SheetCount).UsedRange.Cells(ActiveSheet.UsedRange.Rows.Count, 1).Row
        
        ' Reset my variables to zero
        ' SVol(2) will store my highest volume on the sheet
        SVol(2) = 0
        SPrice(3) = 0
        SPrice(4) = 0
        OpenRow = 2
            For i = 2 To LastRow
            ' Cut the volume figure by 1000for faster computing
                Range("H" & i).Value = Range("G" & i).Value / 1000
                
            ' Set open annual stock price and
            ' Check if we are still within the same ticker symbol

                SPrice(1) = Cells(OpenRow, 3).Value
                SVol(1) = Cells(i, 8).Value
            If Left(Cells(i + 1, 2), 4) <> Left(Cells(i, 2), 4) Or _
                Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                SPrice(2) = Cells(i, 6).Value
                                                
                ' Set the stock symbol
                    StkSymbl = Cells(i, 1).Value
                    Range("I" & Summary_Table_Row).Value = StkSymbl '& Space & Left(Cells(i, 2), 4)
                ' Set the annual price change
                If SPrice(2) - SPrice(1) <= 0 Then
                    Range("J" & Summary_Table_Row, "K" & Summary_Table_Row).Interior.Color = VBA.vbRed
                Else
                    Range("J" & Summary_Table_Row, "K" & Summary_Table_Row).Interior.Color = VBA.vbGreen
                End If
                
                    Range("J" & Summary_Table_Row).Value = SPrice(2) - SPrice(1)
                                
                ' Calculate percent change
                    If SPrice(1) <> 0 Then
                        Range("K" & Summary_Table_Row).Value = (Range("J" & Summary_Table_Row).Value / SPrice(1)) * 100
                    Else
                        Range("K" & Summary_Table_Row).Value = 0
                    End If
                ' Push Annual volume total
                    Range("L" & Summary_Table_Row).Value = SVol(1) * 1000
                ' add one to thesummary table row
                    Summary_Table_Row = Summary_Table_Row + 1
                    Else
                    SVol(1) = SVol(1) + Cells(i, 8).Value
                    OpenRow = i + 1
            End If
                    
                'Find Greatest percent increase per sheet
                    If Range("K" & Summary_Table_Row).Value > SPrice(4) Then
                        SPrice(4) = Range("K" & Summary_Table_Row).Value
                    End If
                    
                 'Find Greatest percent decrese per sheet
                    If Range("K" & Summary_Table_Row).Value < SPrice(3) Then
                        SPrice(3) = Range("K" & Summary_Table_Row).Value
                    End If
                    
                'Find highest volume per sheet
                    If SVol(2) < SVol(1) Then
                        SVol(2) = SVol(1)
                    End If
                ' Post the highest volume
                Range("Q2").Value = SPrice(4)
                Range("Q3").Value = SPrice(3)
                Range("Q2:Q3").NumberFormat = "0.00%"
                Range("Q4").Value = SVol(2) * 1000
            
            Next i
        Columns("A:Q").Select
        Selection.Columns.AutoFit
        Range("M1").Select
    Next SheetCount
    
    ' Debug.Print ActiveWorkbook.FullName
    ' Debug.Print myRange.Address
        
End Sub
