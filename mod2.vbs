Attribute VB_Name = "Module1"
Sub Worksheetloop()

    Dim ws As Worksheet
    Dim Need_Summary_Table_Header As Boolean
    Dim COMMAND_SPREADSHEET As Boolean
    
    Need_Summary_Table_Header = False
    COMMAND_SPREADSHEET = True
    
' Loop through all of the worksheets in workbook
    For Each ws In Worksheets
    
' Set variable for holding the ticker name
        Dim Ticker_Name As String
        Ticker_Name = " "
        
        Dim Total_Ticker_Volume As Double
        Total_Ticker_Volume = 0
        
' Set variables for Moderate Solution bit
        Dim Open_Price As Double
        Open_Price = 0
        Dim Close_Price As Double
        Close_Price = 0
        Dim Yearly_Change As Double
        Yearly_Change = 0
        Dim Yearly_percent As Double
        Yearly_percent = 0
' Set variables for Hard Solution bit
        Dim MAX_TICKER_NAME As String
        MAX_TICKER_NAME = " "
        Dim MIN_TICKER_NAME As String
        MIN_TICKER_NAME = " "
        Dim MAX_PERCENT As Double
        MAX_PERCENT = 0
        Dim MIN_PERCENT As Double
        MIN_PERCENT = 0
        Dim MAX_VOLUME_TICKER As String
        MAX_VOLUME_TICKER = " "
        Dim MAX_VOLUME As Double
        MAX_VOLUME = 0
'---------------------

' in the summary table for the current worksheet
        Dim Summary_Table_Row As Long
        Summary_Table_Row = 2
        
' Set row count for the current worksheet
        Dim Lastrow As Long
        Dim i As Long
        
        Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

' For all ws except the first
        If Need_Summary_Table_Header Then
' Set title names for the Summary Table for current ws
            ws.Range("I1").Value = "Ticker"
            ws.Range("J1").Value = "Yearly Change"
            ws.Range("K1").Value = "Percent Change"
            ws.Range("L1").Value = "Total Stock Volume"
' Set Titles for new Summary Table on the right for ws
            ws.Range("O2").Value = "Greatest % Increase"
            ws.Range("O3").Value = "Greatest % Decrease"
            ws.Range("O4").Value = "Greatest Total Volume"
            ws.Range("P1").Value = "Ticker"
            ws.Range("Q1").Value = "Value"
        Else
            Need_Summary_Table_Header = True
        End If
'---------------------
'for the first ws
            ws.Range("I1").Value = "Ticker"
            ws.Range("J1").Value = "Yearly Change"
            ws.Range("K1").Value = "Percent Change"
            ws.Range("L1").Value = "Total Stock Volume"

            ws.Range("O2").Value = "Greatest % Increase"
            ws.Range("O3").Value = "Greatest % Decrease"
            ws.Range("O4").Value = "Greatest Total Volume"
            ws.Range("P1").Value = "Ticker"
            ws.Range("Q1").Value = "Value"
'---------------------

' Set value of Open Price for the first Ticker of ws
'open price for loop below
        Open_Price = ws.Cells(2, 3).Value
        
' Loop from the beginning of the current worksheet(Row2) till its last row
        For i = 2 To Lastrow
' Check if we are still within the same ticker name
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
' Set the ticker name, we are ready to insert this ticker name data
                Ticker_Name = ws.Cells(i, 1).Value
                
' Calculate Yearly_Change and Yearly_percent
                Close_Price = ws.Cells(i, 6).Value
                Yearly_Change = Close_Price - Open_Price
' Check divi by 0 condition
                If Open_Price <> 0 Then
                    Yearly_percent = (Yearly_Change / Open_Price) * 100
                Else
                    MsgBox ("For " & Ticker_Name & ", Row " & CStr(i) & ": Open Price =" & Open_Price & ". Fix <open> field manually and save the spreadsheet.")
                End If
                
' Add to the Ticker name total volume
                Total_Ticker_Volume = Total_Ticker_Volume + ws.Cells(i, 7).Value
 '-----------------------------------------
            
                ws.Range("I" & Summary_Table_Row).Value = Ticker_Name
                ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
 ' Fill yearly_change to grn red
                If (Yearly_Change > 0) Then
'Fill column with grn color
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                ElseIf (Yearly_Change <= 0) Then
'Fill column with red color
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                End If
'-----------------------------------------
                ws.Range("K" & Summary_Table_Row).Value = (CStr(Yearly_percent) & "%")
                ws.Range("L" & Summary_Table_Row).Value = Total_Ticker_Volume
' Add 1 to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
                Yearly_Change = 0
' Hard solution do this in top loop Yearly_percent = 0
                Close_Price = 0
                Open_Price = ws.Cells(i + 1, 3).Value

                If (Yearly_percent > MAX_PERCENT) Then
                    MAX_PERCENT = Yearly_percent
                    MAX_TICKER_NAME = Ticker_Name
                ElseIf (Yearly_percent < MIN_PERCENT) Then
                    MIN_PERCENT = Yearly_percent
                    MIN_TICKER_NAME = Ticker_Name
                End If
                       
                If (Total_Ticker_Volume > MAX_VOLUME) Then
                    MAX_VOLUME = Total_Ticker_Volume
                    MAX_VOLUME_TICKER = Ticker_Name
                End If
                
' Hard solution adjustments to resetting counters
                Yearly_percent = 0
                Total_Ticker_Volume = 0
                      
'Else - If the cell following a row is still the same ticker name
            Else
                Total_Ticker_Volume = Total_Ticker_Volume + ws.Cells(i, 7).Value
            End If
      
        Next i

 ' put all new counts to the new summary table on the right of spreadsheet
            If Not COMMAND_SPREADSHEET Then
            
                ws.Range("Q2").Value = (CStr(MAX_PERCENT) & "%")
                ws.Range("Q3").Value = (CStr(MIN_PERCENT) & "%")
                ws.Range("P2").Value = MAX_TICKER_NAME
                ws.Range("P3").Value = MIN_TICKER_NAME
                ws.Range("Q4").Value = MAX_VOLUME
                ws.Range("P4").Value = MAX_VOLUME_TICKER
                
            Else
                COMMAND_SPREADSHEET = False
            End If
        
     Next ws
    
End Sub

