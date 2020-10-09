# VBA-Challenge

Sub Stock_Market()

For Each ws In Worksheets

        ' Created a Variable to Hold File Name, Last Row, Last Column, and Year

        Dim Ticker_Name As String
        Ticker_Name = ""
        Dim Total_Stock_Volume As Double
        Dim Begin_Value As Double
        Dim Year_End_Value As Double
        Dim Yearly_Change As Double
        Dim Open_Price As Double
        Dim Close_Price As Double
        Dim Percentage_change As Double
        Dim Summary_Table_Row As Integer
        Dim LastRow As Long
        Dim WorkSheetName As String
        
 
        Summary_Table_Row = 2
        Total_Stock_Volume = 0
        WorkSheetName = ws.Name
      
        
        'Determine the Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
            ws.Range("I1").Value = "Ticker"
            ws.Range("J1").Value = "Yearly_Change"
            ws.Range("K1").Value = "Percentage_Change"
            ws.Range("L1").Value = "Total_Stock_Volume"
        
        Begin_Value = ws.Cells(2, 3).Value

        For i = 2 To LastRow
        
           
        
            ' Differentiate between ticker name & take Begin value of current ticker
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                Ticker_Name = ws.Cells(i, 1).Value
                ws.Cells(1, 9).Value = "Ticker"
                ws.Cells(1, 10).Value = "Yearly_Change"
                ws.Cells(1, 11).Value = "Percentage Change"
                ws.Cells(i, 3).Value = "Begin_Value"
                ws.Cells(i, 5).Value = "Year_End_Value"
                ws.Cells(i, 12).Value = "Total Stock Volume"
                
                Total_Stock_Volume = Total_Stock_Volume + CDbl(ws.Cells(i, 7).Value)
                Ticker_Name = ws.Cells(i, 1).Value
                Year_End_Value = ws.Cells(i, 6).Value
                Yearly_Change = (Year_End_Value - Begin_Value)
                
                'print the TIcker Name in the Summary Table
                ws.Range("I" & Summary_Table_Row).Value = Ticker_Name
                
                'To avoide rows with value Zeroes
                    If (Begin_Value > 0) Then
                        Percentage_change = Yearly_Change / Begin_Value
                    
                    End If
                
                
                If (Yearly_Change >= 0) Then
                
                ' Print the Yearly Change to the Summary Table
                ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                
                Else
                ' Print the Yearly Change to the Summary Table
                ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                
                End If
                
                        
                ' Print Percentage Change to the Summary Table
                ws.Range("K" & Summary_Table_Row).Value = Percentage_change
                
                ' Print the Total Volume to the Summary Table
                ws.Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
                
                'Format change for Percentage change column
        
                ws.Range("K" & Summary_Table_Row).NumberFormat = "##,##0.00%"

                Summary_Table_Row = Summary_Table_Row + 1

            Total_Stock_Volume = 0
            Percentage_change = 0

            End If
        
        Next i
        
    Next ws
    
    End Sub
