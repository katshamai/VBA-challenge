Attribute VB_Name = "Module1"
Sub TickerAnalysis()
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call TickerAnalysis1
    Next
    Application.ScreenUpdating = True
    
End Sub

Private Sub TickerAnalysis1()

               ' Create a script to loop through ticker symbol
                Dim ticker As String
                Dim yearlychange As Double
                Dim percentagechange As Double
                Dim totalstock As Double
                Dim openprice As Double
                Dim closeprice As Double
                Dim grtperinc As Double
                Dim grtperdec As Double
                Dim grttotval As Double
        
        
                ' Count the number of rows in the data set
                Dim last_row As Long
                last_row = Cells(Rows.Count, 1).End(xlUp).Row
                
                ' Keep track of the data for each ticker in the summary table
                Dim Summary_Table_Row As Integer
                Summary_Table_Row = 2
                Summary_Table_Header = 1
                  
                ' Keep track of greatest data for table
                Dim Greatest_Table_Row As Integer
                Greatest_Table_Row = 2
                Greatest_Table_Header = 1
                  
                'Set up Summary Table Headers
                  Range("I" & Summary_Table_Header).Value = "Ticker"
                  Range("I" & Summary_Table_Header).Font.FontStyle = "Bold"
                  Range("I" & Summary_Table_Header).HorizontalAlignment = xlCenter
                  Range("J" & Summary_Table_Header).Value = "Yearly Change"
                  Range("J" & Summary_Table_Header).Font.FontStyle = "Bold"
                  Range("J" & Summary_Table_Header).HorizontalAlignment = xlCenter
                  Range("K" & Summary_Table_Header).Value = "Percent Change"
                  Range("K" & Summary_Table_Header).Font.FontStyle = "Bold"
                  Range("K" & Summary_Table_Header).HorizontalAlignment = xlCenter
                  Range("L" & Summary_Table_Header).Value = "Total Stock Volume"
                  Range("L" & Summary_Table_Header).Font.FontStyle = "Bold"
                  Range("L" & Summary_Table_Header).HorizontalAlignment = xlCenter
                
                ' Set opening price
                    openprice = Cells(2, 3)
                                                                                  
                ' Loop through all ticker
                  For i = 2 To last_row
                                                                      
                    ' Check if we are still within the ticker
                   If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                                      
                      ' Set the Ticker
                      ticker = Cells(i, 1).Value
                                          
                      ' Total Stock
                      totalstock = totalstock + Cells(i, 7).Value
                                            
                      ' Identify close price
                      closeprice = Cells(i, 6).Value
                                                              
                      ' Calculate yearly change
                      yearlychange = closeprice - openprice
                                                                                                                      
                      'Calculate percentage change
                      If openprice <> 0 Then
                      percentagechange = yearlychange / openprice
                      Else
                      percentagechange = 0 'on the assumption that opening price is based on value at "the beginning of a given year" but not the first available opening price value
                      End If
                                                                                               
                      ' Print results
                      Range("I" & Summary_Table_Row).Value = ticker
                      Range("J" & Summary_Table_Row).Value = yearlychange
                      Range("K" & Summary_Table_Row).Value = percentagechange
                      Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                      Range("L" & Summary_Table_Row).Value = totalstock
                      
                      ' Add one to the summary table row
                      Summary_Table_Row = Summary_Table_Row + 1
                                            
                      ' Identify opening price
                                                                                   
                      openprice = Cells(i + 1, 3).Value
                      
                      
                                                                                                                         
                      ' Reset total stock
                      totalstock = 0
                      
                      ' If the cell immediately following a row is the same ticker
                    
                    Else
                         
                      ' Add to the Stock Total
                      totalstock = totalstock + Cells(i, 7).Value
                                                    
                    End If
                   
                    Next i
                
                        'Conditional Formatting for Yearly Change
                        
                        For i = 2 To last_row
                        
                        ' Check if the change is positive
                          If Cells(i, 10).Value >= 0 Then
                        
                              ' Color the positive change green
                              Cells(i, 10).Interior.ColorIndex = 4
                        
                              
                          ' Check if the change is negative
                          ElseIf Cells(i, 10).Value <= 0 Then
                        
                              ' Color the negative change red
                              Cells(i, 10).Interior.ColorIndex = 3
                        
                          End If
                        
                        Next i
                        
                    'Set up Greatest Table
                    Range("O" & Greatest_Table_Header).Value = "Ticker"
                    Range("O" & Greatest_Table_Header).Font.FontStyle = "Bold"
                    Range("O" & Greatest_Table_Header).HorizontalAlignment = xlCenter
                    Range("P" & Greatest_Table_Header).Value = "Value"
                    Range("P" & Greatest_Table_Header).Font.FontStyle = "Bold"
                    Range("P" & Greatest_Table_Header).HorizontalAlignment = xlCenter
                    Range("N" & Greatest_Table_Row).Value = "Greatest % Increase"
                    Range("N" & Greatest_Table_Row).Font.FontStyle = "Bold"
                    Range("N" & Greatest_Table_Row).HorizontalAlignment = xlCenter
                    Greatest_Table_Row = Greatest_Table_Row + 1
                    Range("N" & Greatest_Table_Row).Value = "Greatest % Decrease"
                    Range("N" & Greatest_Table_Row).Font.FontStyle = "Bold"
                    Range("N" & Greatest_Table_Row).HorizontalAlignment = xlCenter
                    Greatest_Table_Row = Greatest_Table_Row + 1
                    Range("N" & Greatest_Table_Row).Value = "Greatest Total Volume"
                    Range("N" & Greatest_Table_Row).Font.FontStyle = "Bold"
                    Range("N" & Greatest_Table_Row).HorizontalAlignment = xlCenter
                          
                    Dim last_row2 As Long
                    last_row2 = Cells(Rows.Count, 12).End(xlUp).Row
                                                                                                                                       
                    For K = 2 To last_row2
                    
                    If Cells(K + 1, 11) > Max Then
                    Max = Cells(K + 1, 11)
                    Tag = Cells(K + 1, 9)
                    End If
                    
                    Cells(2, "P").Value = Max
                    Cells(2, "P").NumberFormat = "0.00%"
                    Cells(2, "O").Value = Tag
                    
                    If Cells(K + 1, 11) < Min Then
                    Min = Cells(K + 1, 11)
                    tagmin = Cells(K + 1, 9)
                    End If
                    
                    Cells(3, "P").Value = Min
                    Cells(3, "P").NumberFormat = "0.00%"
                    Cells(3, "O").Value = tagmin
                    
                    If Cells(K + 1, 12) > Maxvol Then
                    Maxvol = Cells(K + 1, 12)
                    tagvol = Cells(K + 1, 9)
                    End If
                    
                    Cells(4, "P").Value = Maxvol
                    Cells(4, "O").Value = tagvol
                    
                    Next K
                    
                                                                                                                                
                    ' Format columns to fit content
                    ActiveSheet.Columns("A:T").AutoFit
            
End Sub


