Attribute VB_Name = "Module1"
'UCBC HW 2 VBA
Sub HW2_VBA()


'Creating the variables
Dim ticker_name As String
Dim ticker_value As Double
Dim yearly_start As Double
Dim yearly_end As Double




Dim sum_table As Integer
    sum_table = 2

'Looping through the stock data
For i = 2 To 70926

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        ticker_name = Cells(i, 1).Value
        ticker_value = Cells(i, 7).Value + ticker_value
        
        'Printing out the ticker name and value
         Range("H" & sum_table).Value = ticker_name
         Range("I" & sum_table).Value = ticker_value
         
         
         'Moving the sum_table down a cell
         sum_table = sum_table + 1
         
         
         'Reset ticker_value
         ticker_value = 0
    
    'Adding up the total stock volume
    Else
        
        ticker_value = Cells(i, 7).Value + ticker_value
        
    End If
    
Next i

End Sub



        
        
        

        
    

