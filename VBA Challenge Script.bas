Attribute VB_Name = "Module1"
Sub stock_homework()

' Turn off screen updating and automatic calculations
Application.Calculation = xlManual
Application.ScreenUpdating = False

' Set Ws as Worksheet
Dim Ws As Worksheet
Dim Need_Table_Header As Boolean
Dim Spreadsheet As Boolean

Need_Table_Header = True
Spreadsheet = True

' Loop Through worksheets in the workbook

For Each Ws In Worksheets
    
' Defining Variables

    Dim Ticker As String
    Ticker = " "
    Dim Change_In_Price As Double
    Dim Percent_Change As Double
    Dim Total_Stock As LongLong
    Dim Lastrow As Long
    Dim i As Long
    Dim Open_Price As Double
    Dim Close_Price As Double
  


    Lastrow = Ws.Cells(Rows.Count, 1).End(xlUp).Row

    Total_Stock = 0
    Open_Price = 0
    Close_Price = 0
    Change_In_Price = 0
    Percent_Change = 0
   

    ' Output Table
    Dim Output_Table As Integer

    Output_Table = 2

    ' Create Table Headers
    If Need_Table_Header Then
    
        Ws.Range("J1").Value = ("Ticker")
        Ws.Range("K1").Value = ("Yearly Change")
        Ws.Range("L1").Value = ("Percent Change")
        Ws.Range("M1").Value = ("Total Stock Volume")
        ' Additional Headers for Bonus
        Ws.Range("P2").Value = ("Greatest % Increase")
        Ws.Range("P3").Value = "Greatest % Decrease"
        Ws.Range("P4").Value = ("Greatest Total Volume")
        Ws.Range("Q1").Value = ("Ticker")
        Ws.Range("R1").Value = ("Value")
    Else
        
        Need_Table_Header = True
    End If


' Setting initial value for Open Price
Open_Price = Ws.Range("C2")

' Loop through all Tickers
For i = 2 To Lastrow
    If Ws.Cells(i + 1, 1).Value <> Ws.Cells(i, 1).Value Then
    
        ' Setting Ticker Name
        Ticker = Ws.Cells(i, 1).Value
        
        'Calculate Change in Price and Percent Change
        Close_Price = Ws.Cells(i, 6).Value
        Change_In_Price = Close_Price - Open_Price
        
        If Open_Price <> 0 Then
            Percent_Change = (Change_In_Price / Open_Price)
            
        Else
            Percent_Change = 0
            
        End If
    
        
        ' Add to the Stock Volume total
        Total_Stock = Total_Stock + Cells(i, 7).Value
        
        ' Print Stock Ticker Name in Table
        Ws.Range("J" & Output_Table).Value = Ticker
        
        ' Print Yearly Change
        Ws.Range("K" & Output_Table).Value = Change_In_Price
        Ws.Range("K" & Output_Table).NumberFormat = "0.00"
        
        ' Print Percent Change
        Ws.Range("L" & Output_Table).Value = Percent_Change
        Ws.Range("L" & Output_Table).NumberFormat = "0.00%"
    
        'Print Stock Total in Table
        Ws.Range("M" & Output_Table).Value = Total_Stock
        
        If Ws.Range("K" & Output_Table).Value >= 0 Then
            Ws.Range("K" & Output_Table).Interior.ColorIndex = 4
        Else
            Ws.Range("K" & Output_Table).Interior.ColorIndex = 3
        End If
        
        
        Output_Table = Output_Table + 1
        
        ' Reset Stock Total
        
        Total_Stock = 0
        
    Else
    
        Total_Stock = Total_Stock + Cells(i, 7).Value

      
        
    End If
    
Next i

' Bonus: Greatest % increase, Greatest % Decrease, and Greatest Total Volume
Lastrow = Ws.Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To Lastrow
    If Ws.Range("L" & i).Value > Ws.Range("R2").Value Then
    Ws.Range("R2").Value = Ws.Range("L" & i).Value
    Ws.Range("Q2").Value = Ws.Range("J" & i).Value
    
    End If
    
    If Ws.Range("L" & i).Value < Ws.Range("R3").Value Then
        Ws.Range("R3").Value = Ws.Range("L" & i).Value
        Ws.Range("Q3").Value = Ws.Range("J" & i).Value
    End If
    
    If Ws.Range("M" & i).Value > Ws.Range("R4").Value Then
        Ws.Range("R4").Value = Ws.Range("M" & i).Value
        Ws.Range("Q4").Value = Ws.Range("J" & i).Value
    End If
Next i

' Format to include %
Ws.Range("R2").NumberFormat = "0.00%"
Ws.Range("R3").NumberFormat = "0.00%"
    
' Turn on Screen Updating and Automatic Calculations
Application.Calculation = xlAutomatic
Application.ScreenUpdating = True


Ws.Columns("J:R").AutoFit

Next Ws

End Sub
