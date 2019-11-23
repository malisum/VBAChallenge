Attribute VB_Name = "TickerData"
Option Explicit


Sub analyzeTickerData()


' ------------------------------------------------------------------------------------------------------------------------------- '
'
' Date: 22 Nov 2019
' Procedure Log:
' This procedure read the data in the 3 worksheets within the current workbook.
' Each sheet has data for a specific year:
' Sheet1 "2014" is for 2014 data
' Sheet2 "2015" is for 2015 data
' Sheet3 "2016" is for 2016 data
'
' Each sheet has data for stock prices and volume traded for a date within the year.
' The procdure will analyze the yearly data per sheet and provide 2 outputs:
' 1. Summary of each stock ticker in columns J through M (column number 9 through 12)
'    Column J (9): Stock Symbol
'    Column K (10): Yearly change from opening price at the beginning of a given year to the closing price at the end of that year
'    Column L (11): Percent change from opening price at the beginning of a given year to the closing price at the end of that year
'    Column M (12): The total stock volume of the stock.
' 2. The stock with the Greatest % increase, Greatest % Decrease and Greatest total volume
'    Column R (15): Label
'    Column P (16): Ticker
'    Column Q (17): Volume
'
' Note: The list provided in the sample data iss sorted, but the procedure is designed to work on an unsorted list too.
'       For performance (when the list is unsorted), the logic will handle this by adding 4 columns for the 1st summary table,
'       which will record the lowest & highest date per stock and the corresponding price.
'       When a match is found in the 1st summary, lowest & highest date and corresponding price is be maintained as
'       part of the list. Also with each iteration of 1st summary, the 2nd summary is updated with the highest
'       percent increase, highest percent decrease & highest total volume (1st summary has highest & lowes date data per stock).
'       For perforamce, second loop to search in 1st summary will check if the stock ticker is same as the last one processed,
'       if same then it will use it & skip the second loop on 1st summary. If different, it will search in the for loop.
'       At the end of the process, the 4 temporary columns will be deleted & the columns for the 2 summary will be formatted.
'
' ------------------------------------------------------------------------------------------------------------------------------- '


' Decalre variables:
' Variable for Worksheet object
Dim currentSheet As Worksheet
' Last data row in the Worksheet
Dim lastDataRow As Long
' Last summary row in the Worksheet
Dim lastSmryRow As Long
' Last data column in the Worksheet
Dim lastDataCol As Integer
' Counter variable for main data loop
Dim dataCounter As Long
' Counter variable for 1st summary table
Dim smryCounter As Long
' Ticket Found Flag
Dim isTickerFound As Boolean
' Summary count
Dim smryCount As Long
' prevSmrycounter (for faster search if ticker is same as last one read)
Dim prevSmrycounter As Long
' Simple loopCounter
Dim loopCounter As Integer

' Initialize any data common to all sheets


' Main logic:
' Repeat for each sheet in the Workbook
For Each currentSheet In ThisWorkbook.Worksheets


    ' Initialize / Set (per sheet)
    
    ' Delete any data from prior macro run (Columns 10 to 26) - this will clear any custom number formatting also
    For loopCounter = 26 To 10 Step -1
        currentSheet.Columns(loopCounter).Delete
    Next loopCounter
        
    ' Set Last Data Row
    lastDataRow = currentSheet.Cells(Rows.Count, 1).End(xlUp).Row
    ' Set Last Column Row (Use the last row for, otherwise the 1st row will get summary table on the right side -
    ' - which will give wrong answer if the macro is run more than once
    lastDataCol = currentSheet.Cells(lastDataRow, Columns.Count).End(xlToLeft).Column
        
    
    ' Populate summary table header data (leave 2 blanks columns from last data column)
    currentSheet.Cells(1, 10).Value = "Ticker"
    currentSheet.Cells(1, 11).Value = "Yearly Change"
    currentSheet.Cells(1, 12).Value = "Percent Change"
    currentSheet.Cells(1, 13).Value = "Total Stock Volume"
    ' Temporary columns - store begining & ending year date & price (if data is not sorted)
    ' These 4 columns will be deleted at the end
    currentSheet.Cells(1, 14).Value = "Lowest Date"
    currentSheet.Cells(1, 15).Value = "Price at Lowest Date"
    currentSheet.Cells(1, 16).Value = "Highest Date"
    currentSheet.Cells(1, 17).Value = "Price at Highest Date"
    
    
    ' Populate summary All table header data & label (leave 2 columns and 6 columns for summary from data column)
    currentSheet.Cells(1, 20).Value = " "
    currentSheet.Cells(1, 21).Value = "Ticker"
    currentSheet.Cells(1, 22).Value = "Value"
    currentSheet.Cells(2, 20).Value = "Greatest % increase"
    currentSheet.Cells(3, 20).Value = "Greatest % Decrease"
    currentSheet.Cells(4, 20).Value = "Greatest total volume"

    
    ' Inititialze summary total = 0
    smryCount = 0
    ' Intialize saved summary counter = 0
    prevSmrycounter = 0
    
    ' Initialize final totals for second summary table (per sheet)
    currentSheet.Cells(2, 21).Value = ""
    currentSheet.Cells(3, 21).Value = ""
    currentSheet.Cells(4, 21).Value = ""
    currentSheet.Cells(2, 22).Value = ""
    currentSheet.Cells(3, 22).Value = ""
    currentSheet.Cells(4, 22).Value = ""


    ' Data is sorted but assume & code if it is not sorted for comprehensive logic
    ' Read the data in the current sheet
    For dataCounter = 2 To lastDataRow
    
        ' Each Row (dataCounter)
        ' Start with no match found for summary
        isTickerFound = False
            
        ' Try the last read (saved) summary (to avoid second loop if it is for same ticker symbol as last one)
        If prevSmrycounter <> 0 Then
            If currentSheet.Cells(dataCounter, 1).Value = currentSheet.Cells(prevSmrycounter, 10).Value Then
                isTickerFound = True
                ' Match found - Add to total for the symbol
                currentSheet.Cells(prevSmrycounter, 13).Value = currentSheet.Cells(prevSmrycounter, 13).Value + currentSheet.Cells(dataCounter, 7).Value
                ' Check & replace lowest date and price for the match found
                If currentSheet.Cells(dataCounter, 2).Value < currentSheet.Cells(prevSmrycounter, 14).Value Then
                    ' Replace lowest date & opening value
                    currentSheet.Cells(prevSmrycounter, 14).Value = currentSheet.Cells(dataCounter, 2).Value
                    currentSheet.Cells(prevSmrycounter, 15).Value = currentSheet.Cells(dataCounter, 3).Value
                    ' Recalculate yearly change and percentage change (date data changed)
                    ' Yearly change = Highest date value (col 17) - Lowest Date value (col 15)
                    currentSheet.Cells(prevSmrycounter, 11).Value = currentSheet.Cells(prevSmrycounter, 17).Value - currentSheet.Cells(prevSmrycounter, 15).Value
                    ' Percentage change  = Yearly change / Lowest Date value (col 15)
                    ' Avoid divide by zero
                    If currentSheet.Cells(prevSmrycounter, 15).Value = 0 Then
                        currentSheet.Cells(prevSmrycounter, 12).Value = 0
                    Else
                        currentSheet.Cells(prevSmrycounter, 12).Value = currentSheet.Cells(prevSmrycounter, 11).Value / currentSheet.Cells(prevSmrycounter, 15).Value
                    End If
                End If
                ' Check & replace highest date and price for the match found
                If currentSheet.Cells(dataCounter, 2).Value > currentSheet.Cells(prevSmrycounter, 16).Value Then
                    ' Replace highest date & closing value
                    currentSheet.Cells(prevSmrycounter, 16).Value = currentSheet.Cells(dataCounter, 2).Value
                    currentSheet.Cells(prevSmrycounter, 17).Value = currentSheet.Cells(dataCounter, 6).Value
                    ' Recalculate yearly change and percentage change (date data changed)
                    ' Yearly change = Highest date value (col 17) - Lowest Date value (col 15)
                    currentSheet.Cells(prevSmrycounter, 11).Value = currentSheet.Cells(prevSmrycounter, 17).Value - currentSheet.Cells(prevSmrycounter, 15).Value
                    ' Percentage change  = Yearly change / Lowest Date value (col 15)
                    ' Avoid divide by zero
                    If currentSheet.Cells(prevSmrycounter, 15).Value = 0 Then
                        currentSheet.Cells(prevSmrycounter, 12).Value = 0
                    Else
                        currentSheet.Cells(prevSmrycounter, 12).Value = currentSheet.Cells(prevSmrycounter, 11).Value / currentSheet.Cells(prevSmrycounter, 15).Value
                    End If
                End If
            End If ' Not same as prvious summary record
        End If ' zero previous summary counter
                
        ' Proceed with the loop to check in summary
        ' if this was not same as the last summary checked above And there is atleast one summary record
        If isTickerFound = False And smryCount >= 1 Then
            ' Loop through to chck if it exists in Suummary
            For smryCounter = 2 To (smryCount + 2)
                If currentSheet.Cells(dataCounter, 1).Value = currentSheet.Cells(smryCounter, 10).Value Then
                    isTickerFound = True
                    prevSmrycounter = smryCounter
                    ' Match found - Add to total for the symbol
                    currentSheet.Cells(smryCounter, 13).Value = currentSheet.Cells(smryCounter, 13).Value + currentSheet.Cells(dataCounter, 7).Value
                   ' Check & replace lowest date and price for the match found
                    If currentSheet.Cells(dataCounter, 2).Value < currentSheet.Cells(smryCounter, 14).Value Then
                        ' Replace lowest date & opening value
                        currentSheet.Cells(smryCounter, 14).Value = currentSheet.Cells(dataCounter, 2).Value
                        currentSheet.Cells(smryCounter, 15).Value = currentSheet.Cells(dataCounter, 3).Value
                        ' Recalculate yearly change and percentage change (date data changed)
                        ' Yearly change = Highest date value (col 17) - Lowest Date value (col 15)
                        currentSheet.Cells(smryCounter, 11).Value = currentSheet.Cells(smryCounter, 17).Value - currentSheet.Cells(smryCounter, 15).Value
                        ' Percentage change  = Yearly change / Lowest Date value (col 15)
                        ' Avoid divide by zero
                        If currentSheet.Cells(prevSmrycounter, 15).Value = 0 Then
                            currentSheet.Cells(prevSmrycounter, 12).Value = 0
                        Else
                            currentSheet.Cells(prevSmrycounter, 12).Value = currentSheet.Cells(prevSmrycounter, 11).Value / currentSheet.Cells(prevSmrycounter, 15).Value
                        End If
                    End If
                    ' Check & replace highest date and price for the match found
                    If currentSheet.Cells(dataCounter, 2).Value > currentSheet.Cells(smryCounter, 16).Value Then
                        ' Replace highest date & closing value
                        currentSheet.Cells(smryCounter, 16).Value = currentSheet.Cells(dataCounter, 2).Value
                        currentSheet.Cells(smryCounter, 17).Value = currentSheet.Cells(dataCounter, 6).Value
                        ' Recalculate yearly change and percentage change (date data changed)
                        ' Yearly change = Highest date value (col 17) - Lowest Date value (col 15)
                        currentSheet.Cells(smryCounter, 11).Value = currentSheet.Cells(smryCounter, 17).Value - currentSheet.Cells(smryCounter, 15).Value
                        ' Percentage change  = Yearly change / Lowest Date value (col 15)
                        ' Avoid divide by zero
                        If currentSheet.Cells(prevSmrycounter, 15).Value = 0 Then
                            currentSheet.Cells(prevSmrycounter, 12).Value = 0
                        Else
                            currentSheet.Cells(prevSmrycounter, 12).Value = currentSheet.Cells(prevSmrycounter, 11).Value / currentSheet.Cells(prevSmrycounter, 15).Value
                        End If
                    End If
                    ' Exit the for loop if match found
                    Exit For
                End If
            ' Next summary Row
            Next smryCounter
        End If
        
        
        ' If no match found - Add one summary row
        If isTickerFound = False Then
            
            ' Increment summary count by 1 for new ticker symbol
            smryCount = smryCount + 1
            ' Ticker symbol
            currentSheet.Cells(smryCount + 1, 10).Value = currentSheet.Cells(dataCounter, 1).Value
            ' Total volume
            currentSheet.Cells(smryCount + 1, 13).Value = currentSheet.Cells(dataCounter, 7).Value
            ' Lowest date for the year this symbol
            currentSheet.Cells(smryCount + 1, 14).Value = currentSheet.Cells(dataCounter, 2).Value
            ' Opening price on lowes date for this year this symbol
            currentSheet.Cells(smryCount + 1, 15).Value = currentSheet.Cells(dataCounter, 3).Value
            ' Highest date for the year this symbol
            currentSheet.Cells(smryCount + 1, 16).Value = currentSheet.Cells(dataCounter, 2).Value
            ' Closing price on lowes date for this year this symbol
            currentSheet.Cells(smryCount + 1, 17).Value = currentSheet.Cells(dataCounter, 6).Value
            ' Recalculate yearly change and percentage change (date data changed)
            ' Yearly change = Highest date value (col 17) - Lowest Date value (col 15)
            currentSheet.Cells(smryCount + 1, 11).Value = currentSheet.Cells(smryCount + 1, 17).Value - currentSheet.Cells(smryCount + 1, 15).Value
            ' Percentage change  = Yearly change / Lowest Date value (col 15)
            ' Avoid divide by zero
            If currentSheet.Cells(smryCount + 1, 15).Value = 0 Then
                currentSheet.Cells(smryCount + 1, 12).Value = 0
            Else
                currentSheet.Cells(smryCount + 1, 12).Value = currentSheet.Cells(smryCount + 1, 11).Value / currentSheet.Cells(smryCount + 1, 15).Value
            End If

            ' save this as the counter for previous symbol of summary
            prevSmrycounter = smryCount + 1
            
        End If
    
    ' Next Row
    Next dataCounter

    

    ' Populate the 2nd summary table after the 1st summary table processing is complete
    ' Since this procedure is coded to work on an unsorted list - this has to be done outside of the 1st summary population logic
    For smryCounter = 2 To (smryCount + 2) Step 1
        If (currentSheet.Cells(smryCounter, 12).Value > currentSheet.Cells(2, 22).Value) Or IsEmpty(currentSheet.Cells(2, 22)) Then
            currentSheet.Cells(2, 21).Value = currentSheet.Cells(smryCounter, 10).Value
            currentSheet.Cells(2, 22).Value = currentSheet.Cells(smryCounter, 12).Value
        End If
        ' Highest % Loser
        If currentSheet.Cells(smryCounter, 12).Value < currentSheet.Cells(3, 22).Value Or IsEmpty(currentSheet.Cells(3, 22)) Then
            currentSheet.Cells(3, 21).Value = currentSheet.Cells(smryCounter, 10).Value
            currentSheet.Cells(3, 22).Value = currentSheet.Cells(smryCounter, 12).Value
        End If
        ' Highest Volume
        If currentSheet.Cells(smryCounter, 13).Value > currentSheet.Cells(4, 22).Value Or IsEmpty(currentSheet.Cells(4, 22)) Then
            currentSheet.Cells(4, 21).Value = currentSheet.Cells(smryCounter, 10).Value
            currentSheet.Cells(4, 22).Value = currentSheet.Cells(smryCounter, 13).Value
        End If
    Next smryCounter
    
    
    ' Delete the 4 temporary columns created to store high & low values with dates - Cols 14 to 17
    For loopCounter = 17 To 14 Step -1
        currentSheet.Columns(loopCounter).Delete
    Next loopCounter
    
    
    ' Format the cells of Both summary tables appropriately
    ' 1st Summary:
    '   Format Column J as General
    '   Format Column K as Percent
    '   Format Column L as Percent
    '   Format Column M as General
    currentSheet.Columns("J").NumberFormat = "General"
    currentSheet.Columns("J").EntireColumn.AutoFit
    currentSheet.Columns("K").NumberFormat = "General"
    currentSheet.Columns("K").EntireColumn.AutoFit
    currentSheet.Columns("L").NumberFormat = "0.00%"
    currentSheet.Columns("L").EntireColumn.AutoFit
    currentSheet.Columns("M").NumberFormat = "General"
    currentSheet.Columns("M").EntireColumn.AutoFit
    
    ' 2nd Summary:
    '   Format Column P as General
    '   Format Column Q as General
    '   Format Column R as Row 2 & 3 as Percent
    '   Format Column R as Row 4 as General
    currentSheet.Columns("P").NumberFormat = "General"
    currentSheet.Columns("P").EntireColumn.AutoFit
    currentSheet.Columns("Q").NumberFormat = "General"
    currentSheet.Columns("Q").EntireColumn.AutoFit
    currentSheet.Cells(1, 18).NumberFormat = "General"
    currentSheet.Cells(2, 18).NumberFormat = "0.00%"
    currentSheet.Cells(3, 18).NumberFormat = "0.00%"
    currentSheet.Cells(4, 18).NumberFormat = "General"
    currentSheet.Columns("R").EntireColumn.AutoFit

'Next Sheet
Next


End Sub




