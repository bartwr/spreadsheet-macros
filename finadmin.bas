REM  *****  BASIC  *****

Sub Main

End Sub

Sub HelloMacro

  Print "Hello"

End Sub

REM https://www.debugpoint.com/2014/10/range-processing-using-macro-in-libreoffice-calc-part-1/
Sub rang_processing_demo
	dim my_doc as object
	dim my_sheets as object
	dim my_range as object
	my_doc = ThisComponent
	REM Get sheets
	my_sheets = my_doc.Sheets

	REM Set cell data
	dim my_data(0,0)
	my_data(0,0) = "Apple2"
	
	REM Fill column with data
	REM my_range = my_sheets(0).getCellRangebyName("B2")
	REM my_range.setDataArray(my_data)

	' https://www.debugpoint.com/2014/09/libreoffice-workbook-worksheet-and-cell-processing-using-macro/
	dim my_cell as object
	my_cell = ThisComponent.Sheets(0).getCellByPosition(1,2)
	print(my_cell.VALUE)

End Sub

Sub processing_sheets_cells

	dim my_doc   as object
	Dim my_sheets as object
	Dim my_cell as object

	Dim sheet_count, i, row, col, cell_value, str
	
	my_doc = ThisComponent
	my_sheets = my_doc.Sheets 
	sheet_count = my_sheets.Count

	for i = 0 to sheet_count - 1		
		str = str & chr(13) & "--------" & chr(13)  
		for row=1 to 4
				for col=0 to 1
					my_cell = ThisComponent.Sheets(i).getCellByPosition(col,row)
					Select Case my_cell.Type
						Case com.sun.star.table.CellContentType.VALUE
							cell_value = my_cell.Value
						Case com.sun.star.table.CellContentType.TEXT
							cell_value = my_cell.String
					End Select
					str = str & " " & cell_value
				next col
				str = str & Chr(13)
		next row
	next i
	msgbox str
	
End Sub

Sub testB
	dim my_doc   as object
	Dim my_sheets as object
	Dim my_cell as object

	Dim sheet_count, i, col, cell_value, str
	dim ledgerSheet, transactionSheet
	dim currentYear, forMonth
	dim ledgerRow, transactionRow, categoryCell, sumValue
	dim transactionDateCell, transactionCategoryCell, transactionAmountCell, cellYear, cellMonth, cellToUpdate
	dim my_data(0,0)

	my_doc = ThisComponent
	my_sheets = my_doc.Sheets 

	currentYear = "2020"
	ledgerSheet = my_sheets(1)
	transactionSheet = my_sheets(2)

	'GET LEDGER CATEGORIES
'	for ledgerRow = 53 to 53
	for ledgerRow = 12 to 152
'	for ledgerRow = 131 to 131
      categoryCell = ledgerSheet.getCellByPosition(0,ledgerRow)
	  'LOOP MONTHS
	  for forMonth = 1 to 12
	  	'FOR EVERY MONTH: GET SUM OF TRANSACTIONS FOR THIS CATEGORY
	  	sumValue = 0
		for transactionRow = 0 to 3000
		    transactionCategoryCell = transactionSheet.getCellByPosition(0,transactionRow)
		    transactionDateCell = transactionSheet.getCellByPosition(5,transactionRow)
		    transactionAmountCell = transactionSheet.getCellByPosition(7,transactionRow)
	        cellYear = Format(transactionDateCell.String, "yyyy")
	        cellMonth = Format(transactionDateCell.String, "m")
			if cellYear = currentYear and CInt(cellMonth) = forMonth and categoryCell.VALUE = transactionCategoryCell.VALUE then
				sumValue = CDbl(sumValue) + CDbl(transactionAmountCell.VALUE)
			'	print(CDbl(transactionAmountCell.VALUE))
			'	print("sumValue: " & sumValue)
			endif
		next transactionRow
		'FOR EVERY CATEGORY+MONTH COMBINATION: STORE SUM IN SPREADSHEET
		if sumValue <> 0 then
			cellToUpdate = ledgerSheet.getCellByPosition(9+forMonth,ledgerRow)
			my_data(0,0) = sumValue
			'Do not update header rows
			if categoryCell.VALUE Mod 100 <> 0 or categoryCell.VALUE < 1000 then
				cellToUpdate.setDataArray(my_data)
			endif
		endif
	  next forMonth
    next ledgerRow

End Sub

Sub updateSumForMonthAndCategory

	dim my_doc   as object
	Dim my_sheets as object
	Dim my_cell as object

	Dim sheet_count, i, row, col, cell_value, str
	
	my_doc = ThisComponent
	my_sheets = my_doc.Sheets 
	sheet_count = my_sheets.Count

	' Get months
		' Get categories
			' Sum amount for category & update

	for i = -1 to -1
		str = str & chr(13) & "--------" & chr(13)  
		for row=1 to 4
				for col=0 to 1
					my_cell = ThisComponent.Sheets(i).getCellByPosition(col,row)
					Select Case my_cell.Type
						Case com.sun.star.table.CellContentType.VALUE
							cell_value = my_cell.Value
						Case com.sun.star.table.CellContentType.TEXT
							cell_value = my_cell.String
					End Select
					str = str & " " & cell_value
				next col
				str = str & Chr(13)
		next row
	next i
	msgbox str
	
End Sub
