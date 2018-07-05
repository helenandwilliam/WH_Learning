#code - UTF-8

## coding:utf-8

import openpyxl

Pre_data = openpyxl.load_workbook('5.14美团于洪区.xlsx')
Next_data = openpyxl.load_workbook('5.23美团于洪区.xlsx')


work = openpyxl.Workbook()
table = work.active
table.title = 'standard_list'

Pre_List_name = Pre_data.sheetnames #
Pre_sheet = Pre_data.get_sheet_by_name(Pre_List_name[0])#获取表格
Next_List_name = Next_data.sheetnames #
Next_sheet = Next_data.get_sheet_by_name(Next_List_name[0])#获取表格

Pre_Row_Num = Pre_sheet.max_row
Pre_Col_Num = Pre_sheet.max_column
Next_Row_Num = Next_sheet.max_row
Next_Col_Num = Next_sheet.max_column

List3_Row_Num = 0
Information_Num = 0

table.cell(row= 1, column= 1).value = '标准表店铺（5.14）'
table.cell(row= 1, column= 2).value= '团购信息名'
table.cell(row= 1, column= 3).value= '销售量'
table.cell(row= 1, column= 4).value = '售价'
table.cell(row= 1, column= 5).value = '店铺名（5.23）'
table.cell(row= 1, column= 6).value= '团购信息名'
table.cell(row= 1, column= 7).value= '销售量'
table.cell(row= 1, column= 8).value = '售价'

Same_column = 0

for i in range (2 ,Pre_Row_Num + 1 )
	for j in range (2 , Next_Row_Num + 1) 
		if Pre_sheet.cell(row=i, column=1).value == Next_sheet.cell(row=j, column=1).value or \
				Pre_sheet.cell(row=i, column=2).value == Next_sheet.cell(row=j, column=2).value
			Same_column=j
			for (k in range (0,51))
				if Pre_sheet.cell(row=i, column=3).value == Next_sheet.cell(row=j+k, column=3).value
					table.cell(row= i, column= 1).value = Pre_sheet.cell(row=i, column=1).value
					table.cell(row= i, column= 2).value = Pre_sheet.cell(row=i, column=3).value
					table.cell(row= i, column= 3).value = Pre_sheet.cell(row=i, column=4).value
					table.cell(row= i, column= 4).value = Pre_sheet.cell(row=i, column=5).value
					table.cell(row= i, column= 5).value = Next_sheet.cell(row=j+k, column=4).value
					table.cell(row= i, column= 6).value = Next_sheet.cell(row=j+k, column=5).value
			break		
					
					
				
				
				
				
				
				