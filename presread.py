'''Reading print out from PRESEL'''

import xlsxwriter



''' Function to read lines into a list '''
def toList(nameoffile):
    with open(nameoffile, 'r') as f:
        #llist = [line.strip() for line in f] 
    	llist = [line for line in f]
    return llist

''' Function to clean the list '''
def cleanList(input_list):
	cleaned_list = []
	for item in input_list:
		if item[31:37].strip().isnumeric() :
			cleaned_list.append(item.rstrip('\n'))
	return cleaned_list

''' Function to avoid error in converting to integer '''
def to_integer(string):
	if string != '' :
		string = int(string)	

	return string

''' Function to avoid error in converting to float '''
def to_float(string):
	if string != '' :
		string = float(string)

	return string


''' function to write to excel file '''
def writeToExcel(inverted_loadcase, output_name) :
	''' Create a workbook and add a worksheet. '''
	workbook = xlsxwriter.Workbook(output_name +".xlsx")
	worksheet = workbook.add_worksheet()

	''' starting point '''
	row = 5
	col = 5

	list_top_case = get_top_sort(inverted_loadcase)
	indexed_topcase = column_indexing(list_top_case)

	''' Write header of top LC '''
	for item in indexed_topcase.keys() :
		worksheet.write(row, col + indexed_topcase[item] + 3, item)


	selno = 0
	lcno = 0
	for item in inverted_loadcase :
		selno_temp = item[0]
		lcno_temp = item[1]
		if selno == selno_temp and lcno == lcno_temp :
			pass
		else :
			row = row + 1

		pos_col = col + indexed_topcase[item[2]]
		selno = item[0]
		lcno = item[1]
		factor = item[3]
		pos_row = row

		worksheet.write(pos_row, col + 1, selno)  #Write superelement
		worksheet.write(pos_row, col + 2, lcno)  #Write low lc
		worksheet.write(pos_row, pos_col + 3, factor) #Write factor on corresponding top LC


	print("creating file " + output_name + ".xlsx" )

	workbook.close()
	
	return None

''' Function to get and sort top load case number'''
def get_top_sort(inverted_loadcase) :
	temp_list_1 = []
	for item in inverted_loadcase :
		top_no = item[2]
		if top_no not in temp_list_1 :
			temp_list_1.append(top_no)

	temp_list_1.sort()

	return temp_list_1

''' Function to generate dictionary for column position'''
def column_indexing(list_top_load) :
	i = 1
	indexed_topcase = {}
	for item in list_top_load :
		indexed_topcase[item] = i
		i +=1

	return indexed_topcase





'''Open and read file, and read it line by line'''
list_raw = toList("LCOMBStorm_100.LIS") 

'''Preparing the clean list'''
list_cleaned = cleanList(list_raw)

#print(*list_cleaned, sep='\n')

#Building data structure of load combination
'''

[
	[TOPLC, 
		[SEL, INDEX, 
			[ [LC, FACTOR], [LC, FACTOR], [LC, FACTOR], [], ..]
		]
		[SEL, INDEX,
			[ [LC, FACTOR], [LC, FACTOR], .. ]
		]

	]
]			
'''

load_combination = []

#print(list_cleaned[0][9:15])
#print(len(list_cleaned[0][9:15]))

#print(int(list_cleaned[0][9:15].strip()) + 1)



for line in list_cleaned :
	top_loadcomb_no = to_integer(line[:7].strip())
	low_sel_no = to_integer(line[9:15].strip())
	index_no = to_integer(line[16:22].strip())
	
	load_case_1 = to_integer(line[31:37].strip())
	factor_1	= to_float(line[38:46].strip())
	load_case_2 = to_integer(line[48:54].strip())
	factor_2	= to_float(line[55:63].strip())
	load_case_3	= to_integer(line[65:71].strip())
	factor_3	= to_float(line[72:80].strip())

	if top_loadcomb_no != '' :
			if load_case_2 != '' :
				if load_case_3 != '' :
					load_combination.append([top_loadcomb_no,[
						low_sel_no, index_no,[
						[load_case_1, factor_1], 
						[load_case_2, factor_2], 
						[load_case_3, factor_3]
						]]]
						)

				else :
					load_combination.append([top_loadcomb_no,[
						low_sel_no, index_no,[
						[load_case_1, factor_1], 
						[load_case_2, factor_2]
						]]]
						)

			else :
				load_combination.append([top_loadcomb_no,[
					low_sel_no, index_no,[
					[load_case_1, factor_1]
					]]]
					)


	elif low_sel_no != '' :
		target_load = len(load_combination) - 1
		if load_case_2 != '' :
			if load_case_3 != '' :
				load_combination[target_load].append([low_sel_no, index_no,[
					[load_case_1, factor_1],
					[load_case_2, factor_2],
					[load_case_3, factor_3]
					]]
					)

			else :
				load_combination[target_load].append([low_sel_no, index_no,[
					[load_case_1, factor_1],
					[load_case_2, factor_2]
					]]
					)

		else :
			load_combination[target_load].append([low_sel_no, index_no,[
				[load_case_1, factor_1]
				]]
				)			


	else :
		target_load = len(load_combination) - 1
		target_sel = len(load_combination[target_load]) - 1
		if load_case_2 != '' :
			if load_case_3 != '' :
				load_combination[target_load][target_sel][2].append(
					[load_case_1, factor_1]
					)
				load_combination[target_load][target_sel][2].append(
					[load_case_2, factor_2]
					)
				load_combination[target_load][target_sel][2].append(
					[load_case_3, factor_3]
					)

			else :
				load_combination[target_load][target_sel][2].append(
					[load_case_1, factor_1]
					)
				load_combination[target_load][target_sel][2].append(
					[load_case_2, factor_2]
					)

		else :
			load_combination[target_load][target_sel][2].append(
				[load_case_1, factor_1]
				)


#print(*load_combination[:7], sep='\n \n')


''' inverse the data, per sel and low lc'''

max_LC_no = 101
give_sel_no = [1,2,3,4,5,6,10,50,100]
load_combination_inv = []

for every_sel_no in give_sel_no :							#loop given sel no
	for every_lc_no in range(max_LC_no) :					#loop possible low lc no
		for every_top_lc in load_combination :				#loop every top lc in the database
			for sel_item in every_top_lc :
				if isinstance(sel_item, (list,)) :
					if sel_item[0] == every_sel_no :
						for lc_item in sel_item[2] :
							if lc_item[0] == every_lc_no :
								#print (str(every_sel_no) + " " + str(every_lc_no) + " " + "top load : " + str(every_top_lc[0]) 
								#	+ " factor : " + str(lc_item[1]))  
								load_combination_inv.append([every_sel_no, every_lc_no, every_top_lc[0], lc_item[1]  ])



#print(*load_combination_inv, sep='\n')

#top_load_sorted = get_top_sort(load_combination_inv)
#print(top_load_sorted)

#top_load_indexed = column_indexing(top_load_sorted)
#print(top_load_indexed)

''' write to excel file '''
writeToExcel(load_combination_inv, "LoadCombination_SEL100")










