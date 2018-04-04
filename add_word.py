
import xlsxwriter
from xlrd import open_workbook
import itertools

workbook = xlsxwriter.Workbook('words.xlsx')
worksheet = workbook.add_worksheet('Add_Word')

wb1 = open_workbook('words.xlsx')
ws1 = wb1.sheet_by_name('Add_Word')
row = 0
col = 0
print('Add English Words To Dictionary!')
word = raw_input('Please enter a word: ')
Dword = raw_input('Meaning of Dari: ')
Dword = Dword.decode(encoding='UTF-8',errors='strict')
Pword = raw_input('Meaning of Pashto: ')
Pword = Pword.decode(encoding='UTF-8',errors='strict')


sheet = wb1.sheet_by_index(0)

vals = sheet.col_values(0)
Dvals = sheet.col_values(1)
Pvals = sheet.col_values(2)
for w in vals:
	if w == word:
		print('This word is already exist in dictionary')
		break
else:
	vals.append(word)
	Dvals.append(Dword)
	Pvals.append(Pword)
	print('Word added to dictionary')

for eng, dari, pashto in itertools.izip(vals, Dvals, Pvals):
     worksheet.write(row, col,eng)
     worksheet.write(row, col+1,dari)
     worksheet.write(row, col+2,pashto)
     row += 1


workbook.close()
 