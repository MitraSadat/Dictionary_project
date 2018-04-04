import os
from xlrd import open_workbook

wb = open_workbook('words.xlsx')
ws = wb.sheet_by_index(0)

def one():
    eng = ws.col_values(0, 1)
    dari = ws.col_values(1, 1)
    dic = {a : b for a, b in zip(eng, dari)}

    print('Enter a english word to translate!')
    search_word = raw_input()
    for word in dic.keys():
    	if word == search_word.lower():
    		dar = dic[word]
    		print(dar)
    		break
    else:
    	print('This word is not in dictionary.')
    
def two():
    eng = ws.col_values(0, 1)
    pashto = ws.col_values(2, 1)
    dic2 = {a : b for a, b in zip(eng, pashto)}

    print('Enter a english word to translate!')
    search_word = raw_input()
    for word in dic2.keys():
    	if word == search_word.lower():
    		pash = dic2[word]
    		print(pash)
    		break
    else:
    	print('This word is not in dictionary.')

 
def numbers_to_months(argument):
    switcher = {
        1: one,
        2: two,
    }
    func = switcher.get(argument, lambda: "Invalid Option")
    # Execute the function
    func()

string = input('1: press 1 for translating English to Dari!\n2: press 2 for translating English to Pashto!\n ')
numbers_to_months(string)