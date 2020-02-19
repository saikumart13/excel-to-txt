def main():
	import os
	import openpyxl
	loc = input('Enter file name(with location): ')
	row = int(input('Enter no. of rows to be copied(>0)'))
	col = int(input('Enter no. of columns to be copied(>0)'))
	try:
		wb = openpyxl.load_workbook(loc)
		sheet = wb.active
		with open('output.txt','w') as out_file:
			for i in range(1,row+1):
				for j in range(1,col+1):
					out_file.write(sheet.cell(row = j, column =i).value+' ')
				out_file.write('\n')
		print('File saved as output.txt at location '+os.getcwd())
	except:
		print('some error occured')
	finally:
		input()


if __name__ == '__main__':
	main()