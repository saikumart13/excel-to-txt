def main():
	import os
	import openpyxl
	loc = input('Enter excel file name(with location): ')
	des = input('Enter destination to store text files:')
	row = int(input('Enter no. of rows to be copied(>0)'))
	col = int(input('Enter no. of columns to be copied(>0)'))
	assert (col > 0) & (row > 0), 'Row and Column must be greater than 0'
	assert (os.path.isabs(loc)) & (os.path.isabs(des)), 'Path not available'
	try:
		wb = openpyxl.load_workbook(loc)
		sheet = wb.active
		dest = os.path.join(des,'output.txt')
		with open(dest,'w') as out_file:
			for i in range(1,row+1):
				for j in range(1,col+1):
					out_file.write(sheet.cell(row = j, column =i).value+' ')
				out_file.write('\n')
		print('File saved as output.txt at location '+des)#os.getcwd())
	except:
		print('Error occured')
	finally:
		input()


if __name__ == '__main__':
	main()