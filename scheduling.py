from tkinter import *
from bisect import bisect_left
import os
from os import path
from tkinter import filedialog
import openpyxl
from tkcalendar import Calendar,DateEntry


root = Tk()
root.title("Auto Scheduling Ujian")
root.geometry("300x300")

def browse_file():
	global browse_btn
	root.filename = filedialog.askopenfilename(initialdir=os.getcwd(),title="Select A File",filetypes=[("Comma Separated Values","*.csv"),("all files","*.*")])	
	browse_btn['text'] = path.split(root.filename)[1]

jumlah_sesi = 0
jumlah_hari = 0
weekend_exclude = False

def submit():
	global var
	global jumlah_sesi_input
	global start_date
	global end_date
	jumlah_sesi = jumlah_sesi_input.get()
	weekend_exclude = var.get()
	scheduling(root.filename,jumlah_sesi,weekend_exclude,start_date,end_date)


browse_label = Label(root,text="Input File (.csv) :",anchor=E,width=17)
browse_label.grid(column=0,row= 0)

jumlah_sesi_label = Label(root,text="Jumlah sesi per hari :",anchor=E,width=17)
jumlah_sesi_label.grid(row=2,column=0)

start_date_label = Label(root,text="Hari pertama ujian :",anchor=E,width =17)
start_date_label.grid(row=3,column=0)

end_date_label = Label(root,text="Hari terakhir ujian :",anchor=E,width =17)
end_date_label.grid(row=4,column=0)

browse_btn = Button(root,text="Browse File",padx=15,pady=0,anchor=W,width=13,command=browse_file)
browse_btn.grid(row=0,column=1,padx=10,pady=10)


jumlah_sesi_input = Entry(root)
jumlah_sesi_input.grid(row=2,column=1)

var = StringVar()


start_date = DateEntry(root,width=15,bg="darkblue",fg="white",year=2020)
start_date.grid(row=3,column=1)

end_date = DateEntry(root,width=15,bg='darkblue',fg='white',year=2020)
end_date.grid(row=4,column=1)

weekend = Checkbutton(root, text="Exclude Weekend",variable=var,onvalue="On",offvalue="Off")
weekend.grid(row=5,column=0)
weekend.deselect()


submit_btn = Button(root,text="Submit",padx=15,command=submit)
submit_btn.grid(row=6,column=0,columnspan=2)



# filename = filedialog.askopenfilename(initialdir=os.getcwd(),title="Select A File",filetypes=[("Comma Separated Values","*.csv"),("all files","*.*")])


def scheduling(filename,jumlah_sesi,weekend_exclude,start_date,end_date):
	from datetime import datetime,timedelta
	import numpy as np


	weekend_bool = False
	if weekend_exclude == 'On':
		weekend_bool = True
	jumlah_sesi = int(jumlah_sesi)

	error_text = Label(root,text='',fg='red')
	error_text.grid(row=7,column=0,columnspan=2)

	date = datetime.strptime(start_date.get(), '%m/%d/%y').date().weekday()
	end = datetime.strptime(end_date.get(), '%m/%d/%y').date().weekday()

	if ( date > 4  or end > 4 )and weekend_bool:
		error_text['text']="Tanggal dimulai dan tanggal berakhir tidak boleh weekend"
		return


	all_course = {}

	course2id = {}
	id2course = {}

	student2id = {}
	id2student = {}

	'''
	all_course = {
		course:[students],...
	}
	'''
	#convert student and course to id
	with open(filename) as f:
		data = f.readlines()

		for line in data:
			
			x = line.split(',')
			student,course = x[0],x[1]
			course = course.strip()
			student = student.strip()
			if course not in course2id.keys():
				course2id[course] = len(course2id.keys())
			if student not in student2id.keys():
				student2id[student] = len(student2id.keys()) 
			if course2id[course] not in all_course.keys():
				all_course[course2id[course]] = []
			all_course[course2id[course]].append(student2id[student])


	id2course = dict([(v,k) for k,v in course2id.items()])
	id2student = dict([(v,k) for k,v in student2id.items()])

	all_course_sorted = {}
	for key,value in all_course.items():
		all_course_sorted[key] = sorted(value)
		# print(sorted(value))

	all_color = []
	day_course = {}
	for x in range(len(all_course.keys())): 
		all_color.append(list(range(0,jumlah_sesi)))# 5 = total sesi

	def binary_search(a, x, lo=0, hi=None):  # can't use a to specify default for hi
	    hi = hi if hi is not None else len(a)  # hi defaults to len(a)   
	    pos = bisect_left(a, x, lo, hi)  # find insertion position
	    return pos if pos != hi and a[pos] == x else -1

	color_course = {}

	new_day = []
	count = 0
	print(course2id)
	while True :
		for key,value in all_course_sorted.items():
			if len(all_color[key]) == 0:
				continue
			color_course[key] = all_color[key][0]
			day_course[key] = count
			for key_2,value_2 in all_course_sorted.items():
				if key_2 == key:
					continue

				flag = True
				for x in value:
					if  binary_search(value_2, x) != -1 :
						flag = False
						# print("ketemu",key,key_2)
						break

				if flag is False:
					try:
						all_color[key_2].remove(color_course[key])
						if len(all_color[key_2]) == 0:
							new_day.append(key_2)

					except ValueError:
						pass
					# else:
					# 	print(key,key_2)

		temp = {}
		if len(new_day) > 0:
			for x in new_day:
				all_color[x] = list(range(0,jumlah_sesi))
				temp[x] = all_course_sorted[x]
		else:
			break
		all_course_sorted = temp
		new_day = []
		count += 1


	print(color_course,day_course,sep='\n')

	

	date = datetime.strptime(start_date.get(), '%m/%d/%y')
	end =  datetime.strptime(end_date.get(), '%m/%d/%y')
	
	if weekend_bool:
		print(np.busday_count(date.date(),end.date()) > len(set(day_course.values())))
		if np.busday_count(date.date(),end.date()) < len(set(day_course.values())) :
			error_text['text'] = 'Jumlah hari kurang, tambah lagi!'
			return
		else:
			error_text['text'] = 'Berhasil, cek jadwalnya di output.xlsx'
			error_text.config(fg='green')
	else:
		if abs((date.date() - end.date()).days) < len(set(day_course.values())):
			error_text['text'] = 'Jumlah hari kurang, tambah lagi!'
			return
		else:
			error_text['text'] = 'Berhasil, cek jadwalnya di output.xlsx'
			error_text.config(fg='green')	

	
	try:
		wb = openpyxl.Workbook()
		sheet = wb.active
		for index,x in enumerate(set(day_course.values())):

			date = datetime.strptime(start_date.get(), '%m/%d/%y') + timedelta(days=index)
			if weekend_bool and date.weekday() > 4:
				date = date + timedelta(days=2)
			sheet[chr(ord('B') + index) + '1'] = date.date()
		for i in  range(jumlah_sesi):
			sheet['A' + str(i+2)] = 'sesi_' + str(i)
		
		for key,value in color_course.items():
			hari = chr(ord('B') + day_course[key])
			sesi = str(value+2)
			if sheet[hari+sesi].value is not None :
				sheet[hari+sesi] = sheet[hari+sesi].value + ',' +  id2course[key]
			else:
				sheet[hari+sesi] = id2course[key]


		wb.save('output.xlsx')
	except PermissionError:
		error_text['text'] = 'File output.xlsx nya harus diclose'
		error_text.config(fg='red')
		return


root.mainloop()

