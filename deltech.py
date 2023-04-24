from openpyxl import load_workbook
from tkinter import ttk, filedialog, messagebox
from tkinter import *
import threading
from PIL import Image, ImageTk
import time
import webbrowser

#########################
all_rows = []
check = [""]
interfaces = ["fa1/x", "G1/0/x", "G2/0/x"]
#########################
#opyta sa pre dany excel file a spusti nasledovne thready
def add_file():
	file = filedialog.askopenfilename()
	threading.Thread(target=expand, args=(file, ) ).start()

#########################
#GUI
root = Tk()
root.title("Deltech v1.1")
root.geometry("700x500")
root.configure(bg="white")
root.resizable(False, False)

image_icon = PhotoImage(file="image/logo.png")
root.iconphoto(False, image_icon)

#LOGO
def open_website():
    webbrowser.open("https://www.deltech.sk")

logo = Image.open("image/delbar.png")
logores = logo.resize((200, 100))
logo_ = ImageTk.PhotoImage(logores)
Button(root, image=logo_, bg="white", bd=0, command=open_website).place(x=15, y=10)

cisco = Image.open("image/cisco.png")
ciscores = cisco.resize((100, 50))
cisco_ = ImageTk.PhotoImage(ciscores)
Label(root, image=cisco_, bg="white", bd=0).place(x=125,y=445)

#button pre pridanie excel file
add_button = PhotoImage(file="image/add_file.png")
Button(root, image=add_button, bg="white", bd=0, command=add_file).place(x=87,y=120)

loading = Label(root, text="", bg="white", bd=0)
loading.place(x=50,y=170)

#dropdown menu pre switche
variable_sw = StringVar(root)
variable_sw.set("SWITCH")
drop_menu_sw = OptionMenu(root, variable_sw, *check)
drop_menu_sw.config(width=15)
drop_menu_sw.place(x=50, y=190)

#dropdown menu pre intf
variable_int = StringVar(root)
variable_int.set("INTERFACE")
drop_menu_int = OptionMenu(root, variable_int, *interfaces)
drop_menu_int.config(width=15)
drop_menu_int.place(x=50, y=240)

###############################

#vymeni v intf za "x" dane porty z excelu
def replacer(s, newstring, index, nofail=False):
    if not nofail and index not in range(len(s)):
        raise ValueError("index vonka stringu")

    if index < 0:
    	return newstring + s
    if index > len(s):
    	return s + newstring
    return s[:index] + newstring + s[index + 1:]

#vypise info do gui, vypise do .txt
def output():
	text_box.delete("1.0", END)
	sw = variable_sw.get()
	print(sw)
	print("_" * 30)
	print("*** " + sw + " ***")
	text_box.insert(END, "!*** " + sw + ' ***\n')
	for k in range(len(all_rows) ):
		if sw == all_rows[k]['switch']:
			intf = variable_int.get()
			x_index = intf.index("x")
			intf = replacer(intf, str(all_rows[k]['port']), x_index)

			string = "!" + "*" * 30
			text_box.insert(END, string + "\n")
			#file.write(string + '\n')

			string1 = "interface " + intf
			text_box.insert(END, string1 + "\n")
			#file.write(string1 + '\n')

			string2 = "description " + all_rows[k]['name']
			text_box.insert(END, string2 + "\n")
			#file.write(string2  + '\n')

			string3 = "exit"
			text_box.insert(END, string3 + "\n")
			#file.write(string3  + '\n')
			
	print(check)

def extract_data():
	path = filedialog.askdirectory()
	sw = variable_sw.get()
	text = text_box.get("1.0", 'end')
	with open(path + "/" + sw + ".txt", "w") as file:
		file.write(text)
################################
#submit button co zavola funkciu output
submit_bt = Button(root, width=18, height=3, text="Submit", bg="#FC5956", bd=0, command=output)
submit_bt.place(x=50, y=290)

#save button co ulozi z textboxu text do filu .txt
save_bt = Button(root, width=18, height=3, text="Save", bg="#FC5956", bd=0, command=extract_data)
save_bt.place(x=50, y=350)

#scroll bar pre list
#list v gui kde sa vypisu uz upravene hodnoty
scroll = Scrollbar(root)
#playlist = Listbox(root, width=50 , font=("Arial", 12),bg="#333333", 
					#fg="grey", selectbackground="lightblue", cursor="hand2", 
					#bd=0, yscrollcommand=scroll.set)
text_box = Text(root, width=50, font=("Arial", 12),bg="black", fg="white", wrap="word")

def on_closing():
    root.destroy()
root.protocol("WM_DELETE_WINDOW", on_closing)

#inventors
demko = Label(root, text="S. Demko",)
demko.config(font=("Arial", 6), bg="white")
demko.place(x=10, y=467)

#pozicie list, scroll
scroll.config(command=text_box.yview)
scroll.pack(side=RIGHT, fill=Y)
text_box.pack(side=RIGHT, fill=Y)


###############################

#obnovy dropdown menu a priradi don hodnoty z excelu
def refresh(check):
	loading.config(text="Načítané!", fg="green")
	menu = drop_menu_sw.children["menu"]
	menu.delete(0, "end")
	for option in check:
		menu.add_command(label=option, command=lambda value=option: variable_sw.set(value))

###############################

#vypise z excelu .json format s ktorym sa dalej potom pracuje
def expand(file):
	variable_sw.set("SWITCH")
	variable_int.set("INTERFACE")
	check.clear()
	all_rows.clear()
	wb = load_workbook(file)
	sheet = wb.active

	rows = sheet.rows

	headers = [cell.value for cell in next(rows)]

	for row in rows:
		data = {}
		for title, cell in zip(headers, row):
			data[title] = cell.value

		all_rows.append(data)
	print(all_rows)
	for i in range(len(all_rows) ):
		checked = False
		for j in range(len(check) ):
			if check[j] == all_rows[i]['switch']:
				checked = True

		if not checked:
			check.insert(0, all_rows[i]['switch'])
	loading.config(text="Načítam...", fg="red")
	time.sleep(2)
	threading.Thread(target=refresh, args=(check, ) ).start()


if __name__ == "__main__":
	root.mainloop()


#D3mkO