import tkinter as tk
from openpyxl import *
from tkinter.messagebox import showinfo
from tkinter import filedialog,messagebox

win = tk.Tk()
win.title("Registro") #Nombre de la ventana de registro

#Ruta del archivo usado por openpyxl
path = r'C:\Users\Familia Espinoza\Desktop\Python'
dest_filename = r'C:\Users\user\Desktop\Python\Libro1.xlsx'

#Del modulo 'openpyxl' activa el libro de excel
wb = Workbook()

'''
Funcion definida para guardar los registros, sin validadores de campo, sin verificar
que existan otros registros en el archivo destino.
'''
def save():

    f_name = entry_f.get()
    e_name = entry_e.get()
    num = int (entry_n.get())

    wb = Workbook()
    ws = wb.active      #Activa la hoja

    ws.title = "Equipo" #titulo de hoja
    ws['A1'] = 'Nombre'
    ws['B1'] = 'Equipo'
    ws['C1'] = 'Numero'
    ws['A2'] = f_name
    ws['B2'] = e_name
    ws['C2'] = num

    wb.save(filename = dest_filename)
    showinfo("Guardado", "Registros guardados con exito")
    clear()

def clear():
    entry_f.delete(0, tk.END)
    entry_e.delete(0, tk.END)
    entry_n.delete(0, tk.END)

def open_fold():
    path = filedialog.askdirectory()
    if path == "":
        path = dest_filename
    #print(path)   #Con esta instrucci[on puede ver el cambio de directorio en el terminal o CMD

label = tk.Label(win, text="Conexion de Python con Excel", bg= "yellow")
label.grid(row=0, column=0)

f_label = tk.Label(win, text="Nombre: ")
f_label.grid(row=1, column=0, padx=8, pady=8)

e_label = tk.Label(win, text="Equipo: ")
e_label.grid(row=2, column=0, padx=8, pady=8)

n_label = tk.Label(win, text="Número: ")
n_label.grid(row=3, column=0, padx=8, pady=8)

entry_f = tk.Entry(win)
entry_f.grid(row=1, column=1, padx=8, pady=8)

entry_e = tk.Entry(win)
entry_e.grid(row=2, column=1, padx=8, pady=8)

entry_n = tk.Entry(win)
entry_n.grid(row=3, column=1, padx=8, pady=8)

button = tk.Button(win, text="Registrar", command=save)
button.grid(row=5, column=0, columnspan=2, padx=8, pady=8)

button2 = tk.Button(win, text="Limpiar", command=clear)
button2.grid(row=5, column=1, columnspan=2, padx=8, pady=8)

button3 = tk.Button(win,text = "carpeta", command=open_fold)
button3.grid(row=7, column=0, columnspan= 3, padx=8, pady=8)

wb = load_workbook(filename = 'empty_book.xlsx')
sheet_ranges = wb['range names']
print(sheet_ranges['D18'].value)

win.mainloop()

