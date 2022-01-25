import tkinter as tk
from tkinter import PhotoImage, ttk, messagebox
from tkinter.filedialog import askopenfilenames
import fitz
from datetime import datetime
import win32api
import win32print

################################################################################

with open('./users.txt', 'r') as users:
    usuarios = users.readlines()
    lista_usuarios = [usuario.strip() for usuario in usuarios]

printers = win32print.EnumPrinters(2)
list_printers = []

for i, printer in enumerate(printers):
    n, driver, printer_name, description = printer
    imp = f'{i} - {printer_name}'
    list_printers.append(imp)

files = ''



################################################################################


def get_file_dimension(arquivo):
    doc = fitz.open(arquivo)
    largura_arquivo = 0
    area_arquivo = 0
    total_paginas = doc.page_count
    for page in doc:
        taxa = page.mediabox.height/594
        altura = round(page.mediabox.height/taxa)
        largura = round(page.mediabox.width/taxa)
        largura_arquivo += largura/1000
        area = largura/1000 * altura/1000
        area_arquivo += area

    return arquivo, total_paginas, round(largura_arquivo,2), round(area_arquivo,4)

def load_files():
    global files

    files = askopenfilenames(title="Selecione arquivos para impressão")

    for file in files:
        info_files = ''
        arquivo, total_page, lenght_file, area_file = get_file_dimension(file)
        lista_box_arquivos.insert('end' , arquivo)
        info_files = f'{total_page} / {lenght_file} / {area_file}'
        lista_box_arquivos_info.insert('end' , info_files)

def print_files():
    global files
    selected_user = combo_user.get()
    selected_printer = combo_printer.get()

    if selected_printer and selected_user and files:
        selected_printer = int(selected_printer[0:2].strip())
        win32print.SetDefaultPrinter(printers[selected_printer][2])

        log = f'./log/log de impressão - {datetime.now().strftime("%B")}.txt'
        if messagebox.askokcancel(title="Confirmar impressão", message="Deseja confirmar a impressão de arquivos?"):
            with open(log, 'a') as log:
                    for file in files:
                        win32api.ShellExecute(0, "print", file, None, ".", 0)
                        info_files = ''
                        arquivo, total_page, lenght_file, area_file = get_file_dimension(file)
                        info_files = f'\n{datetime.now().strftime("%d/%m/%Y - %H:%M:%S")};\t{selected_user};\t{arquivo};\t{total_page};\t{lenght_file};\t{area_file};\t{printers[selected_printer][2]};'
                        log.write(info_files)
    else:
        messagebox.showerror(title="Falta informação", message="Você deve selecionar impressora, projetista e os arquivos")

def clear_all():
    global files
    combo_printer.set('')
    combo_user.set('')
    lista_box_arquivos.delete(0, 'end')
    lista_box_arquivos_info.delete(0, 'end')
    files = ''

################################################################################

root = tk.Tk()
root.geometry('1280x600')
root.title("Controle de impressão ENGETEC")
root.columnconfigure(0, weight=1)
root.iconphoto(False, PhotoImage(file='./config/icone.png'))

# root.iconbitmap('icone.ico')

msg = tk.Label(text="Selecione a impressora")
msg.grid(row=0, column=0, padx=15, pady=5, sticky='nsew')
msg = tk.Label(text="Selecione o projetista")
msg.grid(row=0, column=1, padx=15, pady=5, sticky='nsew')
combo_printer = ttk.Combobox(root, values=list_printers)
combo_printer.grid(row=1, column=0, padx=15, sticky='nsew')
combo_user = ttk.Combobox(root, values=lista_usuarios)
combo_user.grid(row=1, column=1, padx=15, sticky='nsew')

bt_load = tk.Button(text="Carregar arquivos", command=load_files)
bt_load.grid(row=2, column=0, padx=15, pady=10, sticky='nsew')
bt_clear = tk.Button(text="Limpar tudo", command=clear_all)
bt_clear.grid(row=2, column=1, padx=15, pady=10, sticky='nsew')
bt_print = tk.Button(text="Imprimir arquivos listados", command=print_files, bg="#EAC7CE", font=30)
bt_print.grid(row=7, column=0, padx=15, pady=15, columnspan=2, ipady=20, sticky='ew')

msg = tk.Label(text="Arquivos selecionados")
msg.grid(row=4, column=0, columnspan=2, sticky='nsew', pady=10)
msg = tk.Label(text="Arquivo")
msg.grid(row=5, column=0, padx=50, sticky='W')
msg = tk.Label(text="Paginas[unid] / Comprimento[m] / Área[m²]")
msg.grid(row=5, column=1, padx=50, sticky='E')

frame = tk.Frame(root)
frame.grid(row=6, column=0, columnspan=2, sticky='nsew')

lista_box_arquivos = tk.Listbox(frame)
lista_box_arquivos.grid(row=0, column=0, sticky='nsew', padx=15, ipadx=444, ipady=130)
lista_box_arquivos_info = tk.Listbox(frame)
lista_box_arquivos_info.grid(row=0, column=1, sticky='nsew', padx=15, ipadx=145, ipady=130)

scrollbar = ttk.Scrollbar(lista_box_arquivos, orient='vertical', command=lista_box_arquivos.yview)
scrollbar.pack(side='right', fill='both')
lista_box_arquivos['yscrollcommand'] = scrollbar.set
scrollbar = ttk.Scrollbar(lista_box_arquivos_info, orient='vertical', command=lista_box_arquivos_info.yview)
scrollbar.pack(side='right', fill='both')
lista_box_arquivos_info['yscrollcommand'] = scrollbar.set

root.mainloop()