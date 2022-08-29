from ftplib import B_CRLF
from fillnprint import FillNPrint
import threading
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog as fd
import os

root = tk.Tk()
root.title('FillNPrint')
root.geometry('640x480+50+50')
root.minsize(640, 480)

exc_var = tk.StringVar(root, None)
cfg_var = tk.StringVar(root, None)
out_var = tk.StringVar(root, None)
sht_var = tk.StringVar(root, None)
cel_var = tk.StringVar(root, None)
lmt_var = tk.StringVar(root, None)

sheets=['']


#open window for selecting excel file
def select_excel_file():
    filetypes = (
        ('Excel files', '*.xlsx'),
        ('All files', '*.*')
    )
    filename = fd.askopenfilename(
        title='Open a file',
        initialdir=os.path.expanduser('~'),
        filetypes=filetypes)
    exc_var = tk.StringVar(root, filename)
    ef_entry.delete(0,tk.END)
    ef_entry.insert(0,filename)
    excel_file('')


#open window for selecting config file
def select_yaml_file():
    filetypes = (('All files', '*.*'),)
    filename = fd.askopenfilename(
        title='Open a file',
        initialdir=os.path.expanduser('~'),
        filetypes=filetypes)
    cfg_var = tk.StringVar(root, filename)
    cg_entry.delete(0,tk.END)
    cg_entry.insert(0,filename)


#open window for selecting output folder
def select_output():
    filetypes = (
        ('PDF files', '*.pdf'),
        ('All files', '*.*')
    )
    filename = fd.asksaveasfilename(
        title='Save as',
        initialdir=os.path.expanduser('~'),
        filetypes=filetypes)
    out_var = tk.StringVar(root, filename)
    op_entry.delete(0,tk.END)
    op_entry.insert(0,filename)


#excel file selected
def excel_file(a):
    sheets = FillNPrint(None, exc_var.get()).get_sheets()
    bs_combobox['values'] = sheets
    bs_combobox.set(sheets[0])

#label
title = ttk.Label(root, text="FillNPrint")
title.config(font=("TkDefaultFont", 32))
title.pack(expand=True, fill='x', padx=10, pady=10)


#excel file section
ef = ttk.LabelFrame(root, text='Excel File') #excel file frame
ef.pack(expand=True, fill='x', padx=10, pady=10)
ef.grid_columnconfigure(0, weight=1)

ef_entry = ttk.Entry(ef, textvariable=exc_var,takefocus=False) #excel file entry input
ef_entry.grid(column=0, row=0, padx=10, pady=10, sticky='ew')
ef_entry.bind("<FocusOut>", excel_file)
ef_entry.bind('<Control-a>', lambda x: ef_entry.selection_range(0, 'end') or "break")

ef_browse = ttk.Button(ef, text='Browse', command=select_excel_file, takefocus=False) #excel file browse button
ef_browse.grid(column=1, row=0, padx=10, pady=10)


#options section
st = ttk.Frame(root)
st.pack(expand=True, fill='x', padx=10)


#sheet
bs = ttk.Frame(st)#box size frame
bs.grid(column=0, row=1,padx=10, pady=5, sticky='w')
bs.grid_columnconfigure(0, weight=1)

bs_combobox = ttk.Combobox(bs, textvariable=sht_var, width=8) #box size spinbox
bs_combobox['values'] = sheets
bs_combobox.set(sheets[0])
bs_combobox.bind("<FocusOut>", lambda event: config_instance.save('box_size', sht_var.get(), True))
bs_combobox.grid(column=1, row=0)

bs_text = ttk.Label(bs, text='Sheet') #box size label
bs_text.grid(column=0, row=0)


#starting cell
sc = ttk.Frame(st)#starting cell frame
sc.grid(column=1, row=1,padx=10, pady=5, sticky='w')
sc.grid_columnconfigure(0, weight=1)

sc_entry = ttk.Entry(sc, textvariable=cel_var, width=5, takefocus=False) #starting cell spinbox
sc_entry.bind("<FocusOut>", lambda event: config_instance.save('starting_cell', cel_var.get(), True))
sc_entry.grid(column=1, row=0)

sc_text = ttk.Label(sc, text='Starting Cell') #starting cell label
sc_text.grid(column=0, row=0)


#limit
lm = ttk.Frame(st)#box size frame
lm.grid(column=2, row=1,padx=10, pady=5, sticky='w')
lm.grid_columnconfigure(0, weight=1)

lm_spinbox = ttk.Spinbox(lm, textvariable=lmt_var, from_=0, to=1000, width=3, takefocus=False) #box size spinbox
lm_spinbox.bind("<FocusOut>", lambda event: config_instance.save('box_size', lmt_var.get(), True))
lm_spinbox.grid(column=1, row=0)

lm_text = ttk.Label(lm, text='Limit') #box size label
lm_text.grid(column=0, row=0)


#config file section
cg = ttk.LabelFrame(root, text='Configuration File') #excel file frame
cg.pack(expand=True, fill='x', padx=10, pady=10)
cg.grid_columnconfigure(0, weight=1)

cg_entry = ttk.Entry(cg, textvariable=cfg_var, takefocus=False) #excel file entry input
cg_entry.grid(column=0, row=0, padx=10, pady=10, sticky='ew')
cg_entry.bind("<FocusOut>", lambda event: config_instance.save('excel_file', cg_var.get(), True))
cg_entry.bind('<Control-a>', lambda x: cg_entry.selection_range(0, 'end') or "break")

cg_browse = ttk.Button(cg, text='Browse', command=select_yaml_file, takefocus=False) #excel file browse button
cg_browse.grid(column=1, row=0, padx=10, pady=10)


#excel file section
op = ttk.LabelFrame(root, text='Output File') #excel file frame
op.pack(expand=True, fill='x', padx=10, pady=10)
op.grid_columnconfigure(0, weight=1)

op_entry = ttk.Entry(op, textvariable=out_var, takefocus=False) #excel file entry input
op_entry.grid(column=0, row=0, padx=10, pady=10, sticky='ew')
op_entry.bind("<FocusOut>", lambda event: config_instance.save('excel_file', op_var.get(), True))
op_entry.bind('<Control-a>', lambda x: op_entry.selection_range(0, 'end') or "break")

op_browse = ttk.Button(op, text='Browse', command=select_output, takefocus=False) #excel file browse button
op_browse.grid(column=1, row=0, padx=10, pady=10)


#generate sectopn
gs = ttk.Frame(root) #generate section frame
gs.pack(expand=True, fill='x',side='bottom', anchor='s')

#generate button
gn = ttk.Button(gs,
                text='Generate',
                #command=generate,
                takefocus=False)
gn.pack(side='right', padx=20, pady=20)

#progress section
ps = ttk.Frame(gs) #progress section text
ps.pack(expand=True, side='right', fill='x')
ps.columnconfigure(0, weight=1)

pt = ttk.Label(ps, text='') #progress text
pt.grid(column=0, row=0, padx=30, sticky='w')

pb = ttk.Progressbar( #progress bar
    ps,
    orient='horizontal',
    mode='determinate',
    length=480,
)


root.mainloop()