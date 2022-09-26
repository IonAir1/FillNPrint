from configparser import ConfigParser
from re import M
import threading
import tkinter as tk
import os
from tkinter import ttk
from tkinter import filedialog as fd
from fillnprint import FillNPrint


#read the save config
def read(file):
    cfg = ConfigParser()
    save = {
        'excel': '',
        'config': '',
        'output': '',
        'cell': '',
        'limit': '',
        'sheet': 0
    }
    cfg.read(file)
    if not cfg.has_section('main'):
        cfg.add_section('main')
    if cfg.has_option('main','excel'):
        save['excel'] = cfg.get('main', 'excel')
    if cfg.has_option('main','config'):
        save['config'] = cfg.get('main', 'config')
    if cfg.has_option('main','output'):
        save['output'] = cfg.get('main', 'output')
    if cfg.has_option('main','cell'):
        save['cell'] = cfg.get('main', 'cell')
    if cfg.has_option('main','limit'):
        save['limit'] = cfg.get('main', 'limit')
    if cfg.has_option('main', 'sheet'):
        save['sheet'] = cfg.get('main', 'sheet')
    return save


#save the save config
def save(file, key, val):
    if key == 'sheet':
        sh_combobox.selection_clear()
    cfg = ConfigParser()
    cfg.read(file)
    if not cfg.has_section('main'):
        cfg.add_section('main')
    cfg.set('main', str(key), str(val))
    with open(file, 'w') as f: #save
        cfg.write(f)


save_file = 'fillnprint.save'
saved = read(save_file)
root = tk.Tk()
root.title('FillNPrint')
root.geometry('640x480+50+50')
root.minsize(640, 480)


exl_var = tk.StringVar(root, saved['excel'])
cfg_var = tk.StringVar(root, saved['config'])
out_var = tk.StringVar(root, saved['output'])
sht_var = tk.StringVar(root, None)
cel_var = tk.StringVar(root, saved['cell'])
lmt_var = tk.StringVar(root, saved['limit'])
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
    
    if filename:
        exl_var = tk.StringVar(root, filename)
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

    if filename:
        cfg_var = tk.StringVar(root, filename)
        cg_entry.delete(0,tk.END)
        cg_entry.insert(0,filename)
        save(save_file, 'config', cfg_var.get())


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

    if filename:
        out_var = tk.StringVar(root, filename)
        op_entry.delete(0,tk.END)
        op_entry.insert(0,filename)
        save(save_file, 'output', out_var.get())


#function to run new thread to generate
def generate():
    gn.focus_set()
    generate = threading.Thread(target=generate_thread)
    generate.start()


#call generate function from FillNPrint
def generate_thread():
    fnp_inst = FillNPrint(cfg_var.get(), exl_var.get())

    #exception handling
    if not exl_var.get().endswith('.xlsx') or not os.path.exists(exl_var.get()):
        print("Invalid excel file")
        pt.config(text="Invalid excel file")
        return
    if not sht_var.get() in sh_combobox['values']:
        print("Selected sheet is not a valid sheet")
        pt.config(text="Selected sheet is not a valid sheet")
        return
    if lmt_var.get().upper().isupper() and len(lmt_var.get()) != 0:
        print("'Limit' setting must be an integer or left empty")
        pt.config(text="'Limit' setting must be an integer or left empty")
        return
    if fnp_inst.cfg == "error: invalid yaml file" or not os.path.exists(cfg_var.get()):
        print("Invalid yaml file")
        pt.config(text="Invalid yaml file")
        return
    if "config error:" in fnp_inst.cfg:
        error = fnp_inst.cfg.split("\n")
        print(error[0])
        pt.config(text=error[0])
        return
    if not out_var.get().endswith('.pdf'):
        print("Output file must be a pdf file")
        pt.config(text="Output file must be a pdf file")
        return

    fnp_inst.assign_progress(pb, pt)
    com = "fnp_inst.generate('{}'".format(out_var.get())
    if sht_var.get() != '':
        com = com + ", sheet='{}'".format(sht_var.get())
    if cel_var.get() != '':
        com = com + ", start='{}'".format(cel_var.get())
    if lmt_var.get() != '':
        com = com + ", limit={}".format(abs(int(lmt_var.get())))
    exec(com+')')


#excel file selected
def excel_file(a):
    sheets = FillNPrint(None, exl_var.get()).get_sheets()
    sh_combobox['values'] = [''] + sheets
    sh_combobox.set(sh_combobox['values'][min(len(sh_combobox['values']), int(saved['sheet']))])
    save(save_file, 'excel', exl_var.get())

#label
title = ttk.Label(root, text="FillNPrint")
title.config(font=("TkDefaultFont", 32))
title.pack(expand=True, fill='x', padx=10, pady=10)


#excel file section
ef = ttk.LabelFrame(root, text='Excel File') #excel file frame
ef.pack(expand=True, fill='x', padx=10, pady=10)
ef.grid_columnconfigure(0, weight=1)

ef_entry = ttk.Entry(ef, textvariable=exl_var,takefocus=False) #excel file entry input
ef_entry.grid(column=0, row=0, padx=10, pady=10, sticky='ew')
ef_entry.bind("<FocusOut>", excel_file)
ef_entry.bind('<Control-a>', lambda _: ef_entry.selection_range(0, 'end') or "break")

ef_browse = ttk.Button(ef, text='Browse', command=select_excel_file, takefocus=False) #excel file browse button
ef_browse.grid(column=1, row=0, padx=10, pady=10)


#options section
st = ttk.Frame(root)
st.pack(expand=True, fill='x', padx=10)


#sheet name
sh = ttk.Frame(st)#box size frame
sh.grid(column=0, row=1,padx=10, pady=5, sticky='w')
sh.grid_columnconfigure(0, weight=1)

sh_combobox = ttk.Combobox(sh, textvariable=sht_var, width=8, state='readonly') #box size spinbox
sh_combobox['values'] = sheets
sh_combobox.set(sheets[0])
sh_combobox.grid(column=1, row=0)
sh_combobox.bind("<<ComboboxSelected>>", lambda _: save(save_file, 'sheet', sh_combobox['values'].index(sht_var.get())))


sh_text = ttk.Label(sh, text='Sheet') #box size label
sh_text.grid(column=0, row=0)


#starting cell
sc = ttk.Frame(st)#starting cell frame
sc.grid(column=1, row=1,padx=10, pady=5, sticky='w')
sc.grid_columnconfigure(0, weight=1)

sc_entry = ttk.Entry(sc, textvariable=cel_var, width=5, takefocus=False) #starting cell spinbox
sc_entry.bind("<FocusOut>", lambda event: save(save_file, 'cell', cel_var.get()))
sc_entry.grid(column=1, row=0)
sc_entry.bind('<Control-a>', lambda _: sc_entry.selection_range(0, 'end') or "break")

sc_text = ttk.Label(sc, text='Starting Cell') #starting cell label
sc_text.grid(column=0, row=0)


#limit
lm = ttk.Frame(st)#box size frame
lm.grid(column=2, row=1,padx=10, pady=5, sticky='w')
lm.grid_columnconfigure(0, weight=1)

lm_spinbox = ttk.Spinbox(lm, textvariable=lmt_var, from_=0, to=1000, width=3, takefocus=False) #box size spinbox
lm_spinbox.bind("<FocusOut>", lambda event: save(save_file, 'limit', lmt_var.get()))
lm_spinbox.grid(column=1, row=0)
lm_spinbox.bind('<Control-a>', lambda _: lm_spinbox.selection_range(0, 'end') or "break")

lm_text = ttk.Label(lm, text='Limit') #box size label
lm_text.grid(column=0, row=0)


#config file section
cg = ttk.LabelFrame(root, text='Configuration File') #excel file frame
cg.pack(expand=True, fill='x', padx=10, pady=10)
cg.grid_columnconfigure(0, weight=1)

cg_entry = ttk.Entry(cg, textvariable=cfg_var, takefocus=False) #excel file entry input
cg_entry.grid(column=0, row=0, padx=10, pady=10, sticky='ew')
cg_entry.bind("<FocusOut>", lambda event: save(save_file, 'config', cfg_var.get()))
cg_entry.bind('<Control-a>', lambda _: cg_entry.selection_range(0, 'end') or "break")

cg_browse = ttk.Button(cg, text='Browse', command=select_yaml_file, takefocus=False) #excel file browse button
cg_browse.grid(column=1, row=0, padx=10, pady=10)


#excel file section
op = ttk.LabelFrame(root, text='Output File') #excel file frame
op.pack(expand=True, fill='x', padx=10, pady=10)
op.grid_columnconfigure(0, weight=1)

op_entry = ttk.Entry(op, textvariable=out_var, takefocus=False) #excel file entry input
op_entry.grid(column=0, row=0, padx=10, pady=10, sticky='ew')
op_entry.bind("<FocusOut>", lambda event: save(save_file, 'output', out_var.get()))
op_entry.bind('<Control-a>', lambda _: op_entry.selection_range(0, 'end') or "break")

op_browse = ttk.Button(op, text='Browse', command=select_output, takefocus=False) #excel file browse button
op_browse.grid(column=1, row=0, padx=10, pady=10)


#generate sectopn
gs = ttk.Frame(root) #generate section frame
gs.pack(expand=True, fill='x',side='bottom', anchor='s')

#generate button
gn = ttk.Button(gs,
                text='Generate',
                command=generate,
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
pb.grid(column=0, row=1, padx=20, sticky='ew')


#load sheet names
excel_file('')


root.mainloop()