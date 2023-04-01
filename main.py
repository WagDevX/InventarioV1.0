import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import pandas as pd
import customtkinter


class InVent:
    def __init__(self):
        self.main_window = customtkinter.CTk()
        customtkinter.set_appearance_mode('System')
        self.main_window.title('WAGNER SOLUÇÕES - INVENTÁRIO v1.0')
        self.entry_refs = {}
        self.adicionar = tk.StringVar()

        self.up_button = customtkinter.CTkButton(self.main_window, text='CIMA',
                                                 border_width=1, border_color='lightblue')
        self.up_button.bind("<Button-1>", self.search_up)
        self.up_button.grid(row=2, column=3)

        self.down_button = customtkinter.CTkButton(self.main_window, text='BAIXO',
                                                   border_width=1, border_color='lightblue')
        self.down_button.bind("<Button-1>", self.search_down)
        self.down_button.grid(row=3, column=3, sticky='n')

        self.search_entry = customtkinter.CTkEntry(self.main_window, width=165)
        self.search_entry.grid(row=1, column=2)
        self.search_entry.bind("<Return>", self.search_treeview)
        self.search_button = customtkinter.CTkButton(self.main_window, text="PROCURAR", border_width=1,
                                                     border_color='lightblue', command=self.search_treeview)
        self.search_button.grid(row=1, column=3)

        self.label = customtkinter.CTkLabel(self.main_window, text='INVENTÁRIO v1.0')
        self.label.grid(row=0, column=0)
        self.label_code = customtkinter.CTkLabel(self.main_window, text='CÓDIGO EAN')
        self.label_code.grid(row=1, column=0)

        self.label_qtd = customtkinter.CTkLabel(self.main_window, text='QUANTIDADE')
        self.label_qtd.grid(row=2, column=0)

        self.cod_ean = customtkinter.CTkEntry(self.main_window, width=250)
        self.cod_ean.grid(row=1, column=1)
        self.cod_ean.bind('<KeyPress>', self.shortcut)

        self.qtd_ean = customtkinter.CTkEntry(self.main_window, width=250)
        self.qtd_ean.grid(row=2, column=1)
        self.qtd_ean.configure(state='disabled')

        self.qtd_ean_2 = tk.Entry(self.main_window, width=50, border=5)

        self.check_box = customtkinter.CTkCheckBox(self.main_window, text='QUANTIDADE MANUAL',
                                                   command=self.check_box_status, variable=self.adicionar,
                                                   onvalue='Sim', offvalue='Não')
        self.check_box.deselect()
        self.check_box.grid(row=2, column=2)

        self.button = customtkinter.CTkButton(self.main_window, text='EXPORTAR',
                                              border_width=1, border_color='lightblue',
                                              command=lambda: self.export_to_excel())
        self.button.grid(row=4, column=0)

        self.button = customtkinter.CTkButton(self.main_window, text='DELETAR', border_width=1,
                                              border_color='lightblue', command=lambda: self.deletar())
        self.button.grid(row=4, column=1, stick='s')

        self.button = customtkinter.CTkButton(self.main_window, text='ADICIONAR',
                                              border_width=1, border_color='lightblue', command=lambda: self.inserir())
        self.button.grid(row=4, column=2, stick='s')

        self.button = customtkinter.CTkButton(self.main_window, text='IMPORTAR',
                                              border_width=1, border_color='lightblue',
                                              command=lambda: self.adicionar_excel())
        self.button.grid(row=4, column=3, stick='s')

        self.tree = ttk.Treeview(self.main_window, columns=('cod', 'qt'), show='headings')
        self.tree.column('cod', minwidth=0, width=200)
        self.tree.column('qt', minwidth=0, width=100)
        self.tree.heading('cod', text='CÓDIGO EAN')
        self.tree.heading('qt', text='QUANTIDADE')
        self.tree.grid(row=3, column=1)

        self.style = ttk.Style()
        self.style.theme_use('default')
        self.style.configure('Treeview',
                             background='silver',
                             foregroundcolor='black',
                             rowheight=25,
                             fieldbackground='silver'
                             )
        self.style.map('Treeview',
                       background=[('selected', 'blue')])

        self.entry_refs['ean'] = self.cod_ean
        self.entry_refs['quantidade'] = self.qtd_ean_2

        self.main_window.mainloop()

    def search_up(self, event=None):
        # Obtenha o texto da entrada de pesquisa
        search_text = self.search_entry.get()

        # Obtenha todos os itens na Treeview
        items = self.tree.get_children()

        # Obtenha o índice do item atualmente selecionado
        current_index = items.index(self.tree.selection()[0])

        # Procure por um item que contenha o texto de pesquisa acima do item atual
        for i in range(current_index - 1, -1, -1):
            item_text = self.tree.item(items[i], "values")[0]
            if item_text.lower().find(search_text) != -1:
                self.tree.selection_remove(self.tree.selection())
                self.tree.selection_add(items[i])
                self.tree.see(items[i])
                break

    def search_down(self, event=None):
        # Obtenha o texto da entrada de pesquisa
        search_text = self.search_entry.get()

        # Obtenha todos os itens na Treeview
        items = self.tree.get_children()

        # Obtenha o índice do item atualmente selecionado
        current_index = items.index(self.tree.selection()[0])

        # Procure por um item que contenha o texto de pesquisa abaixo do item atual
        for i in range(current_index + 1, len(items)):
            item_text = self.tree.item(items[i], "values")[0]
            if item_text.lower().find(search_text) != -1:
                self.tree.selection_remove(self.tree.selection())
                self.tree.selection_add(items[i])
                self.tree.see(items[i])
                break

    def search_treeview(self, event=None):
        # Obtenha o texto da entrada de pesquisa
        search_text = self.search_entry.get()

        # Limpe a seleção atual na Treeview
        self.tree.selection_remove(self.tree.selection())

        # Procure pelo item que corresponde ao texto de pesquisa e selecione-o
        items = self.tree.get_children()
        for item in items:
            item_text = self.tree.item(item, "values")[0]
            if item_text.lower().startswith(search_text.lower()):
                self.tree.selection_add(item)
                self.tree.see(item)
                break

    def disable_entry(self):
        self.qtd_ean.configure(state='disabled')

    def enable_entry(self):
        self.qtd_ean.configure(state='normal')

    def check_box_status(self):
        try:
            if self.adicionar.get() == 'Sim':
                self.enable_entry()
                messagebox.showinfo(title='ATENÇÃO', message='Você precisa adicionar a quantidade manualmente agora!')
            elif self.adicionar.get() == 'Não':
                messagebox.showinfo(title='ATENÇÃO', message='Toda vez que você apertar enter ' 
                                                             'irá adicionar UM na quantidade total!')
                self.disable_entry()
        except():
            messagebox.showinfo(title='ERRO', message='ERRO!')

    def shortcut(self, event):
        try:
            if event.keysym == 'Return':
                try:
                    if self.adicionar.get() == 'Não':
                        self.qtd_ean_2.insert(0, '1')
                        self.inserir_enter()
                    else:
                        self.inserir()
                except():
                    messagebox.showinfo(title='ERRO', message='ERRO DO ENTER!')
                else:
                    return

            self.cod_ean.focus()
        except():
            messagebox.showinfo(title='ERRO', message='Houve um erro ao adicionar esse valor!'
                                                      ' Cheque os dados e tente novamente!')

    def inserir(self):
        try:
            if self.cod_ean.get() == '' or self.qtd_ean.get() == '':
                messagebox.showinfo(title='ERRO', message='Preencha todos os campos!')
                return

            ean = self.cod_ean.get()
            quantidade = int(self.qtd_ean.get())

            for item in self.tree.get_children():
                item_ean = self.tree.item(item, 'values')[0]
                if item_ean == ean:
                    item_quantidade = int(self.tree.item(item, 'values')[1])

                    nova_quantidade = item_quantidade + quantidade

                    self.tree.set(item, 'qt', nova_quantidade)

                    self.entry_refs['ean'].delete(0, 'end')
                    self.entry_refs['quantidade'].delete(0, 'end')
                    return

            self.tree.insert('', 'end', values=(self.cod_ean.get(), self.qtd_ean.get()))
            self.cod_ean.delete(0, 'end')
            self.cod_ean.focus()
        except():
            messagebox.showinfo(title='ERRO', message='Verifique os dados e tente novamente!')

    def inserir_enter(self):
        try:
            if self.cod_ean.get() == '' or self.qtd_ean_2.get() == '':
                return

            ean = self.cod_ean.get()
            quantidade = int(self.qtd_ean_2.get())

            for item in self.tree.get_children():
                item_ean = self.tree.item(item, 'values')[0]
                if item_ean == ean:
                    item_quantidade = int(self.tree.item(item, 'values')[1])

                    nova_quantidade = item_quantidade + quantidade

                    self.tree.set(item, 'qt', nova_quantidade)

                    self.entry_refs['ean'].delete(0, 'end')
                    self.entry_refs['quantidade'].delete(0, 'end')
                    return

            self.tree.insert('', 'end', values=(self.cod_ean.get(), self.qtd_ean_2.get()))
            self.cod_ean.delete(0, 'end')
            self.qtd_ean_2.delete(0, 'end')
            self.cod_ean.focus()
        except():
            messagebox.showinfo(title='ERRO', message='Verifique os dados e tente novamente!')

    def deletar(self):
        try:
            itemselecionado = self.tree.selection()[0]
            self.tree.delete(itemselecionado)
        except():
            messagebox.showinfo(title='ERRO', message='Selecione um elemento a ser deletado!')

    def export_to_excel(self):
        data = []
        for item in self.tree.get_children():
            values = self.tree.item(item, 'values')
            data.append({'CÓDIGO EAN': values[0], 'QUANTIDADE': values[1]})

        df = pd.DataFrame(data)

        file_path = customtkinter.filedialog.asksaveasfilename(defaultextension='.xlsx')
        if file_path:
            # O usuário selecionou um arquivo, continuamos com a exportação
            writer = pd.ExcelWriter(file_path, engine='xlsxwriter')
            df.to_excel(writer, index=False)
            writer.save()

    def read_excel(self, treeview_data):
        df = pd.read_excel(treeview_data)
        for i, row in df.iterrows():
            self.tree.insert('', 'end', values=(row['CÓDIGO EAN'], row['QUANTIDADE']))

    def adicionar_excel(self):
        filename = customtkinter.filedialog.askopenfilename(filetypes=[('Excel', '*.xlsx')])
        if filename:
            self.read_excel(filename)


InVent()
