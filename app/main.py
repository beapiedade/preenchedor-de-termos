# backend imports
from extractor import Extractor
from factory import Factory
# frontend imports
import customtkinter
from tkinter import filedialog, messagebox
import os

class Interface(customtkinter.CTk):
    def __init__(self):
        super().__init__()

        # window
        self.title("Gerador de TEC/D")
        self.geometry("700x550")
        self.grid_columnconfigure(1, weight=1)

        # list selector
        self.excel_label = customtkinter.CTkLabel(self, text="Arquivo Excel (lista):")
        self.excel_label.grid(row=0, column=0, padx=10, pady=10, sticky="w")
        self.excel_entry = customtkinter.CTkEntry(self, placeholder_text="Selecione a lista com os dados")
        self.excel_entry.grid(row=0, column=1, padx=10, pady=10, sticky="ew")
        self.excel_button = customtkinter.CTkButton(self, text="Selecionar", width=100, command=self.select_list)
        self.excel_button.grid(row=0, column=2, padx=10, pady=10)

        # model selector
        self.template_label = customtkinter.CTkLabel(self, text="Modelo Word (termo):")
        self.template_label.grid(row=1, column=0, padx=10, pady=10, sticky="w")
        self.template_entry = customtkinter.CTkEntry(self, placeholder_text="Selecione o modelo do documento")
        self.template_entry.grid(row=1, column=1, padx=10, pady=10, sticky="ew")
        self.template_button = customtkinter.CTkButton(self, text="Selecionar", width=100, command=self.select_template)
        self.template_button.grid(row=1, column=2, padx=10, pady=10)

        # destination folder selector
        self.dest_label = customtkinter.CTkLabel(self, text="Pasta de Destino:")
        self.dest_label.grid(row=2, column=0, padx=10, pady=10, sticky="w")
        self.dest_entry = customtkinter.CTkEntry(self, placeholder_text="Selecione onde salvar os arquivos")
        self.dest_entry.grid(row=2, column=1, padx=10, pady=10, sticky="ew")
        self.dest_button = customtkinter.CTkButton(self, text="Selecionar", width=100, command=self.select_destination)
        self.dest_button.grid(row=2, column=2, padx=10, pady=10)

        # process button
        self.process_button = customtkinter.CTkButton(self, text="Gerar Documentos", height=40, command=self.start)
        self.process_button.grid(row=3, column=0, columnspan=3, padx=10, pady=20, sticky="ew")

        # log area
        self.log_textbox = customtkinter.CTkTextbox(self, state="disabled", height=200)
        self.log_textbox.grid(row=4, column=0, columnspan=3, padx=10, pady=10, sticky="nsew")
        self.grid_rowconfigure(4, weight=1) # Faz a caixa de texto expandir

    def select_list(self):
        filepath = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx *.xls")])
        if filepath:
            self.excel_entry.delete(0, "end")
            self.excel_entry.insert(0, filepath)

    def select_template(self):
        filepath = filedialog.askopenfilename(filetypes=[("Documentos Word", "*.docx")])
        if filepath:
            self.template_entry.delete(0, "end")
            self.template_entry.insert(0, filepath)

    def select_destination(self):
        folderpath = filedialog.askdirectory()
        if folderpath:
            self.dest_entry.delete(0, "end")
            self.dest_entry.insert(0, folderpath)

    def log(self, message):
        self.log_textbox.configure(state="normal")
        self.log_textbox.insert("end", message + "\n")
        self.log_textbox.configure(state="disabled")
        self.update_idletasks()

    def start(self):
        data_path = self.excel_entry.get()
        template_path = self.template_entry.get()
        dest_folder = self.dest_entry.get()

        if not all([data_path, template_path, dest_folder]):
            messagebox.showerror("Erro", "Por favor, selecione todos os arquivos e a pasta de destino.")
            return

        self.log_textbox.configure(state="normal")
        self.log_textbox.delete("1.0", "end")
        self.log_textbox.configure(state="disabled")
        self.log(">>> INICIANDO PROCESSO <<<")
        
        try:
            self.log(f"...Lendo dados do arquivo")
            data = Extractor.from_path(data_path)

            self.log("...verificando e validando dados")
            final_data, data_errors = Extractor.check_data(data)

            if final_data:
                self.log("...iniciando a criação dos documentos")
                for item in final_data:
                    Factory.create_document(dest_folder, template_path, item)
                    self.log(f"Documento gerado para: {item[2]}")
                self.log("\n")

            if data_errors:
                for error in data_errors:
                    self.log(f"Erro encontrado no registro de: {error[2]}")
                self.log("\n")
            
            self.log(">>> PROCESSO FINALIZADO <<<")
            messagebox.showinfo("Sucesso", "Processo finalizado! Verifique o log para detalhes.")

        except Exception as e:
            self.log(f"\n!!! OCORREU UM ERRO INESPERADO !!!")
            self.log(f"Erro: {str(e)}")

if __name__ == "__main__":
    app = Interface()
    app.mainloop()