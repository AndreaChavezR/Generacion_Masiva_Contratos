import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from data_manager import DataManager
from docx import Document
import os
import sys
from threading import Thread
from datetime import datetime
import pandas as pd

class ContractSystem:
    def __init__(self, root):
        self.root = root
        # Configurar rutas base
        if getattr(sys, 'frozen', False):
            self.base_dir = os.path.dirname(sys.executable)
        else:
            self.base_dir = os.getcwd()
            
        self.data_manager = DataManager()
        self.current_contract = None
        self.template_path = ""
        self.output_dir = os.path.join(self.base_dir, "contratos_generados")
        self.template_dir = os.path.join(self.base_dir, "plantillas_word")
        
        # Crear carpetas si no existen
        os.makedirs(self.output_dir, exist_ok=True)
        os.makedirs(self.template_dir, exist_ok=True)
        self.create_interface()
    
    def create_interface(self):
        self.notebook = ttk.Notebook(self.root)
        
        # Pestaña de Registro Completo
        self.tab_registro = ttk.Frame(self.notebook)
        self.create_full_registration_tab()
        
        # Pestaña de Generación
        self.tab_generacion = ttk.Frame(self.notebook)
        self.create_generation_tab()
        
        self.notebook.add(self.tab_registro, text="Registro Completo")
        self.notebook.add(self.tab_generacion, text="Generación")
        self.notebook.pack(expand=True, fill='both')
        
        # Barra de estado
        self.status_bar = ttk.Label(self.root, text="Listo", relief=tk.SUNKEN)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)
    
    def create_full_registration_tab(self):
        main_frame = ttk.Frame(self.tab_registro)
        main_frame.pack(fill='both', expand=True)
        
        # Configuración de Scroll
        canvas = tk.Canvas(main_frame)
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Lista completa de campos según Excel
        fields = [
            ('CONTRATO_ADQUISICION', 'entry'),
            ('NO_CONTRATO', 'entry'),
            ('BIENES', 'text'),
            ('TITULAR_AREA_REQUIRENTE', 'entry'),
            ('TITULAR_AREA', 'entry'),
            ('PROVEEDOR', 'entry'),
            ('NOM_PROVEEDOR', 'entry'),
            ('CARGO_PROVEEDOR', 'entry'),
            ('CARGO_AREA_REQUIRENTE', 'entry'),
            ('FECHA_NOMBRAMIENTO', 'entry'),
            ('FECHA_CELEBRACION', 'text'),
            ('FUNDAMENTO', 'text'),
            ('NO_COMPRA', 'text'),
            ('TIPO_ADQUISICION', 'entry'),
            ('ADQUISICION', 'text'),
            ('NECESIDADES', 'text'),
            ('NO_DENOMINACION', 'entry'),
            ('NO_OFICIO', 'entry'),
            ('NO_ESCRITURA_PUBLICA', 'entry'),
            ('FECHA_PUBLICACION', 'entry'),
            ('TITULAR_NOTARIA', 'entry'),
            ('NO_NOTARIA', 'entry'),
            ('NO_MERCANTIL', 'entry'),
            ('ENTIDAD_FEDERATIVA', 'entry'),
            ('DIA', 'entry'),
            ('OBJETO_SOCIAL', 'text'),
            ('PERSONA_FISICA', 'entry'),
            ('CARACTER_PERSONA_FISICA', 'entry'),
            ('TIPO_DOCUMENTO', 'entry'),
            ('NO_DOCUMENTO', 'entry'),
            ('INSTITUCION', 'entry'),
            ('FOLIO_REGISTRO', 'entry'),
            ('NO_ESCRITURA', 'entry'),
            ('FECHA_PUBLICACION2', 'entry'),
            ('INE_NOTARIO', 'entry'),
            ('RFC', 'entry'),
            ('NO_CONSTANCIA', 'entry'),
            ('FECHA_EXPEDICION', 'entry'),
            ('ANEXOS', 'text'),
            ('CONSTANCIAS', 'text'),
            ('DOMICILIO', 'entry'),
            ('NUMERO', 'entry'),
            ('COLONIA', 'entry'),
            ('ALCALDIA', 'entry'),
            ('CP', 'entry'),
            ('TELEFONOS', 'entry'),
            ('CORREO', 'entry'),
            ('CALLE', 'entry'),
            ('NO_EXT', 'entry'),
            ('DESCRIPCION_ADQUISICION', 'text'),
            ('NO_REQUERIMIENTO', 'entry'),
            ('PARTIDA_PRESUPUESTAL', 'entry'),
            ('MONTO_AUTORIZADO', 'entry'),
            ('CORREO_1', 'entry'),
            ('CORREO_2', 'entry'),
            ('FECHA_ENTREGA', 'entry'),
            ('FECHA_VIGENCIA_ENTREGA', 'entry'),
            ('VIGENCIA_CONTRATO', 'entry'),
            ('NO_PAG', 'entry'),
            ('DIAS', 'entry'),
            ('MES', 'entry'),
            ('DIA_FIRMA', 'entry'),
            ('MES_FIRMA', 'entry'),
            ('DIRECCION_DE', 'entry')
        ]
        
        self.entries = {}
        for idx, (field, ftype) in enumerate(fields):
            row = idx % 25  # 25 campos por columna
            col = idx // 25
            
            lbl_frame = ttk.Frame(scrollable_frame)
            lbl_frame.grid(row=row, column=col*2, padx=5, pady=2, sticky='w')
            
            lbl = ttk.Label(lbl_frame, text=f"{self.format_label(field)}:")
            lbl.pack(side='left')
            
            entry_frame = ttk.Frame(scrollable_frame)
            entry_frame.grid(row=row, column=col*2+1, padx=5, pady=2, sticky='ew')
            
            if ftype == 'entry':
                entry = ttk.Entry(entry_frame, width=25)
                entry.pack(fill='x')
            elif ftype == 'text':
                entry = tk.Text(entry_frame, height=3, width=30)
                entry.pack(fill='x')
            
            self.entries[field] = entry
        
        # Botones
        btn_frame = ttk.Frame(scrollable_frame)
        btn_frame.grid(row=25, column=0, columnspan=4, pady=15)
        
        ttk.Button(btn_frame, text="Guardar Contrato", 
                 command=self.save_data).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="Limpiar Formulario", 
                 command=self.clear_form).pack(side='left', padx=5)

    def create_generation_tab(self):
        frame = ttk.Frame(self.tab_generacion, padding=10)
        frame.pack(fill=tk.BOTH, expand=True)

        # Panel de Control
        control_frame = ttk.Frame(frame)
        control_frame.pack(fill=tk.X, pady=10)
        
        ttk.Button(control_frame, text="Seleccionar Plantilla", 
                 command=self.select_template).pack(side=tk.LEFT)
        ttk.Button(control_frame, text="Generar Todos los Contratos", 
                 command=self.start_bulk_generation).pack(side=tk.LEFT, padx=10)
        ttk.Button(control_frame, text="Abrir Carpeta", 
                 command=self.open_output_dir).pack(side=tk.RIGHT)
        
        # Barra de Progreso
        self.progress = ttk.Progressbar(frame, orient=tk.HORIZONTAL, mode='determinate')
        self.progress.pack(fill=tk.X)

    def format_label(self, text):
        """Formatea los nombres de los campos para mostrar"""
        return text.replace('_', ' ').title()

    def save_data(self):
        data = {}
        for field, entry in self.entries.items():
            if isinstance(entry, tk.Text):
                data[field] = entry.get("1.0", tk.END).strip()
            else:
                data[field] = entry.get().strip()
        
        success, message = self.data_manager.save_record(data)
        if success:
            messagebox.showinfo("Éxito", message)
            self.clear_form()
        else:
            messagebox.showerror("Error", message)

    def clear_form(self):
        for entry in self.entries.values():
            if isinstance(entry, tk.Text):
                entry.delete("1.0", tk.END)
            else:
                entry.delete(0, tk.END)

    def select_template(self):
        """Selección de plantilla desde la carpeta designada"""
        self.template_path = filedialog.askopenfilename(
            initialdir=self.template_dir,
            filetypes=[("Plantillas Word", "*.doc")],
            title="Seleccionar plantilla"
        )
        if self.template_path:
            self.update_status(f"Plantilla seleccionada: {os.path.basename(self.template_path)}")

    def start_bulk_generation(self):
        if not self.template_path:
            messagebox.showwarning("Advertencia", "Seleccione una plantilla primero")
            return
        
        Thread(target=self.generate_all_documents, daemon=True).start()

    def generate_all_documents(self):
        try:
            df = self.data_manager.get_pending_records()
            total = len(df)
            
            if total == 0:
                messagebox.showinfo("Información", "No hay contratos pendientes")
                return
            
            self.progress['maximum'] = total
            os.makedirs(self.output_dir, exist_ok=True)
            
            for idx, (_, row) in enumerate(df.iterrows()):
                doc = Document(self.template_path)
                replacements = row.to_dict()
                replacements['FECHA_GENERACION'] = datetime.now().strftime("%d/%m/%Y %H:%M")
                
                # Reemplazo completo en todo el documento
                for p in doc.paragraphs:
                    for key, value in replacements.items():
                        p.text = p.text.replace(f'{{{{{key}}}}}', str(value))
                
                for table in doc.tables:
                    for row_table in table.rows:
                        for cell in row_table.cells:
                            for key, value in replacements.items():
                                cell.text = cell.text.replace(f'{{{{{key}}}}}', str(value))
                
                filename = f"Contrato_{row['NO_CONTRATO']}.doc"
                output_path = os.path.join(self.output_dir, filename)
                doc.save(output_path)
                
                # Actualizar estado
                self.data_manager.mark_as_generated(row['ID'])
                self.progress['value'] = idx + 1
                self.update_status(f"Generado: {filename}")
            
            messagebox.showinfo("Éxito", f"Se generaron {total} contratos")
            self.progress['value'] = 0
            
        except Exception as e:
            messagebox.showerror("Error", f"Error en generación masiva: {str(e)}")

    def open_output_dir(self):
        try:
            os.startfile(self.output_dir)
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo abrir la carpeta: {str(e)}")

    def update_status(self, message):
        self.status_bar.config(text=message)
        self.root.update_idletasks()

if __name__ == "__main__":
    root = tk.Tk()
    root.geometry("1200x800")
    app = ContractSystem(root)
    root.mainloop()